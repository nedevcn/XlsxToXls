using System.IO.Compression;
using System.Xml;

namespace Nedev.XlsxToXls.Internal;

/// <summary>
/// High-performance XLSX reader using XmlReader (streaming) and ZipArchive.
/// Zero third-party dependencies - uses only System.IO.Compression and System.Xml.
/// </summary>
internal static partial class XlsxReader
{
    private static readonly XmlReaderSettings Sds = new()
    {
        Async = false,
        CloseInput = false,
        IgnoreWhitespace = true,
        DtdProcessing = DtdProcessing.Prohibit,
        ConformanceLevel = ConformanceLevel.Fragment
    };

    public static (List<ReadOnlyMemory<char>> SharedStrings, List<SheetData> Sheets, StylesData? Styles, List<DefinedNameInfo> DefinedNames) Read(Stream xlsxStream)
    {
        using var archive = new ZipArchive(xlsxStream, ZipArchiveMode.Read, leaveOpen: true);
        var sharedStrings = ReadSharedStrings(archive);
        var (sheets, definedNames) = ReadSheetsAndDefinedNames(archive);
        var styles = StylesReader.Read(archive);
        return (sharedStrings, sheets, styles, definedNames);
    }

    private static List<ReadOnlyMemory<char>> ReadSharedStrings(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry == null)
            return [];

        var list = new List<ReadOnlyMemory<char>>();
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Sds);
        var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si" && reader.NamespaceURI == ns)
            {
                var si = ReadSi(reader, ns);
                list.Add((si ?? string.Empty).AsMemory());
            }
        }

        return list;
    }

    private static string? ReadSi(XmlReader reader, string ns)
    {
        if (reader.IsEmptyElement) return string.Empty;
        var sb = new System.Text.StringBuilder();
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "si") break;
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t" && reader.NamespaceURI == ns)
            {
                if (!reader.IsEmptyElement && reader.Read() && reader.NodeType == XmlNodeType.Text)
                    sb.Append(reader.Value);
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "r" && reader.NamespaceURI == ns)
            {
                sb.Append(ReadRichTextRun(reader, ns));
            }
        }
        return sb.ToString();
    }

    private static string ReadRichTextRun(XmlReader reader, string ns)
    {
        if (reader.IsEmptyElement) return string.Empty;
        var sb = new System.Text.StringBuilder();
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "r") break;
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t" && reader.NamespaceURI == ns)
            {
                if (!reader.IsEmptyElement && reader.Read() && reader.NodeType == XmlNodeType.Text)
                    sb.Append(reader.Value);
            }
        }
        return sb.ToString();
    }

    private static (List<SheetData> Sheets, List<DefinedNameInfo> DefinedNames) ReadSheetsAndDefinedNames(ZipArchive archive)
    {
        var rels = ReadWorkbookRels(archive);
        var (sheetIds, definedNames) = ReadSheetIdsAndDefinedNames(archive);
        var list = new List<SheetData>();
        foreach (var (name, rId) in sheetIds)
        {
            if (!rels.TryGetValue(rId, out var path)) continue;
            var (rows, colInfos, mergeRanges, freezePane, rowBreaks, colBreaks, pageSetup, pageMargins, hyperlinks, dataValidations, conditionalFormats) = ReadWorksheet(archive, path);
            var comments = ReadSheetComments(archive, path);
            list.Add(new SheetData(name, rows, colInfos, mergeRanges, freezePane, rowBreaks, colBreaks, pageSetup, pageMargins, hyperlinks, comments, dataValidations, conditionalFormats));
        }
        return (list, definedNames);
    }

    private static Dictionary<string, string> ReadWorkbookRels(ZipArchive archive)
    {
        var dict = new Dictionary<string, string>(StringComparer.Ordinal);
        var entry = archive.GetEntry("xl/_rels/workbook.xml.rels");
        if (entry == null) return dict;

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Sds);
        var ns = "http://schemas.openxmlformats.org/package/2006/relationships";
        string? id = null;
        string? target = null;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.LocalName == "Relationship" && reader.NamespaceURI == ns)
                {
                    id = reader.GetAttribute("Id");
                    target = reader.GetAttribute("Target");
                    if (id != null && target != null)
                    {
                        var path = target.StartsWith("/") ? target.TrimStart('/') : "xl/" + target;
                        dict[id] = path;
                    }
                }
            }
        }
        return dict;
    }

    private static (List<(string Name, string RId)> SheetIds, List<DefinedNameInfo> DefinedNames) ReadSheetIdsAndDefinedNames(ZipArchive archive)
    {
        var sheetIds = new List<(string, string)>();
        var definedNames = new List<DefinedNameInfo>();
        var entry = archive.GetEntry("xl/workbook.xml");
        if (entry == null) return (sheetIds, definedNames);

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Sds);
        var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheet" && reader.NamespaceURI == ns)
            {
                var name = reader.GetAttribute("name") ?? "Sheet";
                var rId = reader.GetAttribute("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id") ?? reader.GetAttribute("id");
                if (rId != null)
                    sheetIds.Add((name, rId));
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "definedName" && reader.NamespaceURI == ns)
            {
                var nameAttr = reader.GetAttribute("name");
                var localSheetId = ParseInt(reader.GetAttribute("localSheetId"));
                if (localSheetId < 0) localSheetId = 0;
                var formula = reader.ReadElementContentAsString();
                if (string.IsNullOrEmpty(nameAttr)) continue;
                var name = nameAttr.Trim();
                byte builtin = name switch
                {
                    "Print_Area" or "_xlnm.Print_Area" => 6,
                    "Print_Titles" or "_xlnm.Print_Titles" => 7,
                    _ => 0
                };
                if (builtin == 0) continue;
                if (!TryParseDefinedNameRange(formula, out var firstRow, out var firstCol, out var lastRow, out var lastCol))
                    continue;
                definedNames.Add(new DefinedNameInfo(localSheetId, builtin, firstRow, firstCol, lastRow, lastCol));
            }
        }
        return (sheetIds, definedNames);
    }

    private static bool TryParseDefinedNameRange(string formula, out int firstRow, out int firstCol, out int lastRow, out int lastCol)
    {
        firstRow = firstCol = 0;
        lastRow = lastCol = -1;
        var refPart = formula;
        var excl = formula.IndexOf('!');
        if (excl >= 0 && excl + 1 < formula.Length)
            refPart = formula[(excl + 1)..].Trim();
        refPart = refPart.Replace("$", "");
        if (string.IsNullOrEmpty(refPart)) return false;
        var colon = refPart.IndexOf(':');
        if (colon < 0) return false;
        var left = refPart[..colon].Trim();
        var right = refPart[(colon + 1)..].Trim();
        if (string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right)) return false;
        if (left.Length <= 5 && right.Length <= 5 && int.TryParse(left, out var row1) && int.TryParse(right, out var row2))
        {
            firstRow = Math.Max(0, row1 - 1);
            lastRow = Math.Min(65535, row2 - 1);
            firstCol = 0;
            lastCol = 255;
            return true;
        }
        if (left.Length >= 1 && char.IsLetter(left[0]) && right.Length >= 1 && char.IsLetter(right[0]) &&
            !int.TryParse(left, out _) && !int.TryParse(right, out _))
        {
            firstCol = ParseColRef(left.AsSpan());
            lastCol = ParseColRef(right.AsSpan());
            firstRow = 0;
            lastRow = 65535;
            return firstCol >= 0 && lastCol >= 0 && firstCol <= 255 && lastCol <= 255;
        }
        var (r1, c1) = ParseCellRef(left);
        var (r2, c2) = ParseCellRef(right);
        firstRow = Math.Min(r1, r2);
        lastRow = Math.Max(r1, r2);
        firstCol = Math.Min(c1, c2);
        lastCol = Math.Max(c1, c2);
        return firstRow <= lastRow && firstCol <= lastCol;
    }

    private static (List<RowData> Rows, List<ColInfo> ColInfos, List<MergeRange> MergeRanges, FreezePaneInfo? FreezePane, List<int> RowBreaks, List<int> ColBreaks, PageSetupInfo? PageSetup, PageMarginsInfo? PageMargins, List<HyperlinkInfo> Hyperlinks, List<DataValidationInfo> DataValidations, List<ConditionalFormatInfo> ConditionalFormats) ReadWorksheet(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path) ?? archive.GetEntry("xl/" + path);
        if (entry == null) return (new List<RowData>(), new List<ColInfo>(), new List<MergeRange>(), null, new List<int>(), new List<int>(), null, null, new List<HyperlinkInfo>(), new List<DataValidationInfo>(), new List<ConditionalFormatInfo>());

        var rows = new List<RowData>();
        var colInfos = new List<ColInfo>();
        var mergeRanges = new List<MergeRange>();
        var rowBreaks = new List<int>();
        var colBreaks = new List<int>();
        var hyperlinks = new List<HyperlinkInfo>();
        var dataValidations = new List<DataValidationInfo>();
        var conditionalFormats = new List<ConditionalFormatInfo>();
        var sharedFormulas = new Dictionary<int, (int BaseRow, int BaseCol, string Formula)>();
        FreezePaneInfo? freezePane = null;
        PageSetupInfo? pageSetup = null;
        PageMarginsInfo? pageMargins = null;
        var inRowBreaks = false;
        var inColBreaks = false;

        var hyperlinkRels = ReadWorksheetRels(archive, path);

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Sds);
        var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        RowData? currentRow = null;
        var cells = new List<CellData>();

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "pane" && reader.NamespaceURI == ns)
            {
                var xSplit = ParseInt(reader.GetAttribute("xSplit"));
                var ySplit = ParseInt(reader.GetAttribute("ySplit"));
                var topLeft = reader.GetAttribute("topLeftCell");
                var (topRow, leftCol) = !string.IsNullOrEmpty(topLeft) ? ParseCellRef(topLeft) : (0, 0);
                if (xSplit > 0 || ySplit > 0)
                    freezePane = new FreezePaneInfo(Math.Max(0, xSplit), Math.Max(0, ySplit), topRow, leftCol);
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "rowBreaks" && reader.NamespaceURI == ns)
                inRowBreaks = true;
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "colBreaks" && reader.NamespaceURI == ns)
                inColBreaks = true;
            else if (reader.NodeType == XmlNodeType.EndElement && (reader.LocalName == "rowBreaks" || reader.LocalName == "colBreaks") && reader.NamespaceURI == ns)
            {
                inRowBreaks = false;
                inColBreaks = false;
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "brk" && reader.NamespaceURI == ns)
            {
                var id = ParseInt(reader.GetAttribute("id"));
                if (id >= 0 && inRowBreaks) rowBreaks.Add(id);
                if (id >= 0 && inColBreaks) colBreaks.Add(id);
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "pageMargins" && reader.NamespaceURI == ns)
            {
                var left = ParseDouble(reader.GetAttribute("left"));
                var right = ParseDouble(reader.GetAttribute("right"));
                var top = ParseDouble(reader.GetAttribute("top"));
                var bottom = ParseDouble(reader.GetAttribute("bottom"));
                var header = ParseDouble(reader.GetAttribute("header"));
                var footer = ParseDouble(reader.GetAttribute("footer"));
                pageMargins = new PageMarginsInfo(left >= 0 ? left : 0.7, right >= 0 ? right : 0.7, top >= 0 ? top : 0.75, bottom >= 0 ? bottom : 0.75, header >= 0 ? header : 0.3, footer >= 0 ? footer : 0.3);
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "pageSetup" && reader.NamespaceURI == ns)
            {
                var orientation = reader.GetAttribute("orientation");
                var scale = ParseInt(reader.GetAttribute("scale"));
                var fitToWidth = ParseInt(reader.GetAttribute("fitToWidth"));
                var fitToHeight = ParseInt(reader.GetAttribute("fitToHeight"));
                var startPage = ParseInt(reader.GetAttribute("useFirstPageNumber")) == 1 ? ParseInt(reader.GetAttribute("firstPageNumber")) : 1;
                pageSetup = new PageSetupInfo(orientation == "landscape", scale > 0 ? scale : 100, fitToWidth >= 0 ? fitToWidth : 0, fitToHeight >= 0 ? fitToHeight : 0, startPage > 0 ? startPage : 1);
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "col" && reader.NamespaceURI == ns)
            {
                var min = ParseInt(reader.GetAttribute("min")) - 1;
                var max = ParseInt(reader.GetAttribute("max")) - 1;
                var w = ParseDouble(reader.GetAttribute("width"));
                var customWidth = reader.GetAttribute("customWidth") == "1";
                var hidden = reader.GetAttribute("hidden") == "1";
                if (min >= 0 && max >= min && (customWidth || hidden))
                    colInfos.Add(new ColInfo(min, max, w > 0 ? w : 10, hidden));
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "dataValidation" && (reader.NamespaceURI == ns || reader.NamespaceURI == "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"))
            {
                var dv = ReadDataValidation(reader, ns);
                if (dv.HasValue && dv.Value.Ranges.Count > 0)
                    dataValidations.Add(dv.Value);
            }
            // Conditional formatting is documented as not supported; we intentionally skip it.
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "conditionalFormatting" && reader.NamespaceURI == ns)
            {
                if (!reader.IsEmptyElement)
                {
                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "conditionalFormatting" && reader.NamespaceURI == ns)
                            break;
                    }
                }
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "hyperlink" && reader.NamespaceURI == ns)
            {
                var refAttr = reader.GetAttribute("ref");
                var rId = reader.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                    ?? reader.GetAttribute("r:id");
                if (!string.IsNullOrEmpty(refAttr) && !string.IsNullOrEmpty(rId) && hyperlinkRels.TryGetValue(rId, out var url) && !string.IsNullOrEmpty(url))
                {
                    var range = ParseRefToRange(refAttr);
                    if (range.HasValue)
                        hyperlinks.Add(new HyperlinkInfo(range.Value.FirstRow, range.Value.FirstCol, range.Value.LastRow, range.Value.LastCol, url));
                }
            }
            // <mergeCells>/<mergeCell>: merged ranges
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "mergeCells" && reader.NamespaceURI == ns)
            {
                if (reader.IsEmptyElement) continue;
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "mergeCells" && reader.NamespaceURI == ns)
                        break;
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "mergeCell" && reader.NamespaceURI == ns)
                    {
                        var refs = reader.GetAttribute("ref");
                        if (!string.IsNullOrEmpty(refs) && ParseMergeRef(refs) is { } m)
                            mergeRanges.Add(new MergeRange(m.FirstRow, m.FirstCol, m.LastRow, m.LastCol));
                    }
                }
            }
            // <row>: physical row (index, height, hidden)
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row" && reader.NamespaceURI == ns)
            {
                if (currentRow.HasValue)
                {
                    currentRow = currentRow.Value with { Cells = cells.ToArray() };
                    rows.Add(currentRow.Value);
                    cells.Clear();
                }

                var rAttr = reader.GetAttribute("r");
                var rowIndex = ParseRowIndex(rAttr);
                var ht = ParseDouble(reader.GetAttribute("ht"));
                var hidden = reader.GetAttribute("hidden") == "1";
                currentRow = new RowData(rowIndex, [], ht, hidden);
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c" && reader.NamespaceURI == ns && currentRow.HasValue)
            {
                var cell = ReadCell(reader, ns, sharedFormulas);
                if (cell.HasValue)
                    cells.Add(cell.Value);
            }
        }

        if (currentRow.HasValue)
        {
            currentRow = currentRow.Value with { Cells = cells.ToArray() };
            rows.Add(currentRow.Value);
        }
        return (rows, colInfos, mergeRanges, freezePane, rowBreaks, colBreaks, pageSetup, pageMargins, hyperlinks, dataValidations, conditionalFormats);
    }

    private static DataValidationInfo? ReadDataValidation(XmlReader reader, string ns)
    {
        var sqref = reader.GetAttribute("sqref");
        var typeStr = reader.GetAttribute("type") ?? "none";
        var opStr = reader.GetAttribute("operator") ?? "between";
        var allowBlank = reader.GetAttribute("allowBlank") == "1";
        var showInputMessage = reader.GetAttribute("showInputMessage") != "0";
        var showErrorMessage = reader.GetAttribute("showErrorMessage") != "0";
        string? formula1 = null, formula2 = null, promptTitle = null, errorTitle = null, promptText = null, errorText = null;
        if (reader.IsEmptyElement) { }
        else
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "dataValidation") break;
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.LocalName == "formula1" && reader.Read() && reader.NodeType == XmlNodeType.Text) formula1 = reader.Value;
                else if (reader.LocalName == "formula2" && reader.Read() && reader.NodeType == XmlNodeType.Text) formula2 = reader.Value;
                else if (reader.LocalName == "promptTitle" && reader.Read() && reader.NodeType == XmlNodeType.Text) promptTitle = reader.Value;
                else if (reader.LocalName == "errorTitle" && reader.Read() && reader.NodeType == XmlNodeType.Text) errorTitle = reader.Value;
                else if (reader.LocalName == "prompt" && reader.Read() && reader.NodeType == XmlNodeType.Text) promptText = reader.Value;
                else if (reader.LocalName == "error" && reader.Read() && reader.NodeType == XmlNodeType.Text) errorText = reader.Value;
            }
        var ranges = new List<(int, int, int, int)>();
        if (!string.IsNullOrEmpty(sqref))
            foreach (var refPart in sqref.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                if (ParseRefToRange(refPart) is (int fr, int fc, int lr, int lc))
                    ranges.Add((fr, fc, lr, lc));
        if (ranges.Count == 0) return null;
        var type = typeStr.ToLowerInvariant() switch { "list" => 3, "whole" => 1, "decimal" => 2, "date" => 4, "time" => 5, "textLength" => 6, "custom" => 7, _ => 0 };
        var op = opStr.ToLowerInvariant() switch { "notBetween" => 1, "equal" => 2, "notEqual" => 3, "greaterThan" => 4, "lessThan" => 5, "greaterThanOrEqual" => 6, "lessThanOrEqual" => 7, _ => 0 };
        return new DataValidationInfo(ranges, type, op, formula1 ?? "", formula2 ?? "", allowBlank, showInputMessage, showErrorMessage, promptTitle ?? "", errorTitle ?? "", promptText ?? "", errorText ?? "");
    }

    private static Dictionary<string, string> ReadWorksheetRels(ZipArchive archive, string worksheetPath)
    {
        var dict = new Dictionary<string, string>(StringComparer.Ordinal);
        var lastSlash = worksheetPath.LastIndexOf('/');
        var relsPath = lastSlash >= 0
            ? worksheetPath[..(lastSlash + 1)] + "_rels/" + worksheetPath[(lastSlash + 1)..] + ".rels"
            : "_rels/" + worksheetPath + ".rels";
        var entry = archive.GetEntry(relsPath) ?? archive.GetEntry("xl/worksheets/_rels/" + (lastSlash >= 0 ? worksheetPath[(lastSlash + 1)..] : worksheetPath) + ".rels");
        if (entry == null) return dict;

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Sds);
        var ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship" && reader.NamespaceURI == ns)
            {
                var type = reader.GetAttribute("Type");
                var id = reader.GetAttribute("Id");
                var target = reader.GetAttribute("Target");
                if (id != null && target != null && type != null && type.EndsWith("hyperlink", StringComparison.OrdinalIgnoreCase))
                    dict[id] = target;
            }
        }
        return dict;
    }

    private static void ReadConditionalFormatting(XmlReader reader, string ns, List<ConditionalFormatInfo> list)
    {
        var sqref = reader.GetAttribute("sqref");
        var ranges = new List<(int FirstRow, int FirstCol, int LastRow, int LastCol)>();
        if (!string.IsNullOrEmpty(sqref))
        {
            foreach (var refPart in sqref.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                if (ParseRefToRange(refPart) is (int fr, int fc, int lr, int lc))
                    ranges.Add((fr, fc, lr, lc));
        }
        if (reader.IsEmptyElement) return;
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "conditionalFormatting" && reader.NamespaceURI == ns)
                break;
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "cfRule" || reader.NamespaceURI != ns)
                continue;

            var typeStr = reader.GetAttribute("type") ?? "cellIs";
            var opStr = reader.GetAttribute("operator") ?? "between";
            var ct = string.Equals(typeStr, "expression", StringComparison.OrdinalIgnoreCase) ? 2 : 1;
            var cp = opStr.ToLowerInvariant() switch
            {
                "between" => 1,
                "notBetween" => 2,
                "equal" => 3,
                "notEqual" => 4,
                "greaterThan" => 5,
                "lessThan" => 6,
                "greaterThanOrEqual" => 7,
                "lessThanOrEqual" => 8,
                _ => 0
            };
            string? f1 = null, f2 = null;
            if (!reader.IsEmptyElement)
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "cfRule" && reader.NamespaceURI == ns)
                        break;
                    if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != ns) continue;
                    if (reader.LocalName == "formula" && reader.Read() && reader.NodeType == XmlNodeType.Text)
                    {
                        if (f1 == null) f1 = reader.Value;
                        else if (f2 == null) f2 = reader.Value;
                    }
                }
            }

            if (ranges.Count > 0)
                list.Add(new ConditionalFormatInfo(new List<(int FirstRow, int FirstCol, int LastRow, int LastCol)>(ranges), ct, cp, f1 ?? string.Empty, f2 ?? string.Empty));
        }
    }

    private static List<CommentInfo> ReadSheetComments(ZipArchive archive, string worksheetPath)
    {
        var list = new List<CommentInfo>();
        var commentsPath = GetWorksheetCommentsPath(archive, worksheetPath);
        if (string.IsNullOrEmpty(commentsPath)) return list;

        var entry = archive.GetEntry(commentsPath) ?? archive.GetEntry("xl/" + commentsPath);
        if (entry == null) return list;

        var authors = new List<string>();
        using (var stream = entry.Open())
        using (var reader = XmlReader.Create(stream, Sds))
        {
            var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "author" && reader.NamespaceURI == ns)
                {
                    if (!reader.IsEmptyElement && reader.Read() && reader.NodeType == XmlNodeType.Text)
                        authors.Add(reader.Value ?? "");
                }
                else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "comment" && reader.NamespaceURI == ns)
                {
                    var refAttr = reader.GetAttribute("ref");
                    var authorId = ParseInt(reader.GetAttribute("authorId"));
                    var text = ReadCommentText(reader, ns);
                    if (!string.IsNullOrEmpty(refAttr) && ParseRefToRange(refAttr) is (int r, int c, _, _))
                    {
                        var author = authorId >= 0 && authorId < authors.Count ? authors[authorId] : "";
                        list.Add(new CommentInfo(r, c, author, text ?? ""));
                    }
                }
            }
        }
        return list;
    }

    private static string ReadCommentText(XmlReader reader, string ns)
    {
        if (reader.IsEmptyElement) return "";
        var sb = new System.Text.StringBuilder();
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "comment") break;
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t" && reader.NamespaceURI == ns)
            {
                if (!reader.IsEmptyElement && reader.Read() && reader.NodeType == XmlNodeType.Text)
                    sb.Append(reader.Value);
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "r" && reader.NamespaceURI == ns)
            {
                sb.Append(ReadCommentRunText(reader, ns));
            }
        }
        return sb.ToString();
    }

    private static string ReadCommentRunText(XmlReader reader, string ns)
    {
        if (reader.IsEmptyElement) return "";
        var sb = new System.Text.StringBuilder();
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "r") break;
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t" && reader.NamespaceURI == ns)
            {
                if (!reader.IsEmptyElement && reader.Read() && reader.NodeType == XmlNodeType.Text)
                    sb.Append(reader.Value);
            }
        }
        return sb.ToString();
    }

    private static string? GetWorksheetCommentsPath(ZipArchive archive, string worksheetPath)
    {
        var lastSlash = worksheetPath.LastIndexOf('/');
        var relsPath = lastSlash >= 0
            ? worksheetPath[..(lastSlash + 1)] + "_rels/" + worksheetPath[(lastSlash + 1)..] + ".rels"
            : "_rels/" + worksheetPath + ".rels";
        var entry = archive.GetEntry(relsPath) ?? archive.GetEntry("xl/worksheets/_rels/" + (lastSlash >= 0 ? worksheetPath[(lastSlash + 1)..] : worksheetPath) + ".rels");
        if (entry == null) return null;

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Sds);
        var ns = "http://schemas.openxmlformats.org/package/2006/relationships";
        string? commentsTarget = null;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship" && reader.NamespaceURI == ns)
            {
                var type = reader.GetAttribute("Type");
                var target = reader.GetAttribute("Target");
                if (target != null && type != null && type.EndsWith("comments", StringComparison.OrdinalIgnoreCase))
                {
                    commentsTarget = target;
                    break;
                }
            }
        }

        if (string.IsNullOrEmpty(commentsTarget)) return null;
        if (commentsTarget.StartsWith("/")) return commentsTarget.TrimStart('/');
        var relsDir = relsPath.Contains('/') ? relsPath[..(relsPath.LastIndexOf('/') + 1)] : "";
        var segments = new List<string>(relsDir.Split('/', StringSplitOptions.RemoveEmptyEntries));
        foreach (var seg in commentsTarget.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (seg == ".." && segments.Count > 0) segments.RemoveAt(segments.Count - 1);
            else if (seg != ".") segments.Add(seg);
        }
        return segments.Count > 0 ? string.Join("/", segments) : null;
    }

    private static (int FirstRow, int FirstCol, int LastRow, int LastCol)? ParseRefToRange(string refAttr)
    {
        var idx = refAttr.IndexOf(':');
        if (idx > 0)
            return ParseMergeRef(refAttr);
        var (r, c) = ParseCellRef(refAttr);
        return (r, c, r, c);
    }

    private static (int FirstRow, int FirstCol, int LastRow, int LastCol)? ParseMergeRef(string refs)
    {
        var idx = refs.IndexOf(':');
        if (idx <= 0) return null;
        var (r1, c1) = ParseCellRef(refs[..idx]);
        var (r2, c2) = ParseCellRef(refs[(idx + 1)..]);
        return (Math.Min(r1, r2), Math.Min(c1, c2), Math.Max(r1, r2), Math.Max(c1, c2));
    }

    private static double ParseDouble(string? s)
    {
        if (string.IsNullOrEmpty(s) || !double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v)) return -1;
        return v;
    }

    private static CellData? ReadCell(XmlReader reader, string ns, Dictionary<int, (int BaseRow, int BaseCol, string Formula)> sharedFormulas)
    {
        var r = reader.GetAttribute("r");
        var t = reader.GetAttribute("t");
        var s = reader.GetAttribute("s");
        if (string.IsNullOrEmpty(r)) return null;

        var (row, col) = ParseCellRef(r);
        var styleIndex = ParseInt(reader.GetAttribute("s"));
        if (reader.IsEmptyElement) return new CellData(row, col, CellKind.Empty, CellKind.Empty, null, -1, styleIndex, null);

        string? value = null;
        string? formula = null;
        string? fType = null;
        var fShared = -1;
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "c") break;
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "v" && reader.NamespaceURI == ns)
            {
                if (reader.Read() && reader.NodeType == XmlNodeType.Text)
                    value = reader.Value;
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "f" && reader.NamespaceURI == ns)
            {
                fType = reader.GetAttribute("t");
                fShared = ParseInt(reader.GetAttribute("si"));
                var text = reader.IsEmptyElement ? "" : reader.ReadElementContentAsString();
                if (!string.IsNullOrEmpty(text))
                    formula = text;
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "is" && reader.NamespaceURI == ns)
            {
                value = ReadInlineStr(reader, ns);
            }
        }

        var cachedKind = CellKind.Empty;
        var sstIndex = -1;
        if (value != null)
        {
            if (t == "s" && int.TryParse(value, out var idx))
            {
                cachedKind = CellKind.SharedString;
                sstIndex = idx;
                value = null;
            }
            else if (t == "str" || t == "inlineStr")
                cachedKind = CellKind.String;
            else if (t == "b" && (value == "1" || value == "0"))
                cachedKind = CellKind.Boolean;
            else if (t == "e")
                cachedKind = CellKind.Error;
            else if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _))
                cachedKind = CellKind.Number;
            else
                cachedKind = CellKind.String;
        }

        if (fType == "shared" && fShared >= 0)
        {
            if (formula != null)
                sharedFormulas[fShared] = (row, col, formula);
            else if (sharedFormulas.TryGetValue(fShared, out var def))
                formula = ExpandSharedFormula(def.Formula, def.BaseRow, def.BaseCol, row, col);
        }

        if (formula != null)
            return new CellData(row, col, CellKind.Formula, cachedKind, value, sstIndex, styleIndex, formula);

        if (value == null && sstIndex < 0)
            return new CellData(row, col, CellKind.Empty, CellKind.Empty, null, -1, styleIndex, null);
        if (sstIndex >= 0)
            return new CellData(row, col, CellKind.SharedString, CellKind.SharedString, null, sstIndex, styleIndex, null);
        if (cachedKind == CellKind.Boolean)
            return new CellData(row, col, CellKind.Boolean, CellKind.Boolean, value, -1, styleIndex, null);
        if (cachedKind == CellKind.Error)
            return new CellData(row, col, CellKind.Error, CellKind.Error, value, -1, styleIndex, null);
        if (cachedKind == CellKind.Number)
            return new CellData(row, col, CellKind.Number, CellKind.Number, value, -1, styleIndex, null);
        return new CellData(row, col, CellKind.String, CellKind.String, value, -1, styleIndex, null);
    }

    private static string ExpandSharedFormula(string baseFormula, int baseRow, int baseCol, int targetRow, int targetCol)
    {
        var dr = targetRow - baseRow;
        var dc = targetCol - baseCol;
        if (dr == 0 && dc == 0) return baseFormula;
        return ShiftA1References(baseFormula, dr, dc);
    }

    private static string ShiftA1References(string formula, int dr, int dc)
    {
        if (string.IsNullOrEmpty(formula) || (dr == 0 && dc == 0)) return formula;
        var sb = new System.Text.StringBuilder(formula.Length + 8);
        var inString = false;
        for (var i = 0; i < formula.Length; i++)
        {
            var ch = formula[i];
            if (ch == '"')
            {
                inString = !inString;
                sb.Append(ch);
                continue;
            }
            if (inString)
            {
                sb.Append(ch);
                continue;
            }

            var j = i;
            var absCol = false;
            var absRow = false;
            if (j < formula.Length && formula[j] == '$') { absCol = true; j++; }
            var colStart = j;
            while (j < formula.Length && char.IsLetter(formula[j])) j++;
            var hasCol = j > colStart;
            if (!hasCol) { sb.Append(ch); continue; }
            var colText = formula.AsSpan(colStart, j - colStart);
            if (j < formula.Length && formula[j] == '$') { absRow = true; j++; }
            var rowStart = j;
            while (j < formula.Length && char.IsDigit(formula[j])) j++;
            var hasRow = j > rowStart;
            if (!hasRow) { sb.Append(ch); continue; }

            var prev = i > 0 ? formula[i - 1] : '\0';
            if (char.IsLetterOrDigit(prev) || prev == '_' || prev == '.')
            {
                sb.Append(ch);
                continue;
            }

            if (!int.TryParse(formula.AsSpan(rowStart, j - rowStart), out var row1))
            {
                sb.Append(ch);
                continue;
            }
            var col1 = ParseColRef(colText);
            if (col1 < 0) { sb.Append(ch); continue; }
            var newCol = absCol ? col1 : col1 + dc;
            var newRow = absRow ? (row1 - 1) : (row1 - 1) + dr;
            if (newCol < 0) newCol = 0;
            if (newCol > 255) newCol = 255;
            if (newRow < 0) newRow = 0;
            if (newRow > 65535) newRow = 65535;

            if (absCol) sb.Append('$');
            sb.Append(ColToA1(newCol));
            if (absRow) sb.Append('$');
            sb.Append(newRow + 1);
            i = j - 1;
        }
        return sb.ToString();
    }

    private static string ColToA1(int col)
    {
        col += 1;
        Span<char> tmp = stackalloc char[4];
        var len = 0;
        while (col > 0)
        {
            var rem = (col - 1) % 26;
            tmp[len++] = (char)('A' + rem);
            col = (col - 1) / 26;
        }
        for (var i = 0; i < len / 2; i++)
            (tmp[i], tmp[len - 1 - i]) = (tmp[len - 1 - i], tmp[i]);
        return new string(tmp[..len]);
    }

    private static string ReadInlineStr(XmlReader reader, string ns)
    {
        if (reader.IsEmptyElement) return string.Empty;
        var sb = new System.Text.StringBuilder();
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "is") break;
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t" && reader.NamespaceURI == ns)
            {
                if (!reader.IsEmptyElement && reader.Read() && reader.NodeType == XmlNodeType.Text)
                    sb.Append(reader.Value);
            }
        }
        return sb.ToString();
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
    private static (int Row, int Col) ParseCellRef(string r)
    {
        var i = 0;
        while (i < r.Length && char.IsLetter(r[i])) i++;
        var col = ParseColRef(r.AsSpan(0, i));
        var row = int.Parse(r.AsSpan(i), System.Globalization.NumberStyles.None) - 1;
        return (row, col);
    }

    private static int ParseColRef(ReadOnlySpan<char> s)
    {
        var col = 0;
        foreach (var c in s)
            col = col * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        return col - 1;
    }

    private static int ParseInt(string? s)
    {
        if (string.IsNullOrEmpty(s) || !int.TryParse(s, out var v)) return -1;
        return v;
    }

    private static int ParseRowIndex(string? r)
    {
        if (string.IsNullOrEmpty(r)) return 0;
        var i = 0;
        while (i < r.Length && char.IsLetter(r[i])) i++;
        return i < r.Length && int.TryParse(r.AsSpan(i), out var row) ? row - 1 : 0;
    }
}

internal record struct SheetData(
    string Name,
    List<RowData> Rows,
    List<ColInfo> ColInfos,
    List<MergeRange> MergeRanges,
    FreezePaneInfo? FreezePane,
    List<int> RowBreaks,
    List<int> ColBreaks,
    PageSetupInfo? PageSetup,
    PageMarginsInfo? PageMargins,
    List<HyperlinkInfo> Hyperlinks,
    List<CommentInfo> Comments,
    List<DataValidationInfo> DataValidations,
    List<ConditionalFormatInfo> ConditionalFormats);

internal record struct DataValidationInfo(
    List<(int FirstRow, int FirstCol, int LastRow, int LastCol)> Ranges,
    int Type,
    int Operator,
    string Formula1,
    string Formula2,
    bool AllowBlank,
    bool ShowPrompt,
    bool ShowError,
    string PromptTitle,
    string ErrorTitle,
    string PromptText,
    string ErrorText);

internal record struct DefinedNameInfo(int SheetIndex0Based, byte BuiltinIndex, int FirstRow, int FirstCol, int LastRow, int LastCol);

internal record struct ConditionalFormatInfo(
    List<(int FirstRow, int FirstCol, int LastRow, int LastCol)> Ranges,
    int Type,
    int Operator,
    string Formula1,
    string Formula2);

internal record struct CommentInfo(int Row, int Col, string Author, string Text);
internal record struct HyperlinkInfo(int FirstRow, int FirstCol, int LastRow, int LastCol, string Url);
internal record struct FreezePaneInfo(int ColSplit, int RowSplit, int TopRowVisible, int LeftColVisible);
internal record struct PageSetupInfo(bool Landscape, int Scale, int FitToWidth, int FitToHeight, int StartPageNumber);
internal record struct PageMarginsInfo(double Left, double Right, double Top, double Bottom, double Header, double Footer);
internal record struct ColInfo(int FirstCol, int LastCol, double Width, bool Hidden);
internal record struct MergeRange(int FirstRow, int FirstCol, int LastRow, int LastCol);
internal record struct RowData(int RowIndex, CellData[] Cells, double Height, bool Hidden);
internal record struct CellData(int Row, int Col, CellKind Kind, CellKind CachedKind, string? Value, int SstIndex, int StyleIndex, string? Formula);
internal enum CellKind { Empty, Number, String, SharedString, Boolean, Error, Formula }
