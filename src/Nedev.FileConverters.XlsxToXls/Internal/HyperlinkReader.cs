using System.IO.Compression;
using System.Xml;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Reads hyperlinks from XLSX files.
/// Parses the xl/worksheets/sheetN.xml.rels and xl/worksheets/_rels/sheetN.xml.rels files
/// to extract hyperlink information.
/// </summary>
public static class HyperlinkReader
{
    private static readonly XmlReaderSettings Settings = new()
    {
        DtdProcessing = DtdProcessing.Ignore,
        XmlResolver = null
    };

    /// <summary>
    /// Reads all hyperlinks from a worksheet in the XLSX archive.
    /// </summary>
    public static HyperlinkCollection ReadHyperlinks(ZipArchive archive, string sheetPath, Action<string>? log = null)
    {
        var collection = new HyperlinkCollection();

        try
        {
            log?.Invoke($"[HyperlinkReader] Reading hyperlinks from {sheetPath}");

            // Get the sheet number from the path
            var sheetName = Path.GetFileNameWithoutExtension(sheetPath);
            if (!int.TryParse(sheetName.Replace("sheet", "", StringComparison.OrdinalIgnoreCase), out var sheetNum))
            {
                log?.Invoke($"[HyperlinkReader] Could not determine sheet number from {sheetPath}");
                return collection;
            }

            // Read the worksheet XML to find hyperlink references
            var worksheetEntry = archive.GetEntry($"xl/worksheets/sheet{sheetNum}.xml");
            if (worksheetEntry == null)
            {
                log?.Invoke($"[HyperlinkReader] Worksheet not found: sheet{sheetNum}.xml");
                return collection;
            }

            // Read relationships file for hyperlinks
            var relsPath = $"xl/worksheets/_rels/sheet{sheetNum}.xml.rels";
            var relationships = ReadRelationships(archive, relsPath, log);

            // Parse hyperlinks from worksheet
            using var stream = worksheetEntry.Open();
            using var reader = XmlReader.Create(stream, Settings);

            var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element)
                    continue;

                if (reader.LocalName == "hyperlink" && reader.NamespaceURI == ns)
                {
                    var hyperlink = ParseHyperlink(reader, relationships, log);
                    if (hyperlink != null)
                    {
                        collection.Add(hyperlink);
                        log?.Invoke($"[HyperlinkReader] Found hyperlink at {hyperlink.GetCellReference()}: {hyperlink.Target}");
                    }
                }
            }

            log?.Invoke($"[HyperlinkReader] Found {collection.Count} hyperlinks");
        }
        catch (Exception ex)
        {
            log?.Invoke($"[HyperlinkReader] Error reading hyperlinks: {ex.Message}");
        }

        return collection;
    }

    /// <summary>
    /// Reads all hyperlinks from all worksheets in the workbook.
    /// </summary>
    public static Dictionary<string, HyperlinkCollection> ReadAllHyperlinks(
        ZipArchive archive,
        List<SheetInfo> sheets,
        Action<string>? log = null)
    {
        var result = new Dictionary<string, HyperlinkCollection>();

        foreach (var sheet in sheets)
        {
            var hyperlinks = ReadHyperlinks(archive, sheet.Path, log);
            if (hyperlinks.Count > 0)
            {
                result[sheet.Name] = hyperlinks;
            }
        }

        return result;
    }

    private static Dictionary<string, Relationship> ReadRelationships(
        ZipArchive archive,
        string relsPath,
        Action<string>? log)
    {
        var relationships = new Dictionary<string, Relationship>(StringComparer.OrdinalIgnoreCase);

        try
        {
            var entry = archive.GetEntry(relsPath);
            if (entry == null)
            {
                log?.Invoke($"[HyperlinkReader] Relationships file not found: {relsPath}");
                return relationships;
            }

            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, Settings);

            var ns = "http://schemas.openxmlformats.org/package/2006/relationships";

            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element)
                    continue;

                if (reader.LocalName == "Relationship" && reader.NamespaceURI == ns)
                {
                    var id = reader.GetAttribute("Id");
                    var type = reader.GetAttribute("Type");
                    var target = reader.GetAttribute("Target");
                    var targetMode = reader.GetAttribute("TargetMode");

                    if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(type))
                    {
                        relationships[id] = new Relationship
                        {
                            Id = id,
                            Type = type,
                            Target = target ?? string.Empty,
                            TargetMode = targetMode
                        };
                    }
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[HyperlinkReader] Error reading relationships: {ex.Message}");
        }

        return relationships;
    }

    private static HyperlinkData? ParseHyperlink(
        XmlReader reader,
        Dictionary<string, Relationship> relationships,
        Action<string>? log)
    {
        try
        {
            // Get cell reference (required)
            var refAttr = reader.GetAttribute("ref");
            if (string.IsNullOrEmpty(refAttr))
            {
                log?.Invoke("[HyperlinkReader] Hyperlink missing ref attribute");
                return null;
            }

            // Parse cell reference
            var (row, col) = CellReferenceConverter.FromReference(refAttr);

            // Get relationship ID
            var rid = reader.GetAttribute("r:id");

            // Get display text (optional)
            var display = reader.GetAttribute("display");

            // Get tooltip (optional)
            var tooltip = reader.GetAttribute("tooltip");

            // Get location within document (optional)
            var location = reader.GetAttribute("location");

            var hyperlink = new HyperlinkData
            {
                Row = row,
                Column = col,
                DisplayText = display,
                Tooltip = tooltip,
                Location = location
            };

            // Determine hyperlink type and target
            if (!string.IsNullOrEmpty(rid) && relationships.TryGetValue(rid, out var rel))
            {
                hyperlink.Target = rel.Target;
                hyperlink.Type = DetermineHyperlinkType(rel);
            }
            else if (!string.IsNullOrEmpty(location))
            {
                // Internal link to a location in the workbook
                hyperlink.Type = HyperlinkType.Internal;
                hyperlink.Target = location;
            }
            else
            {
                log?.Invoke($"[HyperlinkReader] Hyperlink at {refAttr} has no target");
                return null;
            }

            return hyperlink;
        }
        catch (Exception ex)
        {
            log?.Invoke($"[HyperlinkReader] Error parsing hyperlink: {ex.Message}");
            return null;
        }
    }

    private static HyperlinkType DetermineHyperlinkType(Relationship rel)
    {
        var target = rel.Target?.ToLowerInvariant() ?? string.Empty;

        // Check target mode
        if (rel.TargetMode == "External")
        {
            if (target.StartsWith("mailto:"))
                return HyperlinkType.Email;

            if (target.StartsWith("http://") || target.StartsWith("https://"))
                return HyperlinkType.Url;

            if (target.StartsWith("file://"))
                return HyperlinkType.File;

            if (target.StartsWith("\\\\"))
                return HyperlinkType.Unc;

            // Default to URL for external links
            return HyperlinkType.Url;
        }
        else
        {
            // Internal link
            if (target.Contains("!"))
                return HyperlinkType.Worksheet;

            return HyperlinkType.Internal;
        }
    }

    /// <summary>
    /// Reads the shared strings table to find display text for hyperlinks.
    /// </summary>
    public static string? GetDisplayTextFromCell(ZipArchive archive, int sheetNum, int row, int col, Action<string>? log = null)
    {
        try
        {
            var worksheetEntry = archive.GetEntry($"xl/worksheets/sheet{sheetNum}.xml");
            if (worksheetEntry == null)
                return null;

            using var stream = worksheetEntry.Open();
            using var reader = XmlReader.Create(stream, Settings);

            var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var inSheetData = false;
            var currentRow = -1;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.LocalName == "sheetData" && reader.NamespaceURI == ns)
                    {
                        inSheetData = true;
                    }
                    else if (inSheetData && reader.LocalName == "row" && reader.NamespaceURI == ns)
                    {
                        var rowAttr = reader.GetAttribute("r");
                        if (rowAttr != null && int.TryParse(rowAttr, out var r))
                        {
                            currentRow = r - 1; // Convert to 0-based
                        }
                    }
                    else if (inSheetData && reader.LocalName == "c" && reader.NamespaceURI == ns)
                    {
                        var refAttr = reader.GetAttribute("r");
                        if (refAttr != null)
                        {
                            var (cellRow, cellCol) = CellReferenceConverter.FromReference(refAttr);
                            if (cellRow == row && cellCol == col)
                            {
                                // Found the cell, read its value
                                return ReadCellValue(reader, archive, ns, log);
                            }
                        }
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.LocalName == "sheetData" && reader.NamespaceURI == ns)
                    {
                        break;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[HyperlinkReader] Error reading cell value: {ex.Message}");
        }

        return null;
    }

    private static string? ReadCellValue(XmlReader reader, ZipArchive archive, string ns, Action<string>? log)
    {
        var type = reader.GetAttribute("t");

        // Read to the value element
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "v" && reader.NamespaceURI == ns)
            {
                var value = reader.ReadElementContentAsString();

                if (type == "s")
                {
                    // Shared string
                    if (int.TryParse(value, out var sharedStringIndex))
                    {
                        return ReadSharedString(archive, sharedStringIndex, log);
                    }
                }
                else if (type == "str")
                {
                    // Inline string
                    return value;
                }
                else
                {
                    // Regular value
                    return value;
                }
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "is" && reader.NamespaceURI == ns)
            {
                // Inline rich text
                return ReadInlineString(reader, ns);
            }
            else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "c")
            {
                break;
            }
        }

        return null;
    }

    private static string? ReadSharedString(ZipArchive archive, int index, Action<string>? log)
    {
        try
        {
            var entry = archive.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
                return null;

            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, Settings);

            var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var currentIndex = 0;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si" && reader.NamespaceURI == ns)
                {
                    if (currentIndex == index)
                    {
                        return ReadStringItem(reader, ns);
                    }
                    currentIndex++;
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[HyperlinkReader] Error reading shared string: {ex.Message}");
        }

        return null;
    }

    private static string ReadStringItem(XmlReader reader, string ns)
    {
        var result = string.Empty;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t" && reader.NamespaceURI == ns)
            {
                result += reader.ReadElementContentAsString();
            }
            else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "si")
            {
                break;
            }
        }

        return result;
    }

    private static string ReadInlineString(XmlReader reader, string ns)
    {
        var result = string.Empty;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t" && reader.NamespaceURI == ns)
            {
                result += reader.ReadElementContentAsString();
            }
            else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "is")
            {
                break;
            }
        }

        return result;
    }

    private class Relationship
    {
        public string Id { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Target { get; set; } = string.Empty;
        public string? TargetMode { get; set; }
    }
}

/// <summary>
/// Information about a worksheet for hyperlink reading.
/// </summary>
public class SheetInfo
{
    public string Name { get; set; } = string.Empty;
    public string Path { get; set; } = string.Empty;
    public int SheetId { get; set; }
}
