using System.Buffers;
using System.Buffers.Binary;
using System.Text;
using Nedev.XlsxToXls.Internal;

namespace Nedev.XlsxToXls;

/// <summary>
/// High-performance XLSX to XLS converter with zero third-party dependencies.
/// Uses streaming and pooled buffers for optimal memory usage.
/// </summary>
public static class XlsxToXlsConverter
{
    /// <summary>
    /// Converts XLSX stream to XLS format and writes to output stream.
    /// </summary>
    /// <param name="xlsxStream">Input XLSX stream (readable, seekable recommended)</param>
    /// <param name="xlsStream">Output XLS stream (writable)</param>
    public static void Convert(Stream xlsxStream, Stream xlsStream)
    {
        var (sharedStrings, sheets, styles, definedNames) = XlsxReader.Read(xlsxStream);
        var stylesData = styles ?? CreateDefaultStyles();
        var biffSize = EstimateBiffSize(sharedStrings, sheets, stylesData, definedNames);
        var buffer = ArrayPool<byte>.Shared.Rent(Math.Max(biffSize, 256 * 1024));
        try
        {
            var written = WriteBiff(buffer.AsSpan(), sharedStrings, sheets, stylesData, definedNames);
            var ole = new OleCompoundWriter("Workbook");
            ole.WriteStream(buffer.AsSpan(0, written));
            ole.WriteTo(xlsStream);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    /// <summary>
    /// Converts XLSX file to XLS file.
    /// </summary>
    public static void ConvertFile(string xlsxPath, string xlsPath)
    {
        using var xlsx = File.OpenRead(xlsxPath);
        using var xls = File.Create(xlsPath);
        Convert(xlsx, xls);
    }

    private static StylesData CreateDefaultStyles()
    {
        var s = new StylesData();
        s.EnsureMinFonts();
        s.CellXfs.Add(new CellXfInfo(0, 0, 0, 0));
        return s;
    }

    private static int EstimateBiffSize(List<ReadOnlyMemory<char>> sharedStrings, List<SheetData> sheets, StylesData styles, List<DefinedNameInfo> definedNames)
    {
        var n = 1024;
        foreach (var s in sharedStrings)
            n += 16 + s.Length * 2;
        foreach (var sheet in sheets)
        {
            n += 64 * sheet.ColInfos.Count + 32 * sheet.MergeRanges.Count;
            foreach (var hl in sheet.Hyperlinks)
                n += 60 + hl.Url.Length * 2;
            foreach (var comment in sheet.Comments)
            {
                n += 64 + 25 + comment.Author.Length * 2;
                if (!string.IsNullOrEmpty(comment.Text))
                    n += 22 + 5 + comment.Text.Length * 2 + 20;
            }
            foreach (var dv in sheet.DataValidations)
                n += 30 + dv.Ranges.Count * 8 + (dv.PromptTitle.Length + dv.ErrorTitle.Length + dv.PromptText.Length + dv.ErrorText.Length) * 2 + dv.Formula1.Length + dv.Formula2.Length;
        }
        if (sheets.Count > 0)
            n += 8 + (4 + 2 + sheets.Count * 6) + (definedNames.Count > 0 ? definedNames.Count * (4 + 27) : 0);
        foreach (var sheet in sheets)
        {
            foreach (var row in sheet.Rows)
                n += row.Cells.Length * 40 + 32;
        }
        n += styles.Fonts.Count * 64 + styles.NumFmts.Count * 32 + styles.CellXfs.Count * 24;
        return Math.Max(n, 256 * 1024);
    }

    private static int WriteBiff(Span<byte> buffer, List<ReadOnlyMemory<char>> sharedStrings, List<SheetData> sheets, StylesData styles, List<DefinedNameInfo> definedNames)
    {
        var pos = 0;
        var bw = new BiffWriter(buffer);
        bw.WriteBofWorkbook();
        bw.WriteCodepage(0x04E4);

        foreach (var font in styles.Fonts)
        {
            var twips = (ushort)(font.Height * 20);
            bw.WriteFont(font.Name, twips, font.Bold, font.Italic);
        }
        pos = bw.Position;

        bw.WriteBuiltinFmtCount(0);
        for (var i = 0; i < styles.NumFmts.Count; i++)
        {
            var nf = styles.NumFmts[i];
            bw.WriteFormat(164 + i, nf.FormatCode);
        }
        pos = bw.Position;

        for (var i = 0; i < 15; i++)
            bw.WriteXf(0, 0, false);
        foreach (var xf in styles.CellXfs)
        {
            var fontIdx = styles.GetBiffFontIndex(xf.FontId);
            var fmtIdx = styles.GetBiffFormatIndex(xf.NumFmtId);
            bw.WriteXf(fontIdx, fmtIdx, true);
        }
        pos = bw.Position;

        if (sharedStrings.Count > 0)
        {
            var sstSize = WriteSstCorrect(buffer.Slice(pos), sharedStrings);
            pos += sstSize;
        }

        var sheetIndexByName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < sheets.Count; i++)
            if (!sheetIndexByName.ContainsKey(sheets[i].Name))
                sheetIndexByName[sheets[i].Name] = i;

        var cp1252 = Encoding.GetEncoding(1252);
        var tempBuf = ArrayPool<byte>.Shared.Rent(64 * 1024);
        var sheetSizes = new List<int>();
        try
        {
            for (var i = 0; i < sheets.Count; i++)
            {
                var sz = WriteSheet(tempBuf.AsSpan(), sheets[i], sharedStrings, styles, i, sheetIndexByName);
                sheetSizes.Add(sz);
            }

            var boundsheetLen = 0;
            foreach (var sheet in sheets)
            {
                var name = TruncateSheetName(sheet.Name);
                var nameBytes = Math.Min(cp1252.GetByteCount(name), 31);
                boundsheetLen += 4 + 4 + 1 + 1 + 1 + nameBytes;
            }
            var externNameLen = sheets.Count > 0 ? (8 + (4 + 2 + sheets.Count * 6) + definedNames.Count * (4 + 27)) : 0;
            var eofLen = 4;
            var firstSheetOffset = pos + boundsheetLen + externNameLen + eofLen;

            var offset = firstSheetOffset;
            for (var i = 0; i < sheets.Count; i++)
            {
                var sheet = sheets[i];
                bw = new BiffWriter(buffer.Slice(pos));
                bw.WriteBoundSheet(offset, TruncateSheetName(sheet.Name), sheet.Visibility);
                var name = TruncateSheetName(sheet.Name);
                var nameBytes = Math.Min(cp1252.GetByteCount(name), 31);
                pos += 4 + 4 + 1 + 1 + 1 + nameBytes;
                offset += sheetSizes[i];
            }

            if (sheets.Count > 0)
            {
                bw = new BiffWriter(buffer.Slice(pos));
                bw.WriteSupBookInternalRef(sheets.Count);
                pos += 8;
                bw = new BiffWriter(buffer.Slice(pos));
                bw.WriteExternSheet(sheets.Count);
                pos += 4 + 2 + sheets.Count * 6;
                if (definedNames.Count > 0)
                    foreach (var dn in definedNames)
                    {
                        bw = new BiffWriter(buffer.Slice(pos));
                        bw.WriteNameBuiltin(dn, (ushort)dn.SheetIndex0Based);
                        pos += 4 + 27;
                    }
            }

            bw = new BiffWriter(buffer.Slice(pos));
            bw.WriteEof();
            pos += 4;

            for (var i = 0; i < sheets.Count; i++)
            {
                var sz = WriteSheet(tempBuf.AsSpan(), sheets[i], sharedStrings, styles, i, sheetIndexByName);
                tempBuf.AsSpan(0, sz).CopyTo(buffer.Slice(pos));
                pos += sz;
            }
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(tempBuf);
        }

        return pos;
    }

    private static (int FirstRow, int LastRowPlus1, int FirstCol, int LastColPlus1) ComputeUsedRange(SheetData sheet)
    {
        var firstRow = int.MaxValue;
        var lastRow = -1;
        var firstCol = int.MaxValue;
        var lastCol = -1;
        foreach (var row in sheet.Rows)
        {
            if (row.RowIndex < firstRow) firstRow = row.RowIndex;
            if (row.RowIndex > lastRow) lastRow = row.RowIndex;
            foreach (var c in row.Cells)
            {
                if (c.Col < firstCol) firstCol = c.Col;
                if (c.Col > lastCol) lastCol = c.Col;
            }
        }
        if (firstRow == int.MaxValue) firstRow = 0;
        if (lastRow < 0) lastRow = 0;
        if (firstCol == int.MaxValue) firstCol = 0;
        if (lastCol < 0) lastCol = 0;
        return (firstRow, lastRow + 1, firstCol, lastCol + 1);
    }

    private static byte MapXlsxErrorToBiff(string value)
    {
        return value switch
        {
            "#NULL!" => 0x00,
            "#DIV/0!" => 0x07,
            "#VALUE!" => 0x0F,
            "#REF!" => 0x17,
            "#NAME?" => 0x1D,
            "#NUM!" => 0x24,
            "#N/A" => 0x2A,
            _ => 0x2A
        };
    }

    private static string TruncateSheetName(string name) =>
        name.Length > 31 ? name[..31] : name;

    private const int BiffMaxRecordData = 8224;

    private static int WriteSstCorrect(Span<byte> buffer, List<ReadOnlyMemory<char>> strings)
    {
        var outPos = 0;
        var dataLen = 8;
        var isSst = true;

        BinaryPrimitives.WriteInt32LittleEndian(buffer.Slice(4), strings.Count);
        BinaryPrimitives.WriteInt32LittleEndian(buffer.Slice(8), strings.Count);

        foreach (var s in strings)
        {
            var span = s.Span;
            var needs16 = false;
            foreach (var c in span)
            {
                if (c > 255) { needs16 = true; break; }
            }
            var strBytes = 3 + (needs16 ? span.Length * 2 : span.Length);
            if (dataLen + strBytes > BiffMaxRecordData)
            {
                BinaryPrimitives.WriteUInt16LittleEndian(buffer.Slice(outPos), (ushort)(isSst ? 0x00FC : 0x003C));
                BinaryPrimitives.WriteUInt16LittleEndian(buffer.Slice(outPos + 2), (ushort)dataLen);
                outPos += 4 + dataLen;
                dataLen = 0;
                isSst = false;
            }
            var p = outPos + 4 + dataLen;
            BinaryPrimitives.WriteUInt16LittleEndian(buffer.Slice(p), (ushort)span.Length);
            buffer[p + 2] = (byte)(needs16 ? 0 : 1);
            if (needs16)
            {
                for (var i = 0; i < span.Length; i++)
                    BinaryPrimitives.WriteUInt16LittleEndian(buffer.Slice(p + 3 + i * 2), span[i]);
            }
            else
            {
                for (var i = 0; i < span.Length; i++)
                    buffer[p + 3 + i] = (byte)span[i];
            }
            dataLen += strBytes;
        }

        BinaryPrimitives.WriteUInt16LittleEndian(buffer.Slice(outPos), (ushort)(isSst ? 0x00FC : 0x003C));
        BinaryPrimitives.WriteUInt16LittleEndian(buffer.Slice(outPos + 2), (ushort)dataLen);
        return outPos + 4 + dataLen;
    }

    private static int WriteSheet(Span<byte> buffer, SheetData sheet, List<ReadOnlyMemory<char>> sharedStrings, StylesData styles, int sheetIndex, IReadOnlyDictionary<string, int> sheetIndexByName)
    {
        var bw = new BiffWriter(buffer);
        bw.WriteBofWorksheet();
        bw.WriteWsBool();
        bw.WriteDefColWidth(8);

        foreach (var col in sheet.ColInfos)
            bw.WriteColInfo(col.FirstCol, col.LastCol, col.Width > 0 ? col.Width : 10, col.Hidden);

        var (firstRow, lastRowPlus1, firstCol, lastColPlus1) = ComputeUsedRange(sheet);
        bw.WriteDimension(firstRow, lastRowPlus1, firstCol, lastColPlus1);

        foreach (var row in sheet.Rows)
        {
            if (row.Height > 0 || row.Hidden)
                bw.WriteRow(row.RowIndex, row.Height, row.Hidden);
            foreach (var cell in row.Cells)
            {
                var xfIdx = (ushort)styles.GetBiffXfIndex(cell.StyleIndex);
                switch (cell.Kind)
                {
                    case CellKind.Formula:
                    {
                        var rgce = CompileFormulaTokens(cell.Formula, sheetIndex, sheetIndexByName);
                        if (rgce.Length == 0)
                        {
                            // Fallback: emit cached value only
                            WriteCachedValue(bw, cell, xfIdx, sharedStrings);
                            break;
                        }
                        var cachedKind = cell.CachedKind;
                        var cachedNum = 0d;
                        var cachedBool = false;
                        var cachedErr = (byte)0;
                        var cachedStr = "";
                        if (cachedKind == CellKind.Number && cell.Value != null &&
                            double.TryParse(cell.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var n))
                            cachedNum = n;
                        else if (cachedKind == CellKind.Boolean && cell.Value != null)
                            cachedBool = cell.Value == "1";
                        else if (cachedKind == CellKind.Error && cell.Value != null)
                            cachedErr = MapXlsxErrorToBiff(cell.Value);
                        else if (cachedKind == CellKind.String && cell.Value != null)
                            cachedStr = cell.Value;
                        else if (cachedKind == CellKind.SharedString && cell.SstIndex >= 0 && cell.SstIndex < sharedStrings.Count)
                            cachedStr = sharedStrings[cell.SstIndex].ToString();
                        bw.WriteFormula(cell.Row, cell.Col, xfIdx, rgce, cachedKind, cachedNum, cachedBool, cachedErr, cachedStr.AsSpan());
                        break;
                    }
                    case CellKind.Empty:
                        bw.WriteBlank(cell.Row, cell.Col, xfIdx);
                        break;
                    case CellKind.Number when cell.Value != null &&
                        double.TryParse(cell.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var num):
                        bw.WriteNumber(cell.Row, cell.Col, num, xfIdx);
                        break;
                    case CellKind.Boolean when cell.Value != null:
                        bw.WriteBool(cell.Row, cell.Col, cell.Value == "1", xfIdx);
                        break;
                    case CellKind.Error when cell.Value != null:
                        bw.WriteError(cell.Row, cell.Col, MapXlsxErrorToBiff(cell.Value), xfIdx);
                        break;
                    case CellKind.SharedString when cell.SstIndex >= 0 && cell.SstIndex < sharedStrings.Count:
                        bw.WriteLabelSst(cell.Row, cell.Col, cell.SstIndex, xfIdx);
                        break;
                    default:
                        if (cell.Value != null || (cell.SstIndex >= 0 && cell.SstIndex < sharedStrings.Count))
                        {
                            var text = cell.Value ?? sharedStrings[cell.SstIndex].ToString();
                            bw.WriteLabel(cell.Row, cell.Col, text.AsSpan(), xfIdx);
                        }
                        break;
                }
            }
        }

        if (sheet.MergeRanges.Count > 0)
        {
            var arr = sheet.MergeRanges.Select(m => (m.FirstRow, m.FirstCol, m.LastRow, m.LastCol)).ToArray();
            bw.WriteMergeCells(arr);
        }

        if (sheet.RowBreaks.Count > 0)
        {
            var breaks = sheet.RowBreaks.OrderBy(x => x).Distinct().Select(r => (r, 0, 255)).ToArray();
            bw.WriteHorizontalPageBreaks(breaks);
        }
        if (sheet.ColBreaks.Count > 0)
        {
            var breaks = sheet.ColBreaks.OrderBy(x => x).Distinct().Select(c => (c, 0, 16383)).ToArray();
            bw.WriteVerticalPageBreaks(breaks);
        }
        if (sheet.PageMargins is { } margins)
        {
            bw.WriteLeftMargin(margins.Left);
            bw.WriteRightMargin(margins.Right);
            bw.WriteTopMargin(margins.Top);
            bw.WriteBottomMargin(margins.Bottom);
        }
        if (sheet.PageSetup is { } setup)
        {
            bw.WritePageSetup(setup.Landscape, setup.Scale, setup.StartPageNumber, setup.FitToWidth, setup.FitToHeight, sheet.PageMargins?.Header ?? 0.3, sheet.PageMargins?.Footer ?? 0.3);
        }
        if (sheet.FreezePane is { } fp)
        {
            bw.WriteWindow2(freezePanes: true);
            var activePane = (byte)((fp.RowSplit > 0 && fp.ColSplit > 0) ? 0 : (fp.RowSplit > 0 ? 2 : (fp.ColSplit > 0 ? 1 : 0)));
            bw.WritePane((ushort)fp.ColSplit, (ushort)fp.RowSplit, (ushort)fp.TopRowVisible, (ushort)fp.LeftColVisible, activePane);
        }
        else
            bw.WriteWindow2(freezePanes: false);

        foreach (var hl in sheet.Hyperlinks)
            bw.WriteHyperlink(hl.FirstRow, hl.FirstCol, hl.LastRow, hl.LastCol, hl.Url.AsSpan());

        ushort shapeId = 1;
        foreach (var comment in sheet.Comments)
        {
            bw.WriteObjNote(shapeId);
            if (!string.IsNullOrEmpty(comment.Text))
                bw.WriteTxoWithText(comment.Text.AsSpan());
            bw.WriteNote(comment.Row, comment.Col, false, shapeId, comment.Author.AsSpan());
            shapeId++;
        }

        if (sheet.DataValidations.Count > 0)
        {
            bw.WriteDatavalidations(sheet.DataValidations.Count);
            foreach (var dv in sheet.DataValidations)
            {
                var f1 = CompileDataValidationFormula(dv, dv.Formula1, sheetIndex, sheetIndexByName);
                var f2 = CompileFormulaTokens(dv.Formula2, sheetIndex, sheetIndexByName);
                bw.WriteDatavalidation(dv, f1, f2);
            }
        }

        bw.WriteEof();
        return bw.Position;
    }

    private static void WriteCachedValue(BiffWriter bw, CellData cell, ushort xfIdx, List<ReadOnlyMemory<char>> sharedStrings)
    {
        switch (cell.CachedKind)
        {
            case CellKind.Number when cell.Value != null &&
                double.TryParse(cell.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var num):
                bw.WriteNumber(cell.Row, cell.Col, num, xfIdx);
                return;
            case CellKind.Boolean when cell.Value != null:
                bw.WriteBool(cell.Row, cell.Col, cell.Value == "1", xfIdx);
                return;
            case CellKind.Error when cell.Value != null:
                bw.WriteError(cell.Row, cell.Col, MapXlsxErrorToBiff(cell.Value), xfIdx);
                return;
            case CellKind.SharedString when cell.SstIndex >= 0 && cell.SstIndex < sharedStrings.Count:
                bw.WriteLabel(cell.Row, cell.Col, sharedStrings[cell.SstIndex].Span, xfIdx);
                return;
            case CellKind.String:
            default:
                if (cell.Value != null)
                    bw.WriteLabel(cell.Row, cell.Col, cell.Value.AsSpan(), xfIdx);
                else
                    bw.WriteBlank(cell.Row, cell.Col, xfIdx);
                return;
        }
    }

    private static byte[] CompileFormulaTokens(string? formula, int sheetIndex, IReadOnlyDictionary<string, int> sheetIndexByName)
    {
        if (string.IsNullOrWhiteSpace(formula)) return [];
        try
        {
            return FormulaCompiler.Compile(formula, sheetIndex, sheetIndexByName);
        }
        catch
        {
            return [];
        }
    }

    private static byte[] CompileDataValidationFormula(DataValidationInfo dv, string formula, int sheetIndex, IReadOnlyDictionary<string, int> sheetIndexByName)
    {
        if (string.IsNullOrWhiteSpace(formula)) return [];
        // List type: keep explicit list as ptgStr (tStr)
        if (dv.Type == 3 && formula.IndexOf(',') >= 0 && !formula.Contains('!') && !formula.Contains(':') && !formula.StartsWith("=", StringComparison.Ordinal))
            return FormulaCompiler.BuildExplicitListToken(formula);
        return CompileFormulaTokens(formula, sheetIndex, sheetIndexByName);
    }
}
