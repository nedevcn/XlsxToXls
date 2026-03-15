using System.Buffers;
using System.Buffers.Binary;
using System.Text;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Writes page setup information to BIFF8 format.
/// Generates PAGESETUP, PRINTGRIDLINES, PRINTHEADERS, and related records.
/// </summary>
public sealed class PageSetupWriter : IDisposable
{
    private const int DefaultBufferSize = 4096;

    private byte[] _buffer;
    private Memory<byte> _memory;
    private int _position;
    private bool _isRented;

    /// <summary>
    /// Creates a new PageSetupWriter with the specified buffer.
    /// </summary>
    public PageSetupWriter(byte[] buffer)
    {
        _buffer = buffer;
        _memory = buffer;
        _position = 0;
        _isRented = false;
    }

    /// <summary>
    /// Creates a PageSetupWriter using a pooled buffer.
    /// </summary>
    public static PageSetupWriter CreatePooled(out byte[] buffer, int minimumLength = DefaultBufferSize)
    {
        buffer = ArrayPool<byte>.Shared.Rent(minimumLength);
        return new PageSetupWriter(buffer) { _isRented = true };
    }

    /// <summary>
    /// Disposes the writer, returning the buffer to the pool if rented.
    /// </summary>
    public void Dispose()
    {
        if (_isRented)
        {
            ArrayPool<byte>.Shared.Return(_buffer);
            _isRented = false;
        }
    }

    /// <summary>
    /// Gets the written data as a span.
    /// </summary>
    public ReadOnlySpan<byte> GetData() => _memory.Span.Slice(0, _position);

    /// <summary>
    /// Writes all page setup records.
    /// </summary>
    public int WritePageSetup(PageSetupData setup, int sheetIndex)
    {
        // PAGESETUP record (0x00A1)
        WritePageSetupRecord(setup);

        // PRINTGRIDLINES record (0x002B)
        WritePrintGridlines(setup.PrintGridlines);

        // PRINTHEADERS record (0x002A)
        WritePrintHeaders(setup.PrintHeadings);

        // HCENTER record (0x0083)
        WriteHCenter(setup.CenterHorizontally);

        // VCENTER record (0x0084)
        WriteVCenter(setup.CenterVertically);

        // LEFTMARGIN record (0x0026)
        WriteLeftMargin(setup.Margins.Left);

        // RIGHTMARGIN record (0x0027)
        WriteRightMargin(setup.Margins.Right);

        // TOPMARGIN record (0x0028)
        WriteTopMargin(setup.Margins.Top);

        // BOTTOMMARGIN record (0x0029)
        WriteBottomMargin(setup.Margins.Bottom);

        // HEADER record (0x0014)
        if (!string.IsNullOrEmpty(setup.Header))
        {
            WriteHeader(setup.Header);
        }

        // FOOTER record (0x0015)
        if (!string.IsNullOrEmpty(setup.Footer))
        {
            WriteFooter(setup.Footer);
        }

        // HORIZONTALPAGEBREAKS (0x001B) - if needed
        // VERTICALPAGEBREAKS (0x001A) - if needed

        // PRINTAREA defined name
        if (setup.PrintArea.Count > 0)
        {
            WritePrintArea(setup.PrintArea, sheetIndex);
        }

        // PRINTTITLES defined name
        if (setup.PrintTitleRows != null || setup.PrintTitleColumns != null)
        {
            WritePrintTitles(setup.PrintTitleRows, setup.PrintTitleColumns, sheetIndex);
        }

        return _position;
    }

    private void WritePageSetupRecord(PageSetupData setup)
    {
        // PAGESETUP record structure:
        // Offset  Size  Description
        // 0       2     Paper size
        // 2       2     Scaling factor
        // 4       2     Start page number
        // 6       2     Fit to width
        // 8       2     Fit to height
        // 10      2     Options flags
        // 12      2     Print resolution
        // 14      2     Vertical print resolution
        // 16      8     Header margin
        // 24      8     Footer margin
        // 32      2     Number of copies

        const ushort recordId = 0x00A1;
        const ushort recordLength = 34;

        WriteRecordHeader(recordId, recordLength);

        // Paper size
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)setup.PaperSize);
        _position += 2;

        // Scale (10-400, or 0 for fit to pages)
        var scale = setup.Scale ?? 100;
        if (setup.FitToWidth.HasValue || setup.FitToHeight.HasValue)
        {
            scale = 0; // Use fit to pages
        }
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)scale);
        _position += 2;

        // Start page number
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)setup.FirstPageNumber);
        _position += 2;

        // Fit to width
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(setup.FitToWidth ?? 1));
        _position += 2;

        // Fit to height
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(setup.FitToHeight ?? 1));
        _position += 2;

        // Options flags
        ushort options = 0;
        if (setup.Orientation == PageOrientation.Landscape) options |= 0x0002;
        if (setup.BlackAndWhite) options |= 0x0008;
        if (setup.DraftQuality) options |= 0x0010;
        if (setup.PrintComments == PrintComments.AsDisplayed) options |= 0x0020;
        if (setup.PrintComments == PrintComments.AtEnd) options |= 0x0040;
        if (setup.PageOrder == PageOrder.OverThenDown) options |= 0x0080;
        if (setup.CellErrors == CellErrorPrint.Blank) options |= 0x0100;
        if (setup.CellErrors == CellErrorPrint.DashDash) options |= 0x0200;
        if (setup.CellErrors == CellErrorPrint.NA) options |= 0x0300;
        if (setup.CenterVertically) options |= 0x0800;
        if (setup.CenterHorizontally) options |= 0x0400;

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), options);
        _position += 2;

        // Print resolution
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)setup.PrintQuality);
        _position += 2;

        // Vertical print resolution
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)setup.PrintQuality);
        _position += 2;

        // Header margin
        BufferHelpers.WriteDoubleLittleEndian(_buffer.AsSpan(_position), setup.HeaderMargin);
        _position += 8;

        // Footer margin
        BufferHelpers.WriteDoubleLittleEndian(_buffer.AsSpan(_position), setup.FooterMargin);
        _position += 8;

        // Number of copies
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)setup.Copies);
        _position += 2;
    }

    private void WritePrintGridlines(bool print)
    {
        const ushort recordId = 0x002B;
        const ushort recordLength = 2;

        WriteRecordHeader(recordId, recordLength);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), print ? (ushort)1 : (ushort)0);
        _position += 2;
    }

    private void WritePrintHeaders(bool print)
    {
        const ushort recordId = 0x002A;
        const ushort recordLength = 2;

        WriteRecordHeader(recordId, recordLength);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), print ? (ushort)1 : (ushort)0);
        _position += 2;
    }

    private void WriteHCenter(bool center)
    {
        const ushort recordId = 0x0083;
        const ushort recordLength = 2;

        WriteRecordHeader(recordId, recordLength);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), center ? (ushort)1 : (ushort)0);
        _position += 2;
    }

    private void WriteVCenter(bool center)
    {
        const ushort recordId = 0x0084;
        const ushort recordLength = 2;

        WriteRecordHeader(recordId, recordLength);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), center ? (ushort)1 : (ushort)0);
        _position += 2;
    }

    private void WriteLeftMargin(double margin)
    {
        const ushort recordId = 0x0026;
        const ushort recordLength = 8;

        WriteRecordHeader(recordId, recordLength);
        BufferHelpers.WriteDoubleLittleEndian(_buffer.AsSpan(_position), margin);
        _position += 8;
    }

    private void WriteRightMargin(double margin)
    {
        const ushort recordId = 0x0027;
        const ushort recordLength = 8;

        WriteRecordHeader(recordId, recordLength);
        BufferHelpers.WriteDoubleLittleEndian(_buffer.AsSpan(_position), margin);
        _position += 8;
    }

    private void WriteTopMargin(double margin)
    {
        const ushort recordId = 0x0028;
        const ushort recordLength = 8;

        WriteRecordHeader(recordId, recordLength);
        BufferHelpers.WriteDoubleLittleEndian(_buffer.AsSpan(_position), margin);
        _position += 8;
    }

    private void WriteBottomMargin(double margin)
    {
        const ushort recordId = 0x0029;
        const ushort recordLength = 8;

        WriteRecordHeader(recordId, recordLength);
        BufferHelpers.WriteDoubleLittleEndian(_buffer.AsSpan(_position), margin);
        _position += 8;
    }

    private void WriteHeader(string header)
    {
        const ushort recordId = 0x0014;

        // Convert to BIFF8 string format
        var bytes = Encoding.UTF8.GetBytes(header);
        var charCount = bytes.Length;

        // Check if high byte is needed
        var hasHighByte = bytes.Any(b => b > 127);

        var recordLength = (ushort)(3 + charCount + (hasHighByte ? charCount : 0));
        WriteRecordHeader(recordId, recordLength);

        // Character count
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)charCount);
        _position += 2;

        // Options flags
        _buffer[_position++] = (byte)(hasHighByte ? 0x01 : 0x00);

        // String data
        if (hasHighByte)
        {
            // High byte format
            for (var i = 0; i < charCount; i++)
            {
                _buffer[_position++] = bytes[i];
                _buffer[_position++] = 0;
            }
        }
        else
        {
            // Single byte format
            bytes.CopyTo(_buffer.AsSpan(_position));
            _position += charCount;
        }
    }

    private void WriteFooter(string footer)
    {
        const ushort recordId = 0x0015;

        // Convert to BIFF8 string format
        var bytes = Encoding.UTF8.GetBytes(footer);
        var charCount = bytes.Length;

        // Check if high byte is needed
        var hasHighByte = bytes.Any(b => b > 127);

        var recordLength = (ushort)(3 + charCount + (hasHighByte ? charCount : 0));
        WriteRecordHeader(recordId, recordLength);

        // Character count
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)charCount);
        _position += 2;

        // Options flags
        _buffer[_position++] = (byte)(hasHighByte ? 0x01 : 0x00);

        // String data
        if (hasHighByte)
        {
            // High byte format
            for (var i = 0; i < charCount; i++)
            {
                _buffer[_position++] = bytes[i];
                _buffer[_position++] = 0;
            }
        }
        else
        {
            // Single byte format
            bytes.CopyTo(_buffer.AsSpan(_position));
            _position += charCount;
        }
    }

    private void WritePrintArea(List<CellRange> ranges, int sheetIndex)
    {
        // Write as DEFINEDNAME record (0x0018)
        const ushort recordId = 0x0018;

        // Build the formula reference
        var formula = BuildAreaFormula(ranges);
        var formulaBytes = Encoding.ASCII.GetBytes(formula);

        // Calculate record size
        var nameBytes = Encoding.ASCII.GetBytes("Print_Area");
        var nameLen = nameBytes.Length;

        // DEFINEDNAME structure:
        // 2  - Options flags
        // 1  - Keyboard shortcut
        // 1  - Name length
        // 2  - Formula size
        // 2  - Not used
        // 2  - Sheet index (1-based)
        // 4  - Not used
        // N  - Name
        // N  - Formula

        var recordLength = (ushort)(16 + nameLen + formulaBytes.Length);
        WriteRecordHeader(recordId, recordLength);

        // Options flags (hidden, function, etc.)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0x0000);
        _position += 2;

        // Keyboard shortcut
        _buffer[_position++] = 0;

        // Name length
        _buffer[_position++] = (byte)nameLen;

        // Formula size
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)formulaBytes.Length);
        _position += 2;

        // Not used
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 2;

        // Sheet index (1-based)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(sheetIndex + 1));
        _position += 2;

        // Not used
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 4;

        // Name
        nameBytes.CopyTo(_buffer.AsSpan(_position));
        _position += nameLen;

        // Formula
        formulaBytes.CopyTo(_buffer.AsSpan(_position));
        _position += formulaBytes.Length;
    }

    private void WritePrintTitles(CellRange? rows, CellRange? cols, int sheetIndex)
    {
        // Write as DEFINEDNAME record (0x0018)
        const ushort recordId = 0x0018;

        // Build the formula reference
        var formula = BuildTitlesFormula(rows, cols);
        var formulaBytes = Encoding.ASCII.GetBytes(formula);

        // Calculate record size
        var nameBytes = Encoding.ASCII.GetBytes("Print_Titles");
        var nameLen = nameBytes.Length;

        var recordLength = (ushort)(16 + nameLen + formulaBytes.Length);
        WriteRecordHeader(recordId, recordLength);

        // Options flags
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0x0000);
        _position += 2;

        // Keyboard shortcut
        _buffer[_position++] = 0;

        // Name length
        _buffer[_position++] = (byte)nameLen;

        // Formula size
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)formulaBytes.Length);
        _position += 2;

        // Not used
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 2;

        // Sheet index (1-based)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(sheetIndex + 1));
        _position += 2;

        // Not used
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 4;

        // Name
        nameBytes.CopyTo(_buffer.AsSpan(_position));
        _position += nameLen;

        // Formula
        formulaBytes.CopyTo(_buffer.AsSpan(_position));
        _position += formulaBytes.Length;
    }

    private static string BuildAreaFormula(List<CellRange> ranges)
    {
        if (ranges.Count == 0) return "";

        var parts = new List<string>();
        foreach (var range in ranges)
        {
            var startCol = GetColumnName(range.FirstCol);
            var startRow = range.FirstRow + 1;
            var endCol = GetColumnName(range.LastCol);
            var endRow = range.LastRow + 1;

            parts.Add($"{startCol}{startRow}:{endCol}{endRow}");
        }

        return string.Join(",", parts);
    }

    private static string BuildTitlesFormula(CellRange? rows, CellRange? cols)
    {
        var parts = new List<string>();

        if (rows != null)
        {
            parts.Add($"{rows.FirstRow + 1}:{rows.LastRow + 1}");
        }

        if (cols != null)
        {
            var startCol = GetColumnName(cols.FirstCol);
            var endCol = GetColumnName(cols.LastCol);
            parts.Add($"{startCol}:{endCol}");
        }

        return string.Join(",", parts);
    }

    private static string GetColumnName(int colIndex)
    {
        var name = "";
        var index = colIndex + 1;

        while (index > 0)
        {
            index--;
            name = (char)('A' + (index % 26)) + name;
            index /= 26;
        }

        return name;
    }

    private void WriteRecordHeader(ushort recordId, ushort recordLength)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordLength);
        _position += 2;
    }
}
