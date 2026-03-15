using System.Buffers;
using System.Buffers.Binary;
using System.Text;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Writes data validation rules to BIFF8 format.
/// Generates DV (Data Validation) and DVAL (Data Validation List) records.
/// </summary>
public sealed class DataValidationWriter : IDisposable
{
    private const int DefaultBufferSize = 8192;

    private byte[] _buffer;
    private Memory<byte> _memory;
    private int _position;
    private bool _isRented;

    /// <summary>
    /// Creates a new DataValidationWriter with the specified buffer.
    /// </summary>
    public DataValidationWriter(byte[] buffer)
    {
        _buffer = buffer;
        _memory = buffer;
        _position = 0;
        _isRented = false;
    }

    /// <summary>
    /// Creates a DataValidationWriter using a pooled buffer.
    /// </summary>
    public static DataValidationWriter CreatePooled(out byte[] buffer, int minimumLength = DefaultBufferSize)
    {
        buffer = ArrayPool<byte>.Shared.Rent(minimumLength);
        return new DataValidationWriter(buffer) { _isRented = true };
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
    /// Writes all data validation records.
    /// </summary>
    public int WriteDataValidations(List<DataValidationData> validations, int sheetIndex)
    {
        if (validations.Count == 0)
        {
            return 0;
        }

        // Write DVAL record first (container)
        WriteDvalRecord(validations.Count);

        // Write DV records for each validation
        foreach (var validation in validations)
        {
            WriteDvRecord(validation);
        }

        return _position;
    }

    private void WriteDvalRecord(int validationCount)
    {
        // DVAL record (0x01B2) - Data Validation List
        // This record appears before all DV records

        const ushort recordId = 0x01B2;
        const ushort recordLength = 10;

        WriteRecordHeader(recordId, recordLength);

        // Flags (2 bytes) - always 0
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 2;

        // Number of DV records (4 bytes)
        BinaryPrimitives.WriteInt32LittleEndian(_buffer.AsSpan(_position), validationCount);
        _position += 4;

        // Object ID (4 bytes) - always 0xFFFFFFFF
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), 0xFFFFFFFF);
        _position += 4;
    }

    private void WriteDvRecord(DataValidationData validation)
    {
        // DV record (0x01BE) - Data Validation
        // This is a complex record with variable length

        const ushort recordId = 0x01BE;
        var startPosition = _position;

        // Skip header for now, will write after calculating length
        _position += 4;

        // DV record structure:
        // 4 bytes: Flags
        // 4 bytes: Prompt title length and string
        // 4 bytes: Error title length and string
        // 4 bytes: Prompt text length and string
        // 4 bytes: Error text length and string
        // 2 bytes: Formula 1 size
        // N bytes: Formula 1
        // 2 bytes: Formula 2 size
        // N bytes: Formula 2
        // 4 bytes: Cell range count
        // N bytes: Cell ranges

        // Write flags
        var flags = CalculateDvFlags(validation);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), flags);
        _position += 4;

        // Write prompt title (input title)
        WriteDvString(validation.InputTitle ?? "");

        // Write error title
        WriteDvString(validation.ErrorTitle ?? "");

        // Write prompt text (input message)
        WriteDvString(validation.InputMessage ?? "");

        // Write error text (error message)
        WriteDvString(validation.ErrorMessage ?? "");

        // Write formula 1
        WriteDvFormula(validation.Formula1);

        // Write formula 2
        WriteDvFormula(validation.Formula2);

        // Write cell ranges
        WriteDvRanges(validation.Ranges);

        // Calculate and write record length
        var recordLength = (ushort)(_position - startPosition - 4);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(startPosition), recordId);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(startPosition + 2), recordLength);
    }

    private uint CalculateDvFlags(DataValidationData validation)
    {
        uint flags = 0;

        // Type (4 bits, bits 0-3)
        flags |= (uint)((byte)validation.Type & 0x0F);

        // Error style (2 bits, bits 4-5)
        flags |= (uint)(((byte)validation.ErrorAlertType & 0x03) << 4);

        // Operator (4 bits, bits 16-19)
        flags |= (uint)(((byte)validation.Operator & 0x0F) << 16);

        // Allow blank (bit 8)
        if (validation.AllowBlank) flags |= 0x00000100;

        // Suppress drop down (bit 9) - inverted, 1 means suppress
        if (validation.SuppressDropDown) flags |= 0x00000200;

        // Show input message (bit 18)
        if (validation.ShowInputMessage) flags |= 0x00040000;

        // Show error message (bit 19)
        if (validation.ShowErrorMessage) flags |= 0x00080000;

        return flags;
    }

    private void WriteDvString(string value)
    {
        // DV strings use a special format:
        // 2 bytes: character count
        // 1 byte: options (high byte flag)
        // N bytes: string data

        if (string.IsNullOrEmpty(value))
        {
            // Empty string
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
            _position += 2;
            _buffer[_position++] = 0; // Options
            return;
        }

        // Check if we need high byte (Unicode)
        var needsHighByte = value.Any(c => c > 255);

        if (needsHighByte)
        {
            // Unicode string
            var bytes = Encoding.Unicode.GetBytes(value);
            var charCount = value.Length;

            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)charCount);
            _position += 2;
            _buffer[_position++] = 0x01; // High byte flag

            bytes.CopyTo(_buffer.AsSpan(_position));
            _position += bytes.Length;
        }
        else
        {
            // ASCII string
            var bytes = Encoding.ASCII.GetBytes(value);
            var charCount = bytes.Length;

            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)charCount);
            _position += 2;
            _buffer[_position++] = 0x00; // No high byte

            bytes.CopyTo(_buffer.AsSpan(_position));
            _position += bytes.Length;
        }
    }

    private void WriteDvFormula(string? formula)
    {
        // Formula size (2 bytes)
        if (string.IsNullOrEmpty(formula))
        {
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
            _position += 2;
            return;
        }

        // Parse and write formula
        // For simplicity, we'll write the formula as a string token
        // In a full implementation, this would parse the formula into tokens

        var formulaBytes = Encoding.ASCII.GetBytes(formula);
        var formulaSize = (ushort)(formulaBytes.Length + 1); // +1 for token type

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), formulaSize);
        _position += 2;

        // Token type - string constant (0x17)
        _buffer[_position++] = 0x17;

        // String length
        _buffer[_position++] = (byte)formulaBytes.Length;

        // String data
        formulaBytes.CopyTo(_buffer.AsSpan(_position));
        _position += formulaBytes.Length;
    }

    private void WriteDvRanges(List<CellRange> ranges)
    {
        // Number of ranges (4 bytes)
        BinaryPrimitives.WriteInt32LittleEndian(_buffer.AsSpan(_position), ranges.Count);
        _position += 4;

        // Write each range
        foreach (var range in ranges)
        {
            // First row (2 bytes)
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)range.FirstRow);
            _position += 2;

            // Last row (2 bytes)
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)range.LastRow);
            _position += 2;

            // First column (2 bytes)
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)range.FirstCol);
            _position += 2;

            // Last column (2 bytes)
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)range.LastCol);
            _position += 2;
        }
    }

    private void WriteRecordHeader(ushort recordId, ushort recordLength)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordLength);
        _position += 2;
    }
}
