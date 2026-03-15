using System.Buffers;
using System.Buffers.Binary;
using System.Text;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Writes hyperlinks to BIFF8 format.
/// Generates HLINK records for each hyperlink.
/// </summary>
public sealed class HyperlinkWriter : IDisposable
{
    private const int DefaultBufferSize = 65536;

    private byte[] _buffer;
    private Memory<byte> _memory;
    private int _position;
    private bool _isRented;

    /// <summary>
    /// Creates a new HyperlinkWriter with the specified buffer.
    /// </summary>
    public HyperlinkWriter(byte[] buffer)
    {
        _buffer = buffer;
        _memory = buffer;
        _position = 0;
        _isRented = false;
    }

    /// <summary>
    /// Creates a HyperlinkWriter using a pooled buffer.
    /// </summary>
    public static HyperlinkWriter CreatePooled(out byte[] buffer, int minimumLength = DefaultBufferSize)
    {
        buffer = ArrayPool<byte>.Shared.Rent(minimumLength);
        return new HyperlinkWriter(buffer) { _isRented = true };
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
    /// Writes all hyperlinks from a collection.
    /// </summary>
    public int WriteHyperlinks(HyperlinkCollection hyperlinks)
    {
        foreach (var hyperlink in hyperlinks.Hyperlinks)
        {
            WriteHyperlink(hyperlink);
        }

        return _position;
    }

    /// <summary>
    /// Writes a single hyperlink.
    /// </summary>
    public int WriteHyperlink(HyperlinkData hyperlink)
    {
        if (!hyperlink.IsValid())
        {
            return _position;
        }

        // HLINK record (0x01B8)
        // This is a complex record with variable length

        const ushort recordId = 0x01B8;
        var startPosition = _position;

        // Reserve space for record header (will be updated later)
        _position += 4;

        // Row index (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)hyperlink.Row);
        _position += 2;

        // Column index (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)hyperlink.Column);
        _position += 2;

        // Reserved (2 bytes) - must be 0x0002
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0x0002);
        _position += 2;

        // GUID for URL Moniker (16 bytes)
        WriteGuid(HyperlinkGuid.UrlMoniker);

        // Unknown DWORD (4 bytes) - must be 0x00000000
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 4;

        // Option flags (4 bytes)
        var optionFlags = GetOptionFlags(hyperlink);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), optionFlags);
        _position += 4;

        // Write URL/file path
        if (hyperlink.Type == HyperlinkType.Url || hyperlink.Type == HyperlinkType.File)
        {
            WriteUrlOrFile(hyperlink);
        }
        else if (hyperlink.Type == HyperlinkType.Email)
        {
            WriteEmail(hyperlink);
        }
        else if (hyperlink.Type == HyperlinkType.Unc)
        {
            WriteUncPath(hyperlink);
        }
        else if (hyperlink.Type == HyperlinkType.Document || hyperlink.Type == HyperlinkType.Internal)
        {
            WriteDocumentLink(hyperlink);
        }

        // Write description (display text)
        if ((optionFlags & 0x00000014) != 0 && !string.IsNullOrEmpty(hyperlink.DisplayText))
        {
            WriteUnicodeString(hyperlink.DisplayText);
        }

        // Write target frame name
        if ((optionFlags & 0x00000008) != 0)
        {
            WriteUnicodeString("");
        }

        // Write UNC path
        if ((optionFlags & 0x00000010) != 0 && hyperlink.Type == HyperlinkType.Unc)
        {
            WriteUnicodeString(hyperlink.Target);
        }

        // Write bookmark/location
        if ((optionFlags & 0x00000080) != 0 && !string.IsNullOrEmpty(hyperlink.Location))
        {
            WriteUnicodeString(hyperlink.Location);
        }

        // Write Moniker (for file links)
        if ((optionFlags & 0x00000100) != 0 && hyperlink.Type == HyperlinkType.File)
        {
            WriteFileMoniker(hyperlink.Target);
        }

        // Write URL Moniker (for URL links)
        if ((optionFlags & 0x00000300) != 0 && hyperlink.Type == HyperlinkType.Url)
        {
            WriteUrlMoniker(hyperlink.Target);
        }

        // Calculate and write record length
        var recordLength = (ushort)(_position - startPosition - 4);

        // Check if record needs to be continued (BIFF8 max record size is 8224 bytes)
        if (recordLength > 8224)
        {
            // For simplicity, we'll just write the first part
            // In a full implementation, we'd use CONTINUE records
            recordLength = 8224;
        }

        // Write record header
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(startPosition), recordId);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(startPosition + 2), recordLength);

        // If we truncated, adjust position
        if (_position > startPosition + 4 + recordLength)
        {
            _position = startPosition + 4 + recordLength;
        }

        return _position;
    }

    private uint GetOptionFlags(HyperlinkData hyperlink)
    {
        uint flags = 0;

        // Bit 0: Has description (display text)
        if (!string.IsNullOrEmpty(hyperlink.DisplayText))
            flags |= 0x00000014;

        // Bit 1: Has target frame name
        // Not implemented

        // Bit 2: Has UNC path
        if (hyperlink.Type == HyperlinkType.Unc)
            flags |= 0x00000010;

        // Bit 3: Has Moniker or URL Moniker
        if (hyperlink.Type == HyperlinkType.File || hyperlink.Type == HyperlinkType.Url)
            flags |= 0x00000300;

        // Bit 4: Has location string
        if (!string.IsNullOrEmpty(hyperlink.Location))
            flags |= 0x00000080;

        // Bit 8: Absolute path
        if (hyperlink.Type == HyperlinkType.File && Path.IsPathRooted(hyperlink.Target))
            flags |= 0x00000001;

        return flags;
    }

    private void WriteGuid(byte[] guid)
    {
        guid.CopyTo(_buffer.AsSpan(_position));
        _position += 16;
    }

    private void WriteUrlOrFile(HyperlinkData hyperlink)
    {
        // For URL or file links, we need to write the appropriate moniker
        // This is handled by the option flags and moniker writing
    }

    private void WriteEmail(HyperlinkData hyperlink)
    {
        // Email links are essentially URLs with mailto: scheme
        // Write as URL moniker
    }

    private void WriteUncPath(HyperlinkData hyperlink)
    {
        // UNC paths are written as Unicode strings
    }

    private void WriteDocumentLink(HyperlinkData hyperlink)
    {
        // Document links have a location string
    }

    private void WriteUnicodeString(string value)
    {
        // Character count (4 bytes)
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)value.Length);
        _position += 4;

        // Options flags (2 bytes) - 0x0001 = high byte compression (Unicode)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0x0001);
        _position += 2;

        // String data (Unicode, 2 bytes per character)
        var bytes = Encoding.Unicode.GetBytes(value);
        bytes.CopyTo(_buffer.AsSpan(_position));
        _position += bytes.Length;
    }

    private void WriteFileMoniker(string filePath)
    {
        // File Moniker structure
        // This is a simplified implementation

        // Clsid (16 bytes) - FileMoniker class ID
        WriteGuid(HyperlinkGuid.FileMoniker);

        // Stream version (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0x004E); // 'N' for new format
        _position += 2;

        // Reserved (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 2;

        // ANSI path length (4 bytes)
        var ansiPath = Encoding.ASCII.GetBytes(filePath);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)ansiPath.Length);
        _position += 4;

        // ANSI path
        ansiPath.CopyTo(_buffer.AsSpan(_position));
        _position += ansiPath.Length;

        // Unicode path length (4 bytes)
        var unicodePath = Encoding.Unicode.GetBytes(filePath);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)filePath.Length);
        _position += 4;

        // Unicode path
        unicodePath.CopyTo(_buffer.AsSpan(_position));
        _position += unicodePath.Length;
    }

    private void WriteUrlMoniker(string url)
    {
        // URL Moniker structure
        // This is a simplified implementation

        // Clsid (16 bytes) - URLMoniker class ID
        WriteGuid(HyperlinkGuid.UrlMoniker);

        // String length (4 bytes)
        var urlBytes = Encoding.Unicode.GetBytes(url);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)url.Length);
        _position += 4;

        // URL string (Unicode)
        urlBytes.CopyTo(_buffer.AsSpan(_position));
        _position += urlBytes.Length;
    }

    /// <summary>
    /// GUIDs used in hyperlinks.
    /// </summary>
    private static class HyperlinkGuid
    {
        // URL Moniker: {79EAC9E0-BAF9-11CE-8C82-00AA004BA90B}
        public static readonly byte[] UrlMoniker = new byte[]
        {
            0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11,
            0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B
        };

        // File Moniker: {00000303-0000-0000-C000-000000000046}
        public static readonly byte[] FileMoniker = new byte[]
        {
            0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46
        };

        // StdLink: {00000300-0000-0000-C000-000000000046}
        public static readonly byte[] StdLink = new byte[]
        {
            0x00, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46
        };
    }
}

/// <summary>
/// Simplified hyperlink writer that creates basic HLINK records.
/// This is a fallback implementation for when the full writer is too complex.
/// </summary>
public sealed class SimpleHyperlinkWriter : IDisposable
{
    private const int DefaultBufferSize = 65536;

    private byte[] _buffer;
    private Memory<byte> _memory;
    private int _position;
    private bool _isRented;

    public SimpleHyperlinkWriter(byte[] buffer)
    {
        _buffer = buffer;
        _memory = buffer;
        _position = 0;
        _isRented = false;
    }

    public static SimpleHyperlinkWriter CreatePooled(out byte[] buffer, int minimumLength = DefaultBufferSize)
    {
        buffer = ArrayPool<byte>.Shared.Rent(minimumLength);
        return new SimpleHyperlinkWriter(buffer) { _isRented = true };
    }

    public void Dispose()
    {
        if (_isRented)
        {
            ArrayPool<byte>.Shared.Return(_buffer);
            _isRented = false;
        }
    }

    public ReadOnlySpan<byte> GetData() => _memory.Span.Slice(0, _position);

    /// <summary>
    /// Writes a simple URL hyperlink record.
    /// </summary>
    public int WriteUrlHyperlink(int row, int col, string url, string? displayText = null)
    {
        // HLINK record for URL
        const ushort recordId = 0x01B8;
        var startPosition = _position;

        // Calculate record length
        var urlBytes = Encoding.Unicode.GetBytes(url);
        var displayBytes = string.IsNullOrEmpty(displayText)
            ? Array.Empty<byte>()
            : Encoding.Unicode.GetBytes(displayText);

        // Header + row + col + reserved + GUID + DWORD + flags + URL moniker + display
        var recordLength = 4 + 2 + 2 + 2 + 16 + 4 + 4 + (4 + urlBytes.Length) + (displayBytes.Length > 0 ? 6 + displayBytes.Length : 0);

        // Write record header
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)recordLength);
        _position += 2;

        // Row
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)row);
        _position += 2;

        // Column
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)col);
        _position += 2;

        // Reserved
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0x0002);
        _position += 2;

        // URL Moniker GUID
        WriteGuid(new byte[]
        {
            0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11,
            0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B
        });

        // Reserved DWORD
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 4;

        // Option flags (has URL moniker)
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), 0x00000300);
        _position += 4;

        // URL length
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)url.Length);
        _position += 4;

        // URL data
        urlBytes.CopyTo(_buffer.AsSpan(_position));
        _position += urlBytes.Length;

        // Display text (if provided)
        if (displayBytes.Length > 0)
        {
            // Update flags to include description
            var flagsPosition = startPosition + 4 + 2 + 2 + 2 + 16 + 4;
            BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(flagsPosition), 0x00000314);

            // Character count
            BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)displayText!.Length);
            _position += 4;

            // Options (Unicode)
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0x0001);
            _position += 2;

            // Display text data
            displayBytes.CopyTo(_buffer.AsSpan(_position));
            _position += displayBytes.Length;
        }

        return _position;
    }

    private void WriteGuid(byte[] guid)
    {
        guid.CopyTo(_buffer.AsSpan(_position));
        _position += 16;
    }
}
