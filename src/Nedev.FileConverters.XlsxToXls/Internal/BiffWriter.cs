using System.Buffers.Binary;
using System.Text;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Writes BIFF8 records for Excel .xls format. High-performance, minimal allocations.
/// </summary>
internal ref struct BiffWriter
{
    private const int BiffMaxRecordData = 8224;
    private Span<byte> _buffer;
    private int _position;

    public BiffWriter(Span<byte> buffer)
    {
        _buffer = buffer;
        _position = 0;
    }

    public int Position => _position;

    public void WriteBofWorkbook()
    {
        WriteRecordHeader(0x0809, 16);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0600);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0005);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0C0A);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x07CC);
        _position += 2;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0x00000001);
        _position += 4;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0006);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0000);
        _position += 2;
    }

    public void WriteBofWorksheet()
    {
        WriteRecordHeader(0x0809, 16);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0600);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0010);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0C0A);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x07CC);
        _position += 2;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0x00000001);
        _position += 4;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0006);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0000);
        _position += 2;
    }

    public void WriteCodepage(ushort codepage = 0x04E4)
    {
        WriteRecordHeader(0x0042, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), codepage);
        _position += 2;
    }

    public void WriteBoundSheet(int streamPosition, string name, byte sheetType = 0)
    {
        var nameBytes = Encoding.GetEncoding(1252).GetBytes(name);
        var len = Math.Min(nameBytes.Length, 31);
        var recLen = 4 + 1 + 1 + 1 + len;
        WriteRecordHeader(0x0085, recLen);
        BinaryPrimitives.WriteInt32LittleEndian(_buffer.Slice(_position), streamPosition);
        _position += 4;
        _buffer[_position++] = (byte)len;
        _buffer[_position++] = sheetType;
        _buffer[_position++] = 0;
        nameBytes.AsSpan(0, len).CopyTo(_buffer.Slice(_position));
        _position += len;
    }

    public void WriteSupBookInternalRef(int sheetCount)
    {
        WriteRecordHeader(0x01AE, 4);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)sheetCount);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0401);
        _position += 2;
    }

    public void WriteExternSheet(int sheetCount)
    {
        var recLen = 2 + sheetCount * 6;
        WriteRecordHeader(0x0017, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)sheetCount);
        _position += 2;
        for (var i = 0; i < sheetCount; i++)
        {
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)i);
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)i);
            _position += 2;
        }
    }

    public void WriteNameBuiltin(DefinedNameInfo dn, ushort externSheetIndex)
    {
        const int formulaLen = 11;
        var recLen = 2 + 1 + 1 + 2 + 2 + 2 + 4 + 2 + formulaLen;
        WriteRecordHeader(0x0018, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0020);
        _position += 2;
        _buffer[_position++] = 0;
        _buffer[_position++] = 1;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), formulaLen);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(dn.SheetIndex0Based + 1));
        _position += 2;
        _position += 4;
        _buffer[_position++] = 0;
        _buffer[_position++] = dn.BuiltinIndex;
        _buffer[_position++] = 0x3B;
        _position += 1;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), externSheetIndex);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)dn.FirstRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)dn.LastRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)dn.FirstCol);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)dn.LastCol);
        _position += 2;
    }

    public void WriteFont(string name, double heightTwips, bool bold, bool italic)
    {
        var nameBytes = Encoding.GetEncoding(1252).GetBytes(name);
        var len = Math.Min(nameBytes.Length, 255);
        var recLen = 14 + 1 + len;
        WriteRecordHeader(0x0031, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)heightTwips);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)((italic ? 2 : 0)));
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x7FFF);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(bold ? 0x02BC : 0x0190));
        _position += 2;
        _buffer[_position++] = 0;
        _buffer[_position++] = 0;
        _buffer[_position++] = 0;
        _buffer[_position++] = 0;
        _buffer[_position++] = 0;
        _buffer[_position++] = (byte)len;
        nameBytes.AsSpan(0, len).CopyTo(_buffer.Slice(_position));
        _position += len;
    }

    public void WriteFormat(int index, string formatCode)
    {
        var code = Encoding.Unicode.GetBytes(formatCode);
        var recLen = 2 + 2 + code.Length;
        WriteRecordHeader(0x041E, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)index);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(code.Length / 2));
        _position += 2;
        code.AsSpan().CopyTo(_buffer.Slice(_position));
        _position += code.Length;
    }

    public void WriteBuiltinFmtCount(int count)
    {
        WriteRecordHeader(0x001F, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)count);
        _position += 2;
    }

    public void WriteXf(int fontIdx, int formatIdx, bool isCellXf = true, CellXfInfo? xfInfo = null)
    {
        WriteRecordHeader(0x00E0, 20);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)fontIdx);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)formatIdx);
        _position += 2;
        // Type / protection / parent XF index.
        // 保持原有高位行为（0xFFF0），只在低两位编码 Locked/Hidden。
        ushort xfTypeProt = 0xFFF0;
        if (xfInfo.HasValue && isCellXf)
        {
            if (xfInfo.Value.Locked)
                xfTypeProt |= 0x0001;
            if (xfInfo.Value.Hidden)
                xfTypeProt |= 0x0002;
        }
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), xfTypeProt);
        _position += 2;

        // Alignment: bits 0-2 horizontal, bit 3 wrap, bits 4-6 vertical.
        byte align = 0;
        if (xfInfo.HasValue)
        {
            align |= (byte)(xfInfo.Value.HorizontalAlign & 0x07);
            if (xfInfo.Value.WrapText)
                align |= 0x08;
            align |= (byte)((xfInfo.Value.VerticalAlign & 0x07) << 4);
        }
        _buffer[_position++] = align;

        // Text attributes: low 4 bits indent level (0-15). 其余保持 0。
        byte textAttrs = 0;
        if (xfInfo.HasValue)
            textAttrs |= (byte)(xfInfo.Value.Indent & 0x0F);
        _buffer[_position++] = textAttrs;

        // Rotation / other flags 暂时保持 0。
        _buffer[_position++] = 0;
        _buffer[_position++] = 0;

        // Border / background / pattern 等暂时维持为 0。
        _position += 12;
    }

    public void WriteEof()
    {
        WriteRecordHeader(0x000A, 0);
    }

    public void WriteWsBool()
    {
        WriteRecordHeader(0x0081, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0401);
        _position += 2;
    }

    public void WriteDefColWidth(ushort widthChars = 8)
    {
        WriteRecordHeader(0x0055, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), widthChars);
        _position += 2;
    }

    public void WriteDimension(int firstRow, int lastRowPlus1, int firstCol, int lastColPlus1)
    {
        WriteRecordHeader(0x0200, 14);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)Math.Max(0, firstRow));
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)Math.Max(0, lastRowPlus1));
        _position += 4;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)Math.Clamp(firstCol, 0, 255));
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)Math.Clamp(lastColPlus1, 0, 256));
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    public void WriteColInfo(int firstCol, int lastCol, double widthChars, bool hidden)
    {
        WriteRecordHeader(0x007D, 12);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)firstCol);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)lastCol);
        _position += 2;
        var w = (ushort)Math.Clamp((int)(widthChars * 256), 0, 65535);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), w);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 15);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(hidden ? 1 : 0));
        _position += 2;
        _position += 2;
    }

    public void WriteRow(int rowIndex, double heightPoints, bool hidden)
    {
        WriteRecordHeader(0x0208, 16);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)rowIndex);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 255);
        _position += 2;
        var ht = heightPoints > 0 ? (ushort)(heightPoints * 20) : (ushort)0;
        var htWord = (ushort)(ht & 0x7FFF);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), htWord);
        _position += 2;
        _position += 4;
        var flags = 0x100u;
        if (hidden) flags |= 0x20;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;
    }

    public void WriteMergeCells(ReadOnlySpan<(int FirstRow, int FirstCol, int LastRow, int LastCol)> ranges)
    {
        if (ranges.Length == 0) return;
        var recLen = 2 + ranges.Length * 8;
        WriteRecordHeader(0x00E5, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)ranges.Length);
        _position += 2;
        for (var i = 0; i < ranges.Length; i++)
        {
            var r = ranges[i];
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)r.FirstRow);
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)r.LastRow);
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)r.FirstCol);
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)r.LastCol);
            _position += 2;
        }
    }

    public void WriteHorizontalPageBreaks(ReadOnlySpan<(int Row, int FirstCol, int LastCol)> breaks)
    {
        if (breaks.Length == 0) return;
        var recLen = 2 + breaks.Length * 6;
        WriteRecordHeader(0x001B, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)breaks.Length);
        _position += 2;
        for (var i = 0; i < breaks.Length; i++)
        {
            var b = breaks[i];
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)b.Row);
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)b.FirstCol);
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)b.LastCol);
            _position += 2;
        }
    }

    public void WriteVerticalPageBreaks(ReadOnlySpan<(int Col, int FirstRow, int LastRow)> breaks)
    {
        if (breaks.Length == 0) return;
        var recLen = 2 + breaks.Length * 6;
        WriteRecordHeader(0x001A, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)breaks.Length);
        _position += 2;
        for (var i = 0; i < breaks.Length; i++)
        {
            var b = breaks[i];
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)b.Col);
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)b.FirstRow);
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)b.LastRow);
            _position += 2;
        }
    }

    public void WriteLeftMargin(double inches)
    {
        WriteRecordHeader(0x0026, 8);
        BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), inches);
        _position += 8;
    }

    public void WriteRightMargin(double inches)
    {
        WriteRecordHeader(0x0027, 8);
        BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), inches);
        _position += 8;
    }

    public void WriteTopMargin(double inches)
    {
        WriteRecordHeader(0x0028, 8);
        BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), inches);
        _position += 8;
    }

    public void WriteBottomMargin(double inches)
    {
        WriteRecordHeader(0x0029, 8);
        BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), inches);
        _position += 8;
    }

    public void WritePrintGridLines(bool enabled)
    {
        WriteRecordHeader(0x002B, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(enabled ? 1 : 0));
        _position += 2;
    }

    public void WritePrintHeaders(bool enabled)
    {
        WriteRecordHeader(0x002A, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(enabled ? 1 : 0));
        _position += 2;
    }

    public void WriteCenterHorizontal(bool enabled)
    {
        WriteRecordHeader(0x0083, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(enabled ? 1 : 0));
        _position += 2;
    }

    public void WriteCenterVertical(bool enabled)
    {
        WriteRecordHeader(0x0084, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(enabled ? 1 : 0));
        _position += 2;
    }

    public void WritePageSetup(bool landscape, int scale, int startPage, int fitToWidth, int fitToHeight, double headerMargin, double footerMargin)
    {
        const int recLen = 34;
        WriteRecordHeader(0x00A1, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 1);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)Math.Clamp(scale, 0, 65535));
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)Math.Clamp(startPage, 0, 65535));
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)Math.Clamp(fitToWidth, 0, 65535));
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)Math.Clamp(fitToHeight, 0, 65535));
        _position += 2;
        ushort flags = 0x0004;
        if (landscape) flags |= 0x0002;
        flags |= 0x0080;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), flags);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 600);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 600);
        _position += 2;
        BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), headerMargin);
        _position += 8;
        BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), footerMargin);
        _position += 8;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 1);
        _position += 2;
    }

    public void WriteWindow2(bool freezePanes = false)
    {
        WriteRecordHeader(0x003E, 18);
        var flags = (ushort)0x06B6;
        if (freezePanes) flags |= 0x0008;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), flags);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        _position += 12;
    }

    public void WriteHeader(ReadOnlySpan<char> text)
    {
        if (text.Length == 0) return;
        var need16 = HasHighChar(text);
        var byteCount = need16 ? text.Length * 2 : text.Length;
        WriteRecordHeader(0x0014, 3 + byteCount);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)text.Length);
        _position += 2;
        _buffer[_position++] = (byte)(need16 ? 1 : 0);
        if (need16)
        {
            for (var i = 0; i < text.Length; i++)
                BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position + i * 2), text[i]);
            _position += byteCount;
        }
        else
        {
            for (var i = 0; i < text.Length; i++)
                _buffer[_position + i] = (byte)text[i];
            _position += byteCount;
        }
    }

    public void WriteFooter(ReadOnlySpan<char> text)
    {
        if (text.Length == 0) return;
        var need16 = HasHighChar(text);
        var byteCount = need16 ? text.Length * 2 : text.Length;
        WriteRecordHeader(0x0015, 3 + byteCount);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)text.Length);
        _position += 2;
        _buffer[_position++] = (byte)(need16 ? 1 : 0);
        if (need16)
        {
            for (var i = 0; i < text.Length; i++)
                BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position + i * 2), text[i]);
            _position += byteCount;
        }
        else
        {
            for (var i = 0; i < text.Length; i++)
                _buffer[_position + i] = (byte)text[i];
            _position += byteCount;
        }
    }

    public void WritePane(ushort px, ushort py, ushort topRowVisible, ushort leftColVisible, byte activePane)
    {
        WriteRecordHeader(0x0041, 9);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), px);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), py);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), topRowVisible);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), leftColVisible);
        _position += 2;
        _buffer[_position++] = Math.Clamp(activePane, (byte)0, (byte)3);
    }

    public void WriteObjNote(ushort objectId)
    {
        const int totalLen = 60;
        WriteRecordHeader(0x005D, totalLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x15);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 22);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x19);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), objectId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0002);
        _position += 2;
        _position += 16;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0D);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 26);
        _position += 2;
        _position += 26;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    public void WriteTxoWithText(ReadOnlySpan<char> text)
    {
        var need16 = false;
        foreach (var c in text) { if (c > 255) { need16 = true; break; } }
        var maxCch = need16 ? 4111 : 8222;
        var cch = Math.Min(text.Length, maxCch);
        var textSpan = text.Slice(0, cch);
        WriteRecordHeader(0x01B6, 18);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)cch);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 16);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        var textBytes = 1 + (need16 ? cch * 2 : cch);
        WriteRecordHeader(0x003C, textBytes);
        _buffer[_position++] = (byte)(need16 ? 1 : 0);
        if (need16)
            for (var i = 0; i < textSpan.Length; i++)
                BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position + i * 2), textSpan[i]);
        else
            for (var i = 0; i < textSpan.Length; i++)
                _buffer[_position + i] = (byte)textSpan[i];
        _position += need16 ? textSpan.Length * 2 : textSpan.Length;
        WriteRecordHeader(0x003C, 16);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)cch);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        _position += 8;
    }

    public void WriteNote(int row, int col, bool visible, ushort shapeId, ReadOnlySpan<char> author)
    {
        var needs16 = false;
        foreach (var c in author)
        {
            if (c > 255) { needs16 = true; break; }
        }
        var charCount = author.Length;
        var recLen = 8 + 1 + 2 + (needs16 ? charCount * 2 : charCount);
        WriteRecordHeader(0x001C, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)row);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)col);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(visible ? 2 : 0));
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), shapeId);
        _position += 2;
        _buffer[_position++] = (byte)(needs16 ? 1 : 0);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)charCount);
        _position += 2;
        if (needs16)
            for (var i = 0; i < author.Length; i++)
                BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position + i * 2), author[i]);
        else
            for (var i = 0; i < author.Length; i++)
                _buffer[_position + i] = (byte)author[i];
        _position += needs16 ? author.Length * 2 : author.Length;
    }

    private static readonly byte[] StdLinkGuid = { 0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B };
    private static readonly byte[] UrlMonikerGuid = { 0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B };

    public void WriteHyperlink(int firstRow, int firstCol, int lastRow, int lastCol, ReadOnlySpan<char> url)
    {
        var urlLen = url.Length + 1;
        var urlBytes = urlLen * 2;
        var recLen = 8 + 16 + 4 + 4 + 16 + 4 + urlBytes;
        WriteRecordHeader(0x01B8, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)firstRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)lastRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)firstCol);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)lastCol);
        _position += 2;
        StdLinkGuid.CopyTo(_buffer.Slice(_position));
        _position += 16;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0x00000002u);
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0x00000003u);
        _position += 4;
        UrlMonikerGuid.CopyTo(_buffer.Slice(_position));
        _position += 16;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)urlBytes);
        _position += 4;
        for (var i = 0; i < url.Length; i++)
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position + i * 2), url[i]);
        _position += url.Length * 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    public void WriteBlank(int row, int col, ushort xfIndex = 15)
    {
        WriteRecordHeader(0x0201, 6);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)row);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)col);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), xfIndex);
        _position += 2;
    }

    public void WriteBool(int row, int col, bool value, ushort xfIndex = 15)
    {
        WriteRecordHeader(0x0205, 8);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)row);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)col);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), xfIndex);
        _position += 2;
        _buffer[_position++] = (byte)(value ? 1 : 0);
        _buffer[_position++] = 0;
    }

    public void WriteError(int row, int col, byte errorCode, ushort xfIndex = 15)
    {
        WriteRecordHeader(0x0205, 8);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)row);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)col);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), xfIndex);
        _position += 2;
        _buffer[_position++] = Math.Clamp(errorCode, (byte)0, (byte)42);
        _buffer[_position++] = 1;
    }

    public void WriteNumber(int row, int col, double value, ushort xfIndex = 15)
    {
        WriteRecordHeader(0x0203, 14);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)row);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)col);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), xfIndex);
        _position += 2;
        BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), value);
        _position += 8;
    }

    public void WriteFormula(int row, int col, ushort xfIndex, ReadOnlySpan<byte> rgce, CellKind cachedKind, double cachedNumber, bool cachedBool, byte cachedError, ReadOnlySpan<char> cachedString)
    {
        var recLen = 22 + rgce.Length;
        WriteRecordHeader(0x0006, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)row);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)col);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), xfIndex);
        _position += 2;

        // FormulaValue (8 bytes)
        if (cachedKind == CellKind.Number)
        {
            BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), cachedNumber);
        }
        else
        {
            _buffer.Slice(_position, 8).Clear();
            _buffer[_position] = cachedKind switch
            {
                CellKind.String or CellKind.SharedString => (byte)0x00,
                CellKind.Boolean => (byte)0x01,
                CellKind.Error => (byte)0x02,
                _ => (byte)0x03
            };
            if (cachedKind == CellKind.Boolean)
                _buffer[_position + 2] = (byte)(cachedBool ? 1 : 0);
            else if (cachedKind == CellKind.Error)
                _buffer[_position + 2] = cachedError;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position + 6), 0xFFFF);
        }
        _position += 8;

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)rgce.Length);
        _position += 2;
        rgce.CopyTo(_buffer.Slice(_position));
        _position += rgce.Length;

        if (cachedKind is CellKind.String or CellKind.SharedString)
            WriteString(cachedString);
    }

    public void WriteString(ReadOnlySpan<char> text)
    {
        if (text.Length > 32767) text = text[..32767];
        var need16 = HasHighChar(text);
        var maxChars = need16 ? (BiffMaxRecordData - 3) / 2 : (BiffMaxRecordData - 3);
        if (text.Length > maxChars) text = text[..maxChars];
        var recLen = 3 + (need16 ? text.Length * 2 : text.Length);
        WriteRecordHeader(0x0207, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)text.Length);
        _position += 2;
        _buffer[_position++] = (byte)(need16 ? 1 : 0);
        if (need16)
        {
            for (var i = 0; i < text.Length; i++)
                BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position + i * 2), text[i]);
            _position += text.Length * 2;
        }
        else
        {
            for (var i = 0; i < text.Length; i++)
                _buffer[_position + i] = (byte)text[i];
            _position += text.Length;
        }
    }

    public void WriteLabel(int row, int col, ReadOnlySpan<char> text, ushort xfIndex = 15)
    {
        var needs16Bit = false;
        foreach (var c in text)
        {
            if (c > 255) { needs16Bit = true; break; }
        }
        var charCount = text.Length;
        var byteCount = needs16Bit ? charCount * 2 : charCount;
        var recLen = 6 + 1 + 2 + byteCount;
        WriteRecordHeader(0x0204, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)row);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)col);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), xfIndex);
        _position += 2;
        _buffer[_position++] = (byte)(needs16Bit ? 0 : 1);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)charCount);
        _position += 2;
        if (needs16Bit)
        {
            for (var i = 0; i < text.Length; i++)
                BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position + i * 2), text[i]);
            _position += byteCount;
        }
        else
        {
            for (var i = 0; i < text.Length; i++)
                _buffer[_position++] = (byte)text[i];
        }
    }

    public void WriteLabelSst(int row, int col, int sstIndex, ushort xfIndex = 15)
    {
        WriteRecordHeader(0x00FD, 8);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)row);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)col);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), xfIndex);
        _position += 2;
        BinaryPrimitives.WriteInt32LittleEndian(_buffer.Slice(_position), sstIndex);
        _position += 4;
    }

    public void WriteDatavalidations(int count, bool showPrompt = true)
    {
        if (count <= 0) return;
        WriteRecordHeader(0x01B2, 18);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(showPrompt ? 0x0004 : 0));
        _position += 2;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0xFFFFFFFF);
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)count);
        _position += 4;
    }

    public void WriteDatavalidation(DataValidationInfo dv, ReadOnlySpan<byte> f1, ReadOnlySpan<byte> f2)
    {
        var lenPromptTitle = 2 + 1 + (HasHighChar(dv.PromptTitle) ? dv.PromptTitle.Length * 2 : dv.PromptTitle.Length);
        var lenErrorTitle = 2 + 1 + (HasHighChar(dv.ErrorTitle) ? dv.ErrorTitle.Length * 2 : dv.ErrorTitle.Length);
        var lenPromptText = 2 + 1 + (HasHighChar(dv.PromptText) ? dv.PromptText.Length * 2 : dv.PromptText.Length);
        var lenErrorText = 2 + 1 + (HasHighChar(dv.ErrorText) ? dv.ErrorText.Length * 2 : dv.ErrorText.Length);
        var recLen = 4 + lenPromptTitle + lenErrorTitle + lenPromptText + lenErrorText + 4 + f1.Length + 4 + f2.Length + 2 + dv.Ranges.Count * 8;
        WriteRecordHeader(0x01BE, recLen);
        var flags = (uint)(dv.Type & 0x0F) | ((uint)(dv.Operator & 0x0F) << 20) | (dv.AllowBlank ? 0x100u : 0) | (dv.Type == 3 && dv.Formula1.IndexOf(',') >= 0 ? 0x80u : 0) | (dv.ShowPrompt ? 0x40000u : 0) | (dv.ShowError ? 0x80000u : 0);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;
        WriteBiff8UnicodeString16(dv.PromptTitle);
        WriteBiff8UnicodeString16(dv.ErrorTitle);
        WriteBiff8UnicodeString16(dv.PromptText);
        WriteBiff8UnicodeString16(dv.ErrorText);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)f1.Length);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        f1.CopyTo(_buffer.Slice(_position));
        _position += f1.Length;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)f2.Length);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
        f2.CopyTo(_buffer.Slice(_position));
        _position += f2.Length;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)dv.Ranges.Count);
        _position += 2;
        foreach (var (fr, fc, lr, lc) in dv.Ranges)
        {
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)fr);
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position + 2), (ushort)lr);
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position + 4), (ushort)fc);
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position + 6), (ushort)lc);
            _position += 8;
        }
    }

    private static bool HasHighChar(ReadOnlySpan<char> s)
    {
        foreach (var c in s) if (c > 255) return true;
        return false;
    }

    private void WriteBiff8UnicodeString16(ReadOnlySpan<char> s)
    {
        var need16 = false;
        foreach (var c in s) { if (c > 255) { need16 = true; break; } }
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)s.Length);
        _position += 2;
        _buffer[_position++] = (byte)(need16 ? 1 : 0);
        if (need16)
            for (var i = 0; i < s.Length; i++)
                BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position + i * 2), s[i]);
        else
            for (var i = 0; i < s.Length; i++)
                _buffer[_position + i] = (byte)s[i];
        _position += need16 ? s.Length * 2 : s.Length;
    }

    private void WriteRecordHeader(ushort type, int length)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), type);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)length);
        _position += 2;
    }

    // 图表相关记录
    public void WriteObjChart(ushort objectId, ChartData chart)
    {
        // OBJ记录 - 图表对象
        WriteRecordHeader(0x005D, 38);

        // 对象类型 (0x0005 = 图表)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0005);
        _position += 2;

        // 对象ID
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), objectId);
        _position += 2;

        // 选项标志
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x6011);
        _position += 2;

        // 锁定标志
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0000);
        _position += 2;

        // 预留
        _position += 4;

        // 锚点信息 (以1/4000为单位)
        var x = (uint)(chart.Position.X * 4000 / 100);
        var y = (uint)(chart.Position.Y * 4000 / 100);
        var width = (uint)(chart.Position.Width * 4000 / 100);
        var height = (uint)(chart.Position.Height * 4000 / 100);

        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), x);
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), y);
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), width);
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), height);
        _position += 4;

        // 预留
        _position += 8;
    }

    public void WriteMsodrawingChart(ushort objectId, ReadOnlySpan<byte> chartData)
    {
        // MSODRAWING记录 - Office绘图对象
        const ushort msodrawingType = 0x00EC;

        // 写入Escher容器头部
        var headerSize = 8;
        var totalSize = headerSize + chartData.Length;

        if (totalSize <= BiffMaxRecordData)
        {
            WriteRecordHeader(msodrawingType, totalSize);

            // EscherSpgrContainer记录头
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0F00); // 版本和实例
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0000); // 类型 (SpgrContainer)
            _position += 2;
            BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)(8 + chartData.Length)); // 大小
            _position += 4;

            chartData.CopyTo(_buffer.Slice(_position));
            _position += chartData.Length;
        }
        else
        {
            // 需要分块写入
            var offset = 0;
            var remaining = totalSize;
            var isFirst = true;

            while (remaining > 0)
            {
                var chunkSize = Math.Min(BiffMaxRecordData, remaining);
                WriteRecordHeader(msodrawingType, chunkSize);

                if (isFirst)
                {
                    // EscherSpgrContainer记录头
                    BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0F00);
                    _position += 2;
                    BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0000);
                    _position += 2;
                    BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)(8 + chartData.Length));
                    _position += 4;
                    isFirst = false;
                }

                var dataToCopy = Math.Min(chunkSize - (isFirst ? 0 : 0), chartData.Length - offset);
                if (dataToCopy > 0)
                {
                    chartData.Slice(offset, dataToCopy).CopyTo(_buffer.Slice(_position));
                    _position += dataToCopy;
                    offset += dataToCopy;
                }

                remaining -= chunkSize;
            }
        }
    }

    public void WriteBofChart()
    {
        WriteRecordHeader(0x0809, 16);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0600);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0020); // 图表流
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0C0A);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x07CC);
        _position += 2;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0x00000001);
        _position += 4;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0006);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0000);
        _position += 2;
    }

    public void WriteMsodrawingShapes(ReadOnlySpan<byte> shapeData)
    {
        // MSODRAWING记录 - Office绘图对象用于形状
        const ushort msodrawingType = 0x00EC;

        if (shapeData.Length <= BiffMaxRecordData)
        {
            WriteRecordHeader(msodrawingType, shapeData.Length);
            shapeData.CopyTo(_buffer.Slice(_position));
            _position += shapeData.Length;
        }
        else
        {
            // 需要分块写入
            var offset = 0;
            var remaining = shapeData.Length;

            while (remaining > 0)
            {
                var chunkSize = Math.Min(BiffMaxRecordData, remaining);
                WriteRecordHeader(msodrawingType, chunkSize);

                shapeData.Slice(offset, chunkSize).CopyTo(_buffer.Slice(_position));
                _position += chunkSize;
                offset += chunkSize;
                remaining -= chunkSize;
            }
        }
    }
}
