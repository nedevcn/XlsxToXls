using System.Buffers.Binary;
using System.Text;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Writes OLE Compound File Binary Format (CFB) - zero allocations for performance.
/// Implements [MS-CFB] specification for Excel .xls compatibility.
/// </summary>
internal sealed class OleCompoundWriter
{
    private const int SectorSize = 512;
    private const int DirEntrySize = 132;
    private const int FatSectorEntryCount = SectorSize / 4;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FatSector = 0xFFFFFFFD;

    private readonly List<uint> _fat = new List<uint>();
    private readonly List<byte> _data = new List<byte>();
    private readonly string _streamName;
    private int _streamStartSector = -1;

    public OleCompoundWriter(string streamName = "Workbook")
    {
        _streamName = streamName;
    }

    public void WriteStream(ReadOnlySpan<byte> data)
    {
        _data.Clear();
        _data.AddRange(data.ToArray());
    }

    public void WriteTo(Stream output)
    {
        if (_data.Count == 0)
            throw new InvalidOperationException("No stream data to write.");

        var streamSectors = (int)((_data.Count + SectorSize - 1) / SectorSize);
        var dirSector = 0;
        var fatSector = 1;
        _streamStartSector = 2;

        var totalSectors = 2 + streamSectors;
        var fatSectorCount = (totalSectors + FatSectorEntryCount - 1) / FatSectorEntryCount;
        if (fatSectorCount > 1)
        {
            totalSectors += fatSectorCount;
            _streamStartSector = 2 + fatSectorCount;
        }

        _fat.Clear();
        _fat.Add(EndOfChain);
        _fat.Add(FatSector);
        for (var i = 2; i < 2 + fatSectorCount; i++)
            _fat.Add(FatSector);
        for (var i = 0; i < streamSectors; i++)
            _fat.Add(i < streamSectors - 1 ? (uint)(_streamStartSector + i + 1) : EndOfChain);

        Span<byte> header = stackalloc byte[SectorSize];
        header.Clear();
        header[0] = 0xD0;
        header[1] = 0xCF;
        header[2] = 0x11;
        header[3] = 0xE0;
        header[4] = 0xA1;
        header[5] = 0xB1;
        header[6] = 0x1A;
        header[7] = 0xE1;
        BinaryPrimitives.WriteUInt16LittleEndian(header[0x18..], 0x003E);
        header[0x1A] = 0xFE;
        header[0x1B] = 0xFF;
        BinaryPrimitives.WriteUInt16LittleEndian(header[0x1C..], 9);
        BinaryPrimitives.WriteUInt16LittleEndian(header[0x1E..], 6);
        BinaryPrimitives.WriteInt32LittleEndian(header[0x24..], 0);
        BinaryPrimitives.WriteInt32LittleEndian(header[0x28..], fatSectorCount);
        BinaryPrimitives.WriteInt32LittleEndian(header[0x30..], dirSector);
        BinaryPrimitives.WriteInt32LittleEndian(header[0x38..], 0x1000);
        BinaryPrimitives.WriteInt32LittleEndian(header[0x40..], -1);
        BinaryPrimitives.WriteInt32LittleEndian(header[0x44..], -1);
        BinaryPrimitives.WriteInt32LittleEndian(header[0x48..], -1);
        BinaryPrimitives.WriteInt32LittleEndian(header[0x4C..], -1);

        for (var i = 0; i < Math.Min(109, fatSectorCount); i++)
            BinaryPrimitives.WriteInt32LittleEndian(header[(0x4C + i * 4)..], fatSector + i);
        for (var i = fatSectorCount; i < 109; i++)
            BinaryPrimitives.WriteInt32LittleEndian(header[(0x4C + i * 4)..], -1);

        output.Write(header);

        // Sector 0 = directory (ofs_dir)
        // Sector 1.. = FAT, then stream data
        Span<byte> dirEntry = stackalloc byte[DirEntrySize];
        dirEntry.Clear();

        // Root entry
        var rootName = Encoding.Unicode.GetBytes("\u0005Root Entry");
        rootName.AsSpan(0, Math.Min(rootName.Length, 63)).CopyTo(dirEntry);
        dirEntry[64] = (byte)(rootName.Length / 2);
        dirEntry[65] = (byte)((rootName.Length / 2) >> 8);
        dirEntry[66] = 5;  // root storage
        dirEntry[67] = 1;  // black
        BinaryPrimitives.WriteInt32LittleEndian(dirEntry[68..], -1);
        BinaryPrimitives.WriteInt32LittleEndian(dirEntry[72..], -1);
        BinaryPrimitives.WriteInt32LittleEndian(dirEntry[76..], 1);
        BinaryPrimitives.WriteInt32LittleEndian(dirEntry[120..], 0);
        BinaryPrimitives.WriteInt32LittleEndian(dirEntry[124..], 0);  // Low 32 bits
        BinaryPrimitives.WriteInt32LittleEndian(dirEntry[128..], 0);  // High 32 bits
        output.Write(dirEntry);

        // Workbook stream entry
        dirEntry.Clear();
        var wbName = Encoding.Unicode.GetBytes(_streamName);
        wbName.AsSpan(0, Math.Min(wbName.Length, 63)).CopyTo(dirEntry);
        dirEntry[64] = (byte)(wbName.Length / 2);
        dirEntry[65] = (byte)((wbName.Length / 2) >> 8);
        dirEntry[66] = 2;  // stream
        dirEntry[67] = 1;
        BinaryPrimitives.WriteInt32LittleEndian(dirEntry[68..], -1);
        BinaryPrimitives.WriteInt32LittleEndian(dirEntry[72..], -1);
        BinaryPrimitives.WriteInt32LittleEndian(dirEntry[76..], -1);
        BinaryPrimitives.WriteInt32LittleEndian(dirEntry[120..], _streamStartSector);
        BinaryPrimitives.WriteInt32LittleEndian(dirEntry[124..], (int)_data.Count);  // Low 32 bits
        BinaryPrimitives.WriteInt32LittleEndian(dirEntry[128..], 0);  // High 32 bits
        output.Write(dirEntry);

        var dirPadding = SectorSize - DirEntrySize * 2;
        if (dirPadding > 0)
        {
            Span<byte> pad = stackalloc byte[dirPadding];
            pad.Fill(0);
            output.Write(pad);
        }

        // Write FAT sectors
        var fatBytes = _fat.Count * 4;
        var fatSectorBytes = ((fatBytes + SectorSize - 1) / SectorSize) * SectorSize;
        var fatBuffer = new byte[fatSectorBytes];
        for (var i = 0; i < _fat.Count; i++)
            BinaryPrimitives.WriteUInt32LittleEndian(fatBuffer.AsSpan(i * 4), _fat[i]);
        output.Write(fatBuffer);

        // Write stream data sectors
        var offset = 0;
        var remaining = _data.Count;
        var sectorBuffer = new byte[SectorSize];
        while (remaining > 0)
        {
            var toCopy = Math.Min(SectorSize, remaining);
            _data.CopyTo(offset, sectorBuffer, 0, toCopy);
            if (toCopy < SectorSize)
                Array.Clear(sectorBuffer, toCopy, SectorSize - toCopy);
            output.Write(sectorBuffer);
            offset += toCopy;
            remaining -= toCopy;
        }
    }

}
