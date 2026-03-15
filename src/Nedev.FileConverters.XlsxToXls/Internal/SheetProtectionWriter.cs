using System.Buffers.Binary;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Writes sheet protection records in BIFF8 format.
/// </summary>
internal ref struct SheetProtectionWriter
{
    private readonly Span<byte> _buffer;
    private int _position;

    public SheetProtectionWriter(Span<byte> buffer)
    {
        _buffer = buffer;
        _position = 0;
    }

    public int Position => _position;

    /// <summary>
    /// Writes the sheet protection record (PROTECT).
    /// </summary>
    public void WriteSheetProtection(SheetProtectionInfo protection)
    {
        // PROTECT record (0x0012)
        // 2 bytes: fLocked (1 = protected, 0 = not protected)
        WriteRecordHeader(0x0012, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(protection.IsProtected ? 1 : 0));
        _position += 2;
    }

    /// <summary>
    /// Writes the password hash record (SCENPROTECT).
    /// </summary>
    public void WritePasswordHash(ushort passwordHash)
    {
        // PASSWORD record (0x0013)
        // 2 bytes: wPassword (password hash)
        WriteRecordHeader(0x0013, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), passwordHash);
        _position += 2;
    }

    /// <summary>
    /// Writes the protection options record (PROTECTION).
    /// </summary>
    public void WriteProtectionOptions(SheetProtectionInfo protection)
    {
        // PROTECTION record (0x0867) - BIFF8
        // 2 bytes: fLocked (cells locked)
        // 2 bytes: fHidden (cells hidden)
        WriteRecordHeader(0x0867, 4);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 1); // fLocked
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0); // fHidden
        _position += 2;
    }

    /// <summary>
    /// Writes the cell protection flags (DEFAULTROWHEIGHT includes protection info).
    /// </summary>
    public void WriteDefaultProtectionSettings(SheetProtectionInfo protection)
    {
        // DEFAULTROWHEIGHT record (0x0225) with protection flags
        // This record includes default row height and protection settings
        WriteRecordHeader(0x0225, 4);

        // Flags: bit 0 = fUnsynced, bit 1 = fDyZero, bit 2 = fExAsc, bit 3 = fExDsc
        // bit 4-15 = reserved
        var flags = (ushort)0;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), flags);
        _position += 2;

        // Default row height in twips (1/20 of a point)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 255);
        _position += 2;
    }

    /// <summary>
    /// Writes the scenario protection record (SCENPROTECT).
    /// </summary>
    public void WriteScenarioProtection(bool allowEditScenarios)
    {
        // SCENPROTECT record (0x00DD)
        // 2 bytes: fScenProtect (1 = protected, 0 = not protected)
        WriteRecordHeader(0x00DD, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(allowEditScenarios ? 0 : 1));
        _position += 2;
    }

    /// <summary>
    /// Writes the object protection record (OBJPROTECT).
    /// </summary>
    public void WriteObjectProtection(bool allowEditObjects)
    {
        // OBJPROTECT record (0x0063)
        // 2 bytes: fObjectProtect (1 = protected, 0 = not protected)
        WriteRecordHeader(0x0063, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(allowEditObjects ? 0 : 1));
        _position += 2;
    }

    /// <summary>
    /// Writes the FEATHEADR record for enhanced protection (BIFF8).
    /// </summary>
    public void WriteFeatureHeader(SheetProtectionInfo protection)
    {
        // FEATHEADR record (0x0868) - Feature Header
        // This record is used for advanced protection features in BIFF8
        var dataSize = 16; // Minimum size for feature header

        WriteRecordHeader(0x0868, dataSize);

        // Feature type (2 bytes) - 0x0002 for protection
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0002);
        _position += 2;

        // Feature identifier (1 byte)
        _buffer[_position++] = 0x01;

        // Reserved (1 byte)
        _buffer[_position++] = 0x00;

        // cbHdrData (4 bytes) - size of header data
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 8);
        _position += 4;

        // Protection flags (8 bytes)
        var flags = 0u;
        if (!protection.AllowFormatCells) flags |= 0x00000001;
        if (!protection.AllowFormatColumns) flags |= 0x00000002;
        if (!protection.AllowFormatRows) flags |= 0x00000004;
        if (!protection.AllowInsertColumns) flags |= 0x00000008;
        if (!protection.AllowInsertRows) flags |= 0x00000010;
        if (!protection.AllowInsertHyperlinks) flags |= 0x00000020;
        if (!protection.AllowDeleteColumns) flags |= 0x00000040;
        if (!protection.AllowDeleteRows) flags |= 0x00000080;
        if (!protection.AllowSort) flags |= 0x00000100;
        if (!protection.AllowAutoFilter) flags |= 0x00000200;
        if (!protection.AllowPivotTables) flags |= 0x00000400;
        if (!protection.AllowSelectLockedCells) flags |= 0x00000800;
        if (!protection.AllowSelectUnlockedCells) flags |= 0x00001000;

        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // Reserved (4 bytes)
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;
    }

    /// <summary>
    /// Writes all sheet protection records.
    /// </summary>
    public int WriteAllProtectionRecords(SheetProtectionInfo? protection)
    {
        if (protection == null || !protection.Value.IsProtected)
            return _position;

        var prot = protection.Value;

        // Write PROTECT record
        WriteSheetProtection(prot);

        // Write PASSWORD record if password is set
        if (prot.PasswordHash != 0)
        {
            WritePasswordHash(prot.PasswordHash);
        }

        // Write scenario protection
        WriteScenarioProtection(prot.AllowEditScenarios);

        // Write object protection
        WriteObjectProtection(prot.AllowEditObjects);

        // Write feature header for advanced protection
        WriteFeatureHeader(prot);

        return _position;
    }

    private void WriteRecordHeader(ushort type, int length)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), type);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)length);
        _position += 2;
    }
}

/// <summary>
/// Writes workbook protection records in BIFF8 format.
/// </summary>
internal ref struct WorkbookProtectionWriter
{
    private readonly Span<byte> _buffer;
    private int _position;

    public WorkbookProtectionWriter(Span<byte> buffer)
    {
        _buffer = buffer;
        _position = 0;
    }

    public int Position => _position;

    /// <summary>
    /// Writes the workbook protection record (PROT4REV).
    /// </summary>
    public void WriteWorkbookProtection(WorkbookProtectionInfo protection)
    {
        // PROT4REV record (0x001A) - Protection for revision
        WriteRecordHeader(0x001A, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(protection.ProtectStructure ? 1 : 0));
        _position += 2;

        // PROT4REVPASS record (0x001B) - Protection password for revision
        if (protection.PasswordHash != 0)
        {
            WriteRecordHeader(0x001B, 2);
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), protection.PasswordHash);
            _position += 2;
        }

        // WINDOWPROTECT record (0x0019)
        WriteRecordHeader(0x0019, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(protection.ProtectWindows ? 1 : 0));
        _position += 2;
    }

    /// <summary>
    /// Writes all workbook protection records.
    /// </summary>
    public int WriteAllProtectionRecords(WorkbookProtectionInfo? protection)
    {
        if (protection == null || (!protection.Value.ProtectStructure && !protection.Value.ProtectWindows))
            return _position;

        WriteWorkbookProtection(protection.Value);
        return _position;
    }

    private void WriteRecordHeader(ushort type, int length)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), type);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)length);
        _position += 2;
    }
}
