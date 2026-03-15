using System.Buffers;
using System.Buffers.Binary;
using System.Text;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Writes pivot table definitions to BIFF8 format.
/// Generates SXVIEW, SXVD, SXVI, SXVDEX, SXPI, SXDI, and related records.
/// </summary>
public sealed class PivotTableWriter : IDisposable
{
    private const int DefaultBufferSize = 65536;

    private byte[] _buffer;
    private Memory<byte> _memory;
    private int _position;
    private bool _isRented;

    /// <summary>
    /// Creates a new PivotTableWriter with the specified buffer.
    /// </summary>
    public PivotTableWriter(byte[] buffer)
    {
        _buffer = buffer;
        _memory = buffer;
        _position = 0;
        _isRented = false;
    }

    /// <summary>
    /// Creates a PivotTableWriter using a pooled buffer.
    /// </summary>
    public static PivotTableWriter CreatePooled(out byte[] buffer, int minimumLength = DefaultBufferSize)
    {
        buffer = ArrayPool<byte>.Shared.Rent(minimumLength);
        return new PivotTableWriter(buffer) { _isRented = true };
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
    /// Writes a pivot table with its cache definition.
    /// </summary>
    public int WritePivotTable(PivotTableData pivotTable, PivotCacheDefinition cache, int sheetIndex)
    {
        // Write SXVIEW record (pivot table view)
        WriteSxView(pivotTable);

        // Write SXVD records (view dimensions) for each axis
        WriteSxvdRecords(pivotTable);

        // Write SXVI records (view items) for each field
        WriteSxviRecords(pivotTable, cache);

        // Write SXVDEX records (view dimension extensions)
        WriteSxvdexRecords(pivotTable);

        // Write SXPI records (page items) for page fields
        WriteSxpiRecords(pivotTable);

        // Write SXDI records (data items) for data fields
        WriteSxdiRecords(pivotTable);

        // Write SXEX records (extensions)
        WriteSxexRecords(pivotTable);

        return _position;
    }

    /// <summary>
    /// Writes a pivot cache definition.
    /// </summary>
    public int WritePivotCache(PivotCacheDefinition cache)
    {
        // Write SXDB record (cache definition)
        WriteSxdb(cache);

        // Write SXFIELD records for each field
        WriteSxfieldRecords(cache);

        // Write SXSTRING records for shared strings
        WriteSxstringRecords(cache);

        // Write SXDBEX record (cache definition extension)
        WriteSxdbex(cache);

        return _position;
    }

    private void WriteSxView(PivotTableData pivotTable)
    {
        // SXVIEW record (0x00B0) - Pivot Table View
        // This is the main record that defines the pivot table layout

        const ushort recordId = 0x00B0;
        const ushort recordLength = 52; // Fixed length for basic SXVIEW

        WriteRecordHeader(recordId, recordLength);

        // Reference to source range (8 bytes) - SXDBREF structure
        WriteSxdbref(pivotTable.SourceRange);

        // Location row (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)pivotTable.Location.Row);
        _position += 2;

        // Location column (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)pivotTable.Location.Column);
        _position += 2;

        // Cache ID (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)pivotTable.CacheId);
        _position += 2;

        // Reserved (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 2;

        // Row field count (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)pivotTable.RowFields.Count);
        _position += 2;

        // Column field count (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)pivotTable.ColumnFields.Count);
        _position += 2;

        // Page field count (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)pivotTable.PageFields.Count);
        _position += 2;

        // Data field count (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)pivotTable.DataFields.Count);
        _position += 2;

        // Data row count (2 bytes) - number of data rows
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)pivotTable.DataFields.Count);
        _position += 2;

        // Data column count (2 bytes) - number of data columns
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 1);
        _position += 2;

        // Flags (2 bytes)
        ushort flags = 0;
        if (pivotTable.ShowRowGrandTotals) flags |= 0x0001;
        if (pivotTable.ShowColumnGrandTotals) flags |= 0x0002;
        if (pivotTable.ShowError) flags |= 0x0004;
        if (pivotTable.ShowEmpty) flags |= 0x0008;
        if (pivotTable.AutoFormat) flags |= 0x0010;
        if (pivotTable.PreserveFormatting) flags |= 0x0020;
        if (pivotTable.UseCustomLists) flags |= 0x0040;
        if (pivotTable.ShowExpandCollapseButtons) flags |= 0x0080;
        if (pivotTable.ShowFieldHeaders) flags |= 0x0100;
        if (pivotTable.OutlineForm) flags |= 0x0200;
        if (pivotTable.CompactRowAxis) flags |= 0x0400;
        if (pivotTable.CompactColumnAxis) flags |= 0x0800;

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), flags);
        _position += 2;

        // Auto format type (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 2;

        // Reserved (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 2;

        // First row of data (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(pivotTable.Location.Row + 1));
        _position += 2;

        // First column of data (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(pivotTable.Location.Column + 1));
        _position += 2;

        // First row of headers (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)pivotTable.Location.Row);
        _position += 2;

        // Row page count (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 2;

        // Column page count (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 2;
    }

    private void WriteSxdbref(CellRange? range)
    {
        // SXDBREF structure - reference to source data
        if (range == null)
        {
            // Write empty reference
            for (var i = 0; i < 8; i++)
            {
                _buffer[_position++] = 0;
            }
            return;
        }

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

    private void WriteSxvdRecords(PivotTableData pivotTable)
    {
        // SXVD record (0x00B1) - View Dimension
        // One record for each field in the pivot table

        var allFields = new List<PivotField>();
        allFields.AddRange(pivotTable.RowFields);
        allFields.AddRange(pivotTable.ColumnFields);
        allFields.AddRange(pivotTable.PageFields);
        allFields.AddRange(pivotTable.HiddenFields);

        foreach (var field in allFields)
        {
            WriteSxvd(field);
        }
    }

    private void WriteSxvd(PivotField field)
    {
        const ushort recordId = 0x00B1;
        const ushort recordLength = 10;

        WriteRecordHeader(recordId, recordLength);

        // Field index in cache (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)field.FieldIndex);
        _position += 2;

        // Axis type (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)field.Axis);
        _position += 2;

        // Subtotal type (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)field.Subtotal);
        _position += 2;

        // Item count (2 bytes) - will be updated later
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 2;

        // Flags (2 bytes)
        ushort flags = 0;
        if (field.SubtotalTop) flags |= 0x0001;
        if (field.ShowAllItems) flags |= 0x0002;
        if (field.InsertBlankRows) flags |= 0x0004;
        if (field.InsertPageBreaks) flags |= 0x0008;
        if (field.AutoSort) flags |= 0x0010;
        if (field.AutoShow) flags |= 0x0020;
        if (field.SortOrder == SortOrder.Descending) flags |= 0x0040;

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), flags);
        _position += 2;
    }

    private void WriteSxviRecords(PivotTableData pivotTable, PivotCacheDefinition cache)
    {
        // SXVI record (0x00B2) - View Item
        // One record for each item in each field

        var allFields = new List<PivotField>();
        allFields.AddRange(pivotTable.RowFields);
        allFields.AddRange(pivotTable.ColumnFields);

        foreach (var field in allFields)
        {
            // Get the corresponding cache field
            if (field.FieldIndex < cache.Fields.Count)
            {
                var cacheField = cache.Fields[field.FieldIndex];

                // Write item for each shared item in the cache
                for (var i = 0; i < cacheField.SharedItems.Count; i++)
                {
                    WriteSxvi(i, field.HiddenItems.Contains(i));
                }

                // Write grand total item
                WriteSxvi(0x7FFF, false, true);
            }
        }
    }

    private void WriteSxvi(int itemIndex, bool isHidden, bool isGrandTotal = false)
    {
        const ushort recordId = 0x00B2;
        const ushort recordLength = 6;

        WriteRecordHeader(recordId, recordLength);

        // Item index (2 bytes)
        BinaryPrimitives.WriteInt16LittleEndian(_buffer.AsSpan(_position), (short)itemIndex);
        _position += 2;

        // Flags (2 bytes)
        ushort flags = 0;
        if (isHidden) flags |= 0x0001;
        if (isGrandTotal) flags |= 0x0002;

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), flags);
        _position += 2;

        // Reserved (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 2;
    }

    private void WriteSxvdexRecords(PivotTableData pivotTable)
    {
        // SXVDEX record (0x0100) - View Dimension Extension
        // Extended properties for view dimensions

        var allFields = new List<PivotField>();
        allFields.AddRange(pivotTable.RowFields);
        allFields.AddRange(pivotTable.ColumnFields);
        allFields.AddRange(pivotTable.PageFields);

        foreach (var field in allFields)
        {
            WriteSxvdex(field);
        }
    }

    private void WriteSxvdex(PivotField field)
    {
        const ushort recordId = 0x0100;
        const ushort recordLength = 20;

        WriteRecordHeader(recordId, recordLength);

        // Auto sort field (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(field.AutoSortField ?? 0xFFFF));
        _position += 2;

        // Auto show field (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(field.AutoShowField ?? 0xFFFF));
        _position += 2;

        // Auto show count (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)field.AutoShowCount);
        _position += 2;

        // Auto show type (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)field.AutoShowType);
        _position += 2;

        // Number format index (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(field.NumberFormat ?? 0));
        _position += 2;

        // Outline level (1 byte)
        _buffer[_position++] = field.OutlineLevel;

        // Compact flag (1 byte)
        _buffer[_position++] = field.Compact ? (byte)1 : (byte)0;

        // Reserved (6 bytes)
        for (var i = 0; i < 6; i++)
        {
            _buffer[_position++] = 0;
        }
    }

    private void WriteSxpiRecords(PivotTableData pivotTable)
    {
        // SXPI record (0x00B3) - Page Item
        // One record for each page field

        foreach (var field in pivotTable.PageFields)
        {
            WriteSxpi(field);
        }
    }

    private void WriteSxpi(PivotField field)
    {
        const ushort recordId = 0x00B3;
        const ushort recordLength = 6;

        WriteRecordHeader(recordId, recordLength);

        // Field index (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)field.FieldIndex);
        _position += 2;

        // Selected item index (2 bytes) - 0x7FFF means "(All)"
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0x7FFF);
        _position += 2;

        // Flags (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 2;
    }

    private void WriteSxdiRecords(PivotTableData pivotTable)
    {
        // SXDI record (0x00B5) - Data Item
        // One record for each data field

        for (var i = 0; i < pivotTable.DataFields.Count; i++)
        {
            WriteSxdi(pivotTable.DataFields[i], i);
        }
    }

    private void WriteSxdi(PivotDataField field, int position)
    {
        const ushort recordId = 0x00B5;
        const ushort recordLength = 14;

        WriteRecordHeader(recordId, recordLength);

        // Field index (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)field.FieldIndex);
        _position += 2;

        // Aggregation function (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)field.Function);
        _position += 2;

        // Number format index (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(field.NumberFormat ?? 0));
        _position += 2;

        // Position (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)position);
        _position += 2;

        // Base field (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(field.BaseField ?? 0xFFFF));
        _position += 2;

        // Base item (2 bytes)
        BinaryPrimitives.WriteInt16LittleEndian(_buffer.AsSpan(_position), (short)(field.BaseItem ?? 0x7FFF));
        _position += 2;

        // Show data as (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)field.ShowDataAs);
        _position += 2;
    }

    private void WriteSxexRecords(PivotTableData pivotTable)
    {
        // SXEX record (0x00F1) - Pivot Table Extensions
        // Contains additional pivot table properties

        const ushort recordId = 0x00F1;
        const ushort recordLength = 16;

        WriteRecordHeader(recordId, recordLength);

        // Flags (4 bytes)
        uint flags = 0;
        if (pivotTable.ErrorString != null) flags |= 0x00000001;
        if (pivotTable.EmptyString != null) flags |= 0x00000002;

        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), flags);
        _position += 4;

        // Merge labels (4 bytes)
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)pivotTable.MergeLabels);
        _position += 4;

        // Page wrap (4 bytes)
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)pivotTable.PageWrap);
        _position += 4;

        // Page filter order (4 bytes)
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)pivotTable.PageFilterOrder);
        _position += 4;

        // Write error string if present
        if (pivotTable.ErrorString != null)
        {
            WriteSxstring(pivotTable.ErrorString);
        }

        // Write empty string if present
        if (pivotTable.EmptyString != null)
        {
            WriteSxstring(pivotTable.EmptyString);
        }
    }

    private void WriteSxdb(PivotCacheDefinition cache)
    {
        // SXDB record (0x00C0) - Cache Definition
        // Defines the pivot cache

        const ushort recordId = 0x00C0;
        const ushort recordLength = 24;

        WriteRecordHeader(recordId, recordLength);

        // Record count (4 bytes)
        BinaryPrimitives.WriteInt32LittleEndian(_buffer.AsSpan(_position), cache.RecordCount);
        _position += 4;

        // Field count (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)cache.Fields.Count);
        _position += 2;

        // Type (2 bytes) - 1 for worksheet source
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 1);
        _position += 2;

        // Flags (2 bytes)
        ushort flags = 0;
        if (cache.RefreshOnLoad) flags |= 0x0001;

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), flags);
        _position += 2;

        // Reserved (4 bytes)
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 4;

        // Block count (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 2;

        // Created version (1 byte)
        _buffer[_position++] = (byte)cache.CreatedVersion;

        // Refreshed version (1 byte)
        _buffer[_position++] = (byte)cache.RefreshedVersion;

        // Min refreshable version (1 byte)
        _buffer[_position++] = (byte)cache.MinRefreshableVersion;

        // Reserved (3 bytes)
        for (var i = 0; i < 3; i++)
        {
            _buffer[_position++] = 0;
        }
    }

    private void WriteSxfieldRecords(PivotCacheDefinition cache)
    {
        // SXFIELD record (0x00C1) - Cache Field
        // One record for each field in the cache

        foreach (var field in cache.Fields)
        {
            WriteSxfield(field);
        }
    }

    private void WriteSxfield(PivotCacheField field)
    {
        const ushort recordId = 0x00C1;

        // Calculate record length
        var nameBytes = Encoding.Unicode.GetBytes(field.Name);
        var recordLength = (ushort)(6 + nameBytes.Length);

        WriteRecordHeader(recordId, recordLength);

        // Field type (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)field.Type);
        _position += 2;

        // Shared item count (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)field.SharedItems.Count);
        _position += 2;

        // Flags (2 bytes)
        ushort flags = 0;
        if (field.MixedTypes) flags |= 0x0001;

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), flags);
        _position += 2;

        // Name length (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)field.Name.Length);
        _position += 2;

        // Name (Unicode)
        nameBytes.CopyTo(_buffer.AsSpan(_position));
        _position += nameBytes.Length;
    }

    private void WriteSxstringRecords(PivotCacheDefinition cache)
    {
        // SXSTRING record (0x00CD) - Shared String
        // Write all shared items from all fields

        foreach (var field in cache.Fields)
        {
            foreach (var item in field.SharedItems)
            {
                WriteSxstring(item);
            }
        }
    }

    private void WriteSxstring(string value)
    {
        const ushort recordId = 0x00CD;

        var bytes = Encoding.Unicode.GetBytes(value);
        var recordLength = (ushort)(2 + bytes.Length);

        WriteRecordHeader(recordId, recordLength);

        // String length (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)value.Length);
        _position += 2;

        // String data (Unicode)
        bytes.CopyTo(_buffer.AsSpan(_position));
        _position += bytes.Length;
    }

    private void WriteSxdbex(PivotCacheDefinition cache)
    {
        // SXDBEX record (0x0122) - Cache Definition Extension
        // Extended cache properties

        const ushort recordId = 0x0122;
        const ushort recordLength = 12;

        WriteRecordHeader(recordId, recordLength);

        // Flags (4 bytes)
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 4;

        // Reserved (4 bytes)
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 4;

        // Reserved (4 bytes)
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), 0);
        _position += 4;
    }

    private void WriteRecordHeader(ushort recordId, ushort recordLength)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordLength);
        _position += 2;
    }
}
