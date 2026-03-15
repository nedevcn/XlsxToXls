using System.Buffers;
using System.Buffers.Binary;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Writes shape and image records in BIFF8 format.
/// </summary>
internal ref struct ShapeWriter
{
    private Span<byte> _buffer;
    private int _position;

    public ShapeWriter(Span<byte> buffer)
    {
        _buffer = buffer;
        _position = 0;
    }

    public int Position => _position;

    /// <summary>
    /// Creates a ShapeWriter using ArrayPool for buffer management.
    /// </summary>
    public static ShapeWriter CreatePooled(out byte[] buffer, int minSize = 65536)
    {
        buffer = ArrayPool<byte>.Shared.Rent(minSize);
        return new ShapeWriter(buffer.AsSpan());
    }

    /// <summary>
    /// Disposes the ShapeWriter and returns the buffer to the pool.
    /// </summary>
    public void Dispose()
    {
        // Buffer is managed externally
    }

    /// <summary>
    /// Writes all shape records for a worksheet.
    /// </summary>
    public int WriteAllShapes(List<ShapeInfo> shapes, int startingShapeId)
    {
        if (shapes.Count == 0)
            return _position;

        // Write MSODRAWING record containing all shapes
        WriteMsodrawingShapes(shapes, startingShapeId);

        // Write OBJ records for each shape
        var shapeId = startingShapeId;
        foreach (var shape in shapes)
        {
            WriteObjShape(shapeId, shape.Type, shape.Name);
            shapeId++;
        }

        return _position;
    }

    /// <summary>
    /// Writes the MSODRAWING record containing shape data.
    /// </summary>
    private void WriteMsodrawingShapes(List<ShapeInfo> shapes, int startingShapeId)
    {
        // Create Office Art container
        var artData = CreateOfficeArtContainer(shapes, startingShapeId);

        // Write MSODRAWING record (0x00EC)
        // Split into multiple records if data exceeds 8224 bytes
        const int maxRecordData = 8224;
        var offset = 0;

        while (offset < artData.Length)
        {
            var chunkSize = Math.Min(maxRecordData, artData.Length - offset);
            var isFirst = offset == 0;
            var isLast = offset + chunkSize >= artData.Length;

            WriteRecordHeader(0x00EC, chunkSize);
            artData.AsSpan(offset, chunkSize).CopyTo(_buffer.Slice(_position));
            _position += chunkSize;

            offset += chunkSize;
        }
    }

    /// <summary>
    /// Creates the Office Art container for shapes.
    /// </summary>
    private byte[] CreateOfficeArtContainer(List<ShapeInfo> shapes, int startingShapeId)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms);

        // Office Art DggContainer (Drawing Group Container)
        // This is a simplified implementation

        // Write Office Art FDG (File Drawing Group) record
        WriteOfficeArtFdg(writer, shapes.Count);

        // Write Office Art FSPGR (File Shape Group) record
        WriteOfficeArtFspgr(writer);

        // Write shape records
        var shapeId = startingShapeId;
        foreach (var shape in shapes)
        {
            WriteOfficeArtFsp(writer, shapeId, shape.Type);
            WriteOfficeArtFopt(writer, shape);
            WriteOfficeArtClientAnchor(writer, shape.Position);
            WriteOfficeArtClientData(writer);
            shapeId++;
        }

        return ms.ToArray();
    }

    /// <summary>
    /// Writes Office Art FDG record.
    /// </summary>
    private void WriteOfficeArtFdg(BinaryWriter writer, int shapeCount)
    {
        // Office Art Record Header
        // Ver/Instance (2 bytes) + Type (2 bytes) + Length (4 bytes)
        writer.Write((ushort)0x0000); // Version 0, Instance 0
        writer.Write((ushort)0xF000); // Type: FDG
        writer.Write((uint)(8 + shapeCount * 8)); // Length

        // FDG data
        writer.Write((uint)1); // spidMax (maximum shape ID)
        writer.Write((uint)shapeCount); // cidcl (cluster count)
    }

    /// <summary>
    /// Writes Office Art FSPGR (File Shape Group) record.
    /// </summary>
    private void WriteOfficeArtFspgr(BinaryWriter writer)
    {
        writer.Write((ushort)0x0001); // Version 1, Instance 0
        writer.Write((ushort)0xF009); // Type: FSPGR
        writer.Write((uint)16); // Length

        // Group bounds (left, top, right, bottom)
        writer.Write(0);
        writer.Write(0);
        writer.Write(0);
        writer.Write(0);
    }

    /// <summary>
    /// Writes Office Art FSP (File Shape) record.
    /// </summary>
    private void WriteOfficeArtFsp(BinaryWriter writer, int shapeId, ShapeType shapeType)
    {
        writer.Write((ushort)0x0002); // Version 2, Instance 0
        writer.Write((ushort)0xF00A); // Type: FSP
        writer.Write((uint)8); // Length

        // Shape ID and flags
        writer.Write((uint)shapeId);
        writer.Write((uint)0x00000005); // Flags: have anchor + have shape type
    }

    /// <summary>
    /// Writes Office Art FOPT (File Property Table) record.
    /// </summary>
    private void WriteOfficeArtFopt(BinaryWriter writer, ShapeInfo shape)
    {
        var properties = new List<(ushort propId, uint value, bool isComplex)>();

        // Add shape type property
        properties.Add((0x0080, (uint)shape.Type, false));

        // Add rotation if present
        if (shape.Rotation != 0)
        {
            properties.Add((0x0081, (uint)shape.Rotation, false));
        }

        // Add visibility
        if (!shape.Visible)
        {
            properties.Add((0x0082, 0x00010001, false)); // Hidden
        }

        // Calculate size
        var propSize = properties.Count * 6; // Each property is 6 bytes

        writer.Write((ushort)(0x0003 | (properties.Count << 4))); // Version 3, Instance = property count
        writer.Write((ushort)0xF00B); // Type: FOPT
        writer.Write((uint)propSize);

        // Write properties
        foreach (var (propId, value, isComplex) in properties)
        {
            var header = (ushort)(propId | (isComplex ? 0x8000 : 0x0000));
            writer.Write(header);
            writer.Write(value);
        }
    }

    /// <summary>
    /// Writes Office Art Client Anchor record.
    /// </summary>
    private void WriteOfficeArtClientAnchor(BinaryWriter writer, ShapePositionInfo position)
    {
        writer.Write((ushort)0x0000); // Version 0, Instance 0
        writer.Write((ushort)0xF010); // Type: ClientAnchor
        writer.Write((uint)18); // Length

        // Anchor type (0 = absolute, 1 = one-cell, 2 = two-cell)
        writer.Write((byte)(position.AnchorType == AnchorType.TwoCell ? 2 : 1));

        // From position
        writer.Write((ushort)position.FromCol);
        writer.Write((ushort)(position.FromColOffset / 12700)); // Convert EMU to pixels approx
        writer.Write((ushort)position.FromRow);
        writer.Write((ushort)(position.FromRowOffset / 12700));

        // To position
        writer.Write((ushort)position.ToCol);
        writer.Write((ushort)(position.ToColOffset / 12700));
        writer.Write((ushort)position.ToRow);
        writer.Write((ushort)(position.ToRowOffset / 12700));

        // Position relative to cell
        writer.Write((byte)1);
    }

    /// <summary>
    /// Writes Office Art Client Data record.
    /// </summary>
    private void WriteOfficeArtClientData(BinaryWriter writer)
    {
        writer.Write((ushort)0x0000); // Version 0, Instance 0
        writer.Write((ushort)0xF011); // Type: ClientData
        writer.Write((uint)0); // Empty
    }

    /// <summary>
    /// Writes OBJ record for a shape.
    /// </summary>
    public void WriteObjShape(int shapeId, ShapeType type, string name)
    {
        // OBJ record (0x005D)
        var objData = CreateObjData(shapeId, type, name);

        WriteRecordHeader(0x005D, objData.Length);
        objData.CopyTo(_buffer.Slice(_position));
        _position += objData.Length;
    }

    /// <summary>
    /// Creates OBJ record data.
    /// </summary>
    private byte[] CreateObjData(int shapeId, ShapeType type, string name)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms);

        // Common object data subrecord (0x0015)
        writer.Write((ushort)0x0015); // Type
        writer.Write((ushort)18); // Length

        // Object type based on shape type
        var objType = GetObjectType(type);
        writer.Write((ushort)objType);

        // Object ID
        writer.Write((ushort)shapeId);

        // Options
        writer.Write((ushort)0x4011);

        // Macro name (empty)
        writer.Write((ushort)0x0000);

        // More options
        writer.Write((ushort)0x0000);
        writer.Write((uint)0x00000000);

        // End marker (0x0000)
        writer.Write((ushort)0x0000);
        writer.Write((ushort)0x0000);

        return ms.ToArray();
    }

    /// <summary>
    /// Gets the object type for BIFF8 OBJ record.
    /// </summary>
    private static ushort GetObjectType(ShapeType shapeType)
    {
        return shapeType switch
        {
            ShapeType.Picture => 0x0008, // Picture
            ShapeType.TextBox => 0x0006, // Text box
            ShapeType.Rectangle or ShapeType.Oval or ShapeType.Arc => 0x001E, // Shape
            ShapeType.Line or ShapeType.Arrow => 0x0014, // Line
            _ => 0x001E // Default to shape
        };
    }

    /// <summary>
    /// Writes a picture/image to the worksheet.
    /// </summary>
    public int WritePicture(int shapeId, ImageInfo image, ShapePositionInfo position)
    {
        // Write MSODRAWING record with picture data
        var drawingData = CreatePictureDrawingData(shapeId, image, position);

        // Split into MSODRAWING records if needed
        const int maxRecordData = 8224;
        var offset = 0;

        while (offset < drawingData.Length)
        {
            var chunkSize = Math.Min(maxRecordData, drawingData.Length - offset);
            WriteRecordHeader(0x00EC, chunkSize);
            drawingData.AsSpan(offset, chunkSize).CopyTo(_buffer.Slice(_position));
            _position += chunkSize;
            offset += chunkSize;
        }

        // Write OBJ record
        WriteObjPicture(shapeId, image);

        return _position;
    }

    /// <summary>
    /// Creates drawing data for a picture.
    /// </summary>
    private byte[] CreatePictureDrawingData(int shapeId, ImageInfo image, ShapePositionInfo position)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms);

        // Office Art BStore container for image
        WriteOfficeArtBstoreContainer(writer, image);

        // Shape container
        WriteOfficeArtFsp(writer, shapeId, ShapeType.Picture);
        WriteOfficeArtFopt(writer, new ShapeInfo(
            shapeId, ShapeType.Picture, $"Picture {shapeId}", position, image, true, 0, shapeId));
        WriteOfficeArtClientAnchor(writer, position);
        WriteOfficeArtClientData(writer);

        return ms.ToArray();
    }

    /// <summary>
    /// Writes Office Art BStore (Blip Store) container.
    /// </summary>
    private void WriteOfficeArtBstoreContainer(BinaryWriter writer, ImageInfo image)
    {
        // Simplified - in real implementation, this would reference the image data
        writer.Write((ushort)0x0001); // Version 1, Instance 1 (1 blip)
        writer.Write((ushort)0xF001); // Type: BStoreContainer
        writer.Write((uint)8); // Length placeholder

        // Write blip reference
        WriteOfficeArtFBSE(writer, image);
    }

    /// <summary>
    /// Writes Office Art FBSE (File Blip Store Entry) record.
    /// </summary>
    private void WriteOfficeArtFBSE(BinaryWriter writer, ImageInfo image)
    {
        var blipType = GetBlipType(image.Format);

        writer.Write((ushort)0x0002); // Version 2, Instance = blip type
        writer.Write((ushort)0xF007); // Type: FBSE
        writer.Write((uint)(36 + image.Data.Length)); // Length

        // BTWin32
        writer.Write((byte)blipType);
        // BTMacOS
        writer.Write((byte)blipType);

        // UID (16 bytes) - MD4 hash of image data (simplified)
        var uid = ComputeImageHash(image.Data);
        writer.Write(uid, 0, 16);

        // Tag
        writer.Write((ushort)0xFFFE);

        // Size
        writer.Write((uint)image.Data.Length);

        // Ref
        writer.Write((uint)1);

        // Offset
        writer.Write((uint)0);

        // Usage
        writer.Write((byte)0);

        // Name length
        writer.Write((byte)0);

        // Unused
        writer.Write((byte)0);

        // Unused
        writer.Write((byte)0);

        // Blip data would follow here in a full implementation
    }

    /// <summary>
    /// Gets the Office Art blip type for an image format.
    /// </summary>
    private static byte GetBlipType(ImageFormat format)
    {
        return format switch
        {
            ImageFormat.Png => 0x06, // PNG
            ImageFormat.Jpeg => 0x05, // JPEG
            ImageFormat.Gif => 0x02, // GIF
            ImageFormat.Bmp => 0x01, // DIB
            ImageFormat.Tiff => 0x07, // TIFF
            ImageFormat.Wmf => 0x03, // WMF
            ImageFormat.Emf => 0x04, // EMF
            _ => 0x00 // Unknown
        };
    }

    /// <summary>
    /// Computes a simple hash of image data for UID.
    /// </summary>
    private static byte[] ComputeImageHash(byte[] data)
    {
        var hash = new byte[16];
        for (var i = 0; i < data.Length; i++)
        {
            hash[i % 16] ^= data[i];
        }
        return hash;
    }

    /// <summary>
    /// Writes OBJ record for a picture.
    /// </summary>
    private void WriteObjPicture(int shapeId, ImageInfo image)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms);

        // Common object data subrecord
        writer.Write((ushort)0x0015);
        writer.Write((ushort)18);
        writer.Write((ushort)0x0008); // Picture type
        writer.Write((ushort)shapeId);
        writer.Write((ushort)0x6011);
        writer.Write((ushort)0x0000);
        writer.Write((ushort)0x0000);
        writer.Write((uint)0x00000000);

        // Picture options subrecord (0x0008)
        writer.Write((ushort)0x0008);
        writer.Write((ushort)2);
        writer.Write((ushort)0x0000);

        // End marker
        writer.Write((ushort)0x0000);
        writer.Write((ushort)0x0000);

        var objData = ms.ToArray();
        WriteRecordHeader(0x005D, objData.Length);
        objData.CopyTo(_buffer.Slice(_position));
        _position += objData.Length;
    }

    /// <summary>
    /// Writes a BIFF record header.
    /// </summary>
    private void WriteRecordHeader(ushort recordType, int dataLength)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), recordType);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)dataLength);
        _position += 2;
    }
}
