namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Represents a shape or image in a worksheet.
/// </summary>
public sealed class ShapeData
{
    /// <summary>
    /// Gets or sets the shape identifier.
    /// </summary>
    public int ShapeId { get; set; }

    /// <summary>
    /// Gets or sets the shape type.
    /// </summary>
    public ShapeType Type { get; set; }

    /// <summary>
    /// Gets or sets the shape name.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the shape description.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Gets or sets the position and size of the shape.
    /// </summary>
    public ShapePosition Position { get; set; } = new();

    /// <summary>
    /// Gets or sets the image data for picture shapes.
    /// </summary>
    public ImageData? Image { get; set; }

    /// <summary>
    /// Gets or sets the fill properties.
    /// </summary>
    public FillProperties? Fill { get; set; }

    /// <summary>
    /// Gets or sets the line properties.
    /// </summary>
    public LineProperties? Line { get; set; }

    /// <summary>
    /// Gets or sets the text content for text boxes.
    /// </summary>
    public string? Text { get; set; }

    /// <summary>
    /// Gets or sets the text formatting properties.
    /// </summary>
    public TextProperties? TextFormat { get; set; }

    /// <summary>
    /// Gets or sets whether the shape is visible.
    /// </summary>
    public bool Visible { get; set; } = true;

    /// <summary>
    /// Gets or sets the rotation angle in degrees.
    /// </summary>
    public int Rotation { get; set; }

    /// <summary>
    /// Gets or sets the z-order index.
    /// </summary>
    public int ZOrder { get; set; }
}

/// <summary>
/// Shape types supported in BIFF8 format.
/// </summary>
public enum ShapeType
{
    /// <summary>
    /// Picture/image.
    /// </summary>
    Picture = 0x0000,

    /// <summary>
    /// Rectangle shape.
    /// </summary>
    Rectangle = 0x0001,

    /// <summary>
    /// Oval/circle shape.
    /// </summary>
    Oval = 0x0002,

    /// <summary>
    /// Arc shape.
    /// </summary>
    Arc = 0x0004,

    /// <summary>
    /// Text box.
    /// </summary>
    TextBox = 0x0006,

    /// <summary>
    /// Line shape.
    /// </summary>
    Line = 0x0014,

    /// <summary>
    /// Arrow shape.
    /// </summary>
    Arrow = 0x0015,

    /// <summary>
    /// Freeform shape.
    /// </summary>
    Freeform = 0x0018,

    /// <summary>
    /// Group shape.
    /// </summary>
    Group = 0x0064
}

/// <summary>
/// Position and size information for a shape.
/// </summary>
public sealed class ShapePosition
{
    /// <summary>
    /// Gets or sets the anchor type.
    /// </summary>
    public AnchorType AnchorType { get; set; } = AnchorType.OneCell;

    /// <summary>
    /// Gets or sets the column index of the top-left corner.
    /// </summary>
    public int FromCol { get; set; }

    /// <summary>
    /// Gets or sets the row index of the top-left corner.
    /// </summary>
    public int FromRow { get; set; }

    /// <summary>
    /// Gets or sets the offset from the left of the from column (in EMUs).
    /// </summary>
    public int FromColOffset { get; set; }

    /// <summary>
    /// Gets or sets the offset from the top of the from row (in EMUs).
    /// </summary>
    public int FromRowOffset { get; set; }

    /// <summary>
    /// Gets or sets the column index of the bottom-right corner.
    /// </summary>
    public int ToCol { get; set; }

    /// <summary>
    /// Gets or sets the row index of the bottom-right corner.
    /// </summary>
    public int ToRow { get; set; }

    /// <summary>
    /// Gets or sets the offset from the left of the to column (in EMUs).
    /// </summary>
    public int ToColOffset { get; set; }

    /// <summary>
    /// Gets or sets the offset from the top of the to row (in EMUs).
    /// </summary>
    public int ToRowOffset { get; set; }

    /// <summary>
    /// Gets or sets the width in EMUs (English Metric Units).
    /// </summary>
    public int WidthEmu { get; set; }

    /// <summary>
    /// Gets or sets the height in EMUs.
    /// </summary>
    public int HeightEmu { get; set; }
}

/// <summary>
/// Anchor types for shape positioning.
/// </summary>
public enum AnchorType
{
    /// <summary>
    /// Absolute positioning.
    /// </summary>
    Absolute,

    /// <summary>
    /// One-cell anchor (top-left anchored to a cell).
    /// </summary>
    OneCell,

    /// <summary>
    /// Two-cell anchor (both corners anchored to cells).
    /// </summary>
    TwoCell
}

/// <summary>
/// Image data for picture shapes.
/// </summary>
public sealed class ImageData
{
    /// <summary>
    /// Gets or sets the image format.
    /// </summary>
    public ImageFormat Format { get; set; }

    /// <summary>
    /// Gets or sets the image binary data.
    /// </summary>
    public byte[] Data { get; set; } = Array.Empty<byte>();

    /// <summary>
    /// Gets or sets the image relationship ID in XLSX.
    /// </summary>
    public string? RelationshipId { get; set; }

    /// <summary>
    /// Gets or sets the content type (MIME type).
    /// </summary>
    public string? ContentType { get; set; }
}

/// <summary>
/// Supported image formats.
/// </summary>
public enum ImageFormat
{
    /// <summary>
    /// Unknown format.
    /// </summary>
    Unknown,

    /// <summary>
    /// PNG format.
    /// </summary>
    Png,

    /// <summary>
    /// JPEG format.
    /// </summary>
    Jpeg,

    /// <summary>
    /// GIF format.
    /// </summary>
    Gif,

    /// <summary>
    /// BMP format.
    /// </summary>
    Bmp,

    /// <summary>
    /// TIFF format.
    /// </summary>
    Tiff,

    /// <summary>
    /// WMF format.
    /// </summary>
    Wmf,

    /// <summary>
    /// EMF format.
    /// </summary>
    Emf
}

/// <summary>
/// Fill properties for shapes.
/// </summary>
public sealed class FillProperties
{
    /// <summary>
    /// Gets or sets the fill type.
    /// </summary>
    public FillType Type { get; set; } = FillType.Solid;

    /// <summary>
    /// Gets or sets the fill color (ARGB).
    /// </summary>
    public uint Color { get; set; } = 0xFFFFFFFF;

    /// <summary>
    /// Gets or sets the transparency (0-255, 255 = fully transparent).
    /// </summary>
    public byte Transparency { get; set; }
}

/// <summary>
/// Fill types.
/// </summary>
public enum FillType
{
    /// <summary>
    /// No fill.
    /// </summary>
    None,

    /// <summary>
    /// Solid color fill.
    /// </summary>
    Solid,

    /// <summary>
    /// Gradient fill.
    /// </summary>
    Gradient,

    /// <summary>
    /// Pattern fill.
    /// </summary>
    Pattern,

    /// <summary>
    /// Picture/texture fill.
    /// </summary>
    Picture
}

/// <summary>
/// Line properties for shapes.
/// </summary>
public sealed class LineProperties
{
    /// <summary>
    /// Gets or sets the line style.
    /// </summary>
    public ShapeLineStyle Style { get; set; } = ShapeLineStyle.Solid;

    /// <summary>
    /// Gets or sets the line width in EMUs.
    /// </summary>
    public int Width { get; set; } = 9525; // 0.75pt default

    /// <summary>
    /// Gets or sets the line color (ARGB).
    /// </summary>
    public uint Color { get; set; } = 0xFF000000;
}

/// <summary>
/// Shape line styles.
/// </summary>
public enum ShapeLineStyle
{
    /// <summary>
    /// No line.
    /// </summary>
    None,

    /// <summary>
    /// Solid line.
    /// </summary>
    Solid,

    /// <summary>
    /// Dashed line.
    /// </summary>
    Dash,

    /// <summary>
    /// Dotted line.
    /// </summary>
    Dot,

    /// <summary>
    /// Dash-dot pattern.
    /// </summary>
    DashDot,

    /// <summary>
    /// Dash-dot-dot pattern.
    /// </summary>
    DashDotDot
}

/// <summary>
/// Text properties for text boxes.
/// </summary>
public sealed class TextProperties
{
    /// <summary>
    /// Gets or sets the font name.
    /// </summary>
    public string FontName { get; set; } = "Arial";

    /// <summary>
    /// Gets or sets the font size in points.
    /// </summary>
    public double FontSize { get; set; } = 11;

    /// <summary>
    /// Gets or sets whether the text is bold.
    /// </summary>
    public bool Bold { get; set; }

    /// <summary>
    /// Gets or sets whether the text is italic.
    /// </summary>
    public bool Italic { get; set; }

    /// <summary>
    /// Gets or sets the text color (ARGB).
    /// </summary>
    public uint Color { get; set; } = 0xFF000000;

    /// <summary>
    /// Gets or sets the horizontal alignment.
    /// </summary>
    public TextAlignment HorizontalAlignment { get; set; } = TextAlignment.Left;

    /// <summary>
    /// Gets or sets the vertical alignment.
    /// </summary>
    public TextVerticalAlignment VerticalAlignment { get; set; } = TextVerticalAlignment.Top;

    /// <summary>
    /// Gets or sets whether text wraps within the shape.
    /// </summary>
    public bool WrapText { get; set; } = true;
}

/// <summary>
/// Horizontal text alignment.
/// </summary>
public enum TextAlignment
{
    /// <summary>
    /// Left alignment.
    /// </summary>
    Left,

    /// <summary>
    /// Center alignment.
    /// </summary>
    Center,

    /// <summary>
    /// Right alignment.
    /// </summary>
    Right,

    /// <summary>
    /// Justified alignment.
    /// </summary>
    Justify
}

/// <summary>
/// Vertical text alignment.
/// </summary>
public enum TextVerticalAlignment
{
    /// <summary>
    /// Top alignment.
    /// </summary>
    Top,

    /// <summary>
    /// Middle alignment.
    /// </summary>
    Middle,

    /// <summary>
    /// Bottom alignment.
    /// </summary>
    Bottom
}

/// <summary>
/// Internal record for shape information used during conversion.
/// </summary>
internal record struct ShapeInfo(
    int ShapeId,
    ShapeType Type,
    string Name,
    ShapePositionInfo Position,
    ImageInfo? Image,
    bool Visible,
    int Rotation,
    int ZOrder);

/// <summary>
/// Internal record for shape position.
/// </summary>
internal record struct ShapePositionInfo(
    AnchorType AnchorType,
    int FromCol,
    int FromRow,
    int FromColOffset,
    int FromRowOffset,
    int ToCol,
    int ToRow,
    int ToColOffset,
    int ToRowOffset,
    int WidthEmu,
    int HeightEmu)
{
    // Additional constructor for compatibility with code using 13 parameters
    public ShapePositionInfo(
        AnchorType anchorType,
        int fromCol, int fromRow, int fromColOffset, int fromRowOffset,
        int toCol, int toRow, int toColOffset, int toRowOffset,
        int widthEmu, int heightEmu,
        int unused1, int unused2) : this(anchorType, fromCol, fromRow, fromColOffset, fromRowOffset,
        toCol, toRow, toColOffset, toRowOffset, widthEmu, heightEmu)
    {
    }
}

/// <summary>
/// Internal record for image information.
/// </summary>
internal record struct ImageInfo(
    ImageFormat Format,
    byte[] Data,
    string? RelationshipId,
    string? ContentType);
