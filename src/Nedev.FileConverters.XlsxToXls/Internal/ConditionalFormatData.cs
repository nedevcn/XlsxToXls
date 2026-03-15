namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Represents conditional formatting data for BIFF8 format conversion.
/// Contains all information needed to apply conditional formatting rules.
/// </summary>
public sealed class ConditionalFormatData
{
    /// <summary>Gets or sets the unique identifier for this conditional format.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>Gets or sets the cell ranges this format applies to.</summary>
    public List<CellRange> Ranges { get; set; } = [];

    /// <summary>Gets or sets the type of conditional formatting.</summary>
    public ConditionalFormatType Type { get; set; } = ConditionalFormatType.CellIs;

    /// <summary>Gets or sets the operator for cell value comparison.</summary>
    public ComparisonOperator Operator { get; set; } = ComparisonOperator.GreaterThan;

    /// <summary>Gets or sets the first formula/value for comparison.</summary>
    public string? Formula1 { get; set; }

    /// <summary>Gets or sets the second formula/value for comparison (for between operators).</summary>
    public string? Formula2 { get; set; }

    /// <summary>Gets or sets the text value for text-based rules.</summary>
    public string? Text { get; set; }

    /// <summary>Gets or sets the style/format to apply when condition is met.</summary>
    public ConditionalFormatStyle? Style { get; set; }

    /// <summary>Gets or sets the color scale configuration.</summary>
    public ColorScale? ColorScale { get; set; }

    /// <summary>Gets or sets the data bar configuration.</summary>
    public DataBar? DataBar { get; set; }

    /// <summary>Gets or sets the icon set configuration.</summary>
    public IconSet? IconSet { get; set; }

    /// <summary>Gets or sets the priority of this rule (lower = higher priority).</summary>
    public int Priority { get; set; } = 1;

    /// <summary>Gets or sets whether to stop evaluating other rules if this one matches.</summary>
    public bool StopIfTrue { get; set; }
}

/// <summary>Types of conditional formatting rules.</summary>
public enum ConditionalFormatType : byte
{
    /// <summary>Format based on cell value comparison.</summary>
    CellIs = 0,

    /// <summary>Format cells that contain specific text.</summary>
    ContainsText = 1,

    /// <summary>Format cells that do not contain specific text.</summary>
    NotContainsText = 2,

    /// <summary>Format cells that begin with specific text.</summary>
    BeginsWith = 3,

    /// <summary>Format cells that end with specific text.</summary>
    EndsWith = 4,

    /// <summary>Format cells that contain specific date.</summary>
    ContainsDate = 5,

    /// <summary>Format top N values.</summary>
    Top10 = 6,

    /// <summary>Format unique or duplicate values.</summary>
    UniqueValues = 7,

    /// <summary>Format using a formula.</summary>
    Expression = 8,

    /// <summary>Format using a color scale.</summary>
    ColorScale = 9,

    /// <summary>Format using data bars.</summary>
    DataBar = 10,

    /// <summary>Format using icon sets.</summary>
    IconSet = 11,

    /// <summary>Format above average values.</summary>
    AboveAverage = 12,

    /// <summary>Format below average values.</summary>
    BelowAverage = 13
}

/// <summary>Comparison operators for cell value rules.</summary>
public enum ComparisonOperator : byte
{
    /// <summary>No comparison.</summary>
    None = 0,

    /// <summary>Between two values.</summary>
    Between = 1,

    /// <summary>Not between two values.</summary>
    NotBetween = 2,

    /// <summary>Equal to value.</summary>
    Equal = 3,

    /// <summary>Not equal to value.</summary>
    NotEqual = 4,

    /// <summary>Greater than value.</summary>
    GreaterThan = 5,

    /// <summary>Less than value.</summary>
    LessThan = 6,

    /// <summary>Greater than or equal to value.</summary>
    GreaterThanOrEqual = 7,

    /// <summary>Less than or equal to value.</summary>
    LessThanOrEqual = 8
}

/// <summary>Represents a cell range for conditional formatting.</summary>
public sealed class CellRange
{
    /// <summary>Gets or sets the first row index (0-based).</summary>
    public int FirstRow { get; set; }

    /// <summary>Gets or sets the first column index (0-based).</summary>
    public int FirstCol { get; set; }

    /// <summary>Gets or sets the last row index (0-based).</summary>
    public int LastRow { get; set; }

    /// <summary>Gets or sets the last column index (0-based).</summary>
    public int LastCol { get; set; }

    /// <summary>Creates a range from cell references.</summary>
    public static CellRange FromCells(int firstRow, int firstCol, int lastRow, int lastCol)
    {
        return new CellRange
        {
            FirstRow = firstRow,
            FirstCol = firstCol,
            LastRow = lastRow,
            LastCol = lastCol
        };
    }
}

/// <summary>Style to apply when condition is met.</summary>
public sealed class ConditionalFormatStyle
{
    /// <summary>Gets or sets the font color.</summary>
    public ChartColor? FontColor { get; set; }

    /// <summary>Gets or sets whether to make font bold.</summary>
    public bool? Bold { get; set; }

    /// <summary>Gets or sets whether to make font italic.</summary>
    public bool? Italic { get; set; }

    /// <summary>Gets or sets the fill/background color.</summary>
    public ChartColor? FillColor { get; set; }

    /// <summary>Gets or sets the border style.</summary>
    public ConditionalFormatBorder? Border { get; set; }

    /// <summary>Gets or sets the number format.</summary>
    public string? NumberFormat { get; set; }
}

/// <summary>Border style for conditional formatting.</summary>
public sealed class ConditionalFormatBorder
{
    /// <summary>Gets or sets the top border color.</summary>
    public ChartColor? TopColor { get; set; }

    /// <summary>Gets or sets the bottom border color.</summary>
    public ChartColor? BottomColor { get; set; }

    /// <summary>Gets or sets the left border color.</summary>
    public ChartColor? LeftColor { get; set; }

    /// <summary>Gets or sets the right border color.</summary>
    public ChartColor? RightColor { get; set; }
}

/// <summary>Color scale configuration for conditional formatting.</summary>
public sealed class ColorScale
{
    /// <summary>Gets or sets the minimum value configuration.</summary>
    public ColorScalePoint Minimum { get; set; } = new() { Type = ColorScaleValueType.MinValue, Color = ChartColor.White };

    /// <summary>Gets or sets the midpoint value configuration (optional).</summary>
    public ColorScalePoint? Midpoint { get; set; }

    /// <summary>Gets or sets the maximum value configuration.</summary>
    public ColorScalePoint Maximum { get; set; } = new() { Type = ColorScaleValueType.MaxValue, Color = ChartColor.Red };
}

/// <summary>A point in the color scale.</summary>
public sealed class ColorScalePoint
{
    /// <summary>Gets or sets the type of value.</summary>
    public ColorScaleValueType Type { get; set; } = ColorScaleValueType.MinValue;

    /// <summary>Gets or sets the numeric value (if Type is Num or Percent).</summary>
    public double? Value { get; set; }

    /// <summary>Gets or sets the formula (if Type is Formula).</summary>
    public string? Formula { get; set; }

    /// <summary>Gets or sets the color at this point.</summary>
    public ChartColor Color { get; set; } = ChartColor.White;
}

/// <summary>Types of values for color scale points.</summary>
public enum ColorScaleValueType : byte
{
    /// <summary>Numeric value.</summary>
    Num = 0,

    /// <summary>Minimum value in range.</summary>
    MinValue = 1,

    /// <summary>Maximum value in range.</summary>
    MaxValue = 2,

    /// <summary>Percentage of range.</summary>
    Percent = 3,

    /// <summary>Percentile value.</summary>
    Percentile = 4,

    /// <summary>Formula result.</summary>
    Formula = 5
}

/// <summary>Data bar configuration for conditional formatting.</summary>
public sealed class DataBar
{
    /// <summary>Gets or sets the minimum value configuration.</summary>
    public DataBarPoint Minimum { get; set; } = new() { Type = DataBarValueType.MinValue };

    /// <summary>Gets or sets the maximum value configuration.</summary>
    public DataBarPoint Maximum { get; set; } = new() { Type = DataBarValueType.MaxValue };

    /// <summary>Gets or sets the bar color.</summary>
    public ChartColor Color { get; set; } = ChartColor.Blue;

    /// <summary>Gets or sets whether to show the bar only (no value).</summary>
    public bool ShowValue { get; set; } = true;

    /// <summary>Gets or sets the bar border color (optional).</summary>
    public ChartColor? BorderColor { get; set; }

    /// <summary>Gets or sets the bar direction.</summary>
    public DataBarDirection Direction { get; set; } = DataBarDirection.LeftToRight;

    /// <summary>Gets or sets the bar axis position.</summary>
    public DataBarAxisPosition AxisPosition { get; set; } = DataBarAxisPosition.Automatic;
}

/// <summary>A point in the data bar range.</summary>
public sealed class DataBarPoint
{
    /// <summary>Gets or sets the type of value.</summary>
    public DataBarValueType Type { get; set; } = DataBarValueType.MinValue;

    /// <summary>Gets or sets the numeric value (if Type is Num or Percent).</summary>
    public double? Value { get; set; }

    /// <summary>Gets or sets the formula (if Type is Formula).</summary>
    public string? Formula { get; set; }
}

/// <summary>Types of values for data bar points.</summary>
public enum DataBarValueType : byte
{
    /// <summary>Numeric value.</summary>
    Num = 0,

    /// <summary>Minimum value in range.</summary>
    MinValue = 1,

    /// <summary>Maximum value in range.</summary>
    MaxValue = 2,

    /// <summary>Percentile value.</summary>
    Percentile = 3,

    /// <summary>Formula result.</summary>
    Formula = 4,

    /// <summary>Automatic value.</summary>
    Auto = 5
}

/// <summary>Data bar direction.</summary>
public enum DataBarDirection : byte
{
    /// <summary>Left to right.</summary>
    LeftToRight = 0,

    /// <summary>Right to left.</summary>
    RightToLeft = 1
}

/// <summary>Data bar axis position.</summary>
public enum DataBarAxisPosition : byte
{
    /// <summary>Automatic position.</summary>
    Automatic = 0,

    /// <summary>Middle of cell.</summary>
    Middle = 1,

    /// <summary>No axis.</summary>
    None = 2
}

/// <summary>Icon set configuration for conditional formatting.</summary>
public sealed class IconSet
{
    /// <summary>Gets or sets the type of icon set.</summary>
    public IconSetType Type { get; set; } = IconSetType.ThreeTrafficLights;

    /// <summary>Gets or sets whether to show icons only (no values).</summary>
    public bool ShowValue { get; set; } = true;

    /// <summary>Gets or sets whether to reverse the icon order.</summary>
    public bool Reverse { get; set; }

    /// <summary>Gets or sets custom threshold values (optional).</summary>
    public List<IconThreshold>? Thresholds { get; set; }
}

/// <summary>Types of icon sets.</summary>
public enum IconSetType : byte
{
    /// <summary>3 arrows.</summary>
    ThreeArrows = 0,

    /// <summary>3 gray arrows.</summary>
    ThreeArrowsGray = 1,

    /// <summary>3 flags.</summary>
    ThreeFlags = 2,

    /// <summary>3 traffic lights.</summary>
    ThreeTrafficLights = 3,

    /// <summary>3 signs.</summary>
    ThreeSigns = 4,

    /// <summary>3 symbols.</summary>
    ThreeSymbols = 5,

    /// <summary>3 symbols (circled).</summary>
    ThreeSymbols2 = 6,

    /// <summary>4 arrows.</summary>
    FourArrows = 7,

    /// <summary>4 gray arrows.</summary>
    FourArrowsGray = 8,

    /// <summary>4 red to black.</summary>
    FourRedToBlack = 9,

    /// <summary>4 ratings.</summary>
    FourRatings = 10,

    /// <summary>4 traffic lights.</summary>
    FourTrafficLights = 11,

    /// <summary>5 arrows.</summary>
    FiveArrows = 12,

    /// <summary>5 gray arrows.</summary>
    FiveArrowsGray = 13,

    /// <summary>5 ratings.</summary>
    FiveRatings = 14,

    /// <summary>5 quarters.</summary>
    FiveQuarters = 15,

    /// <summary>3 stars.</summary>
    ThreeStars = 16,

    /// <summary>3 triangles.</summary>
    ThreeTriangles = 17,

    /// <summary>5 boxes.</summary>
    FiveBoxes = 18
}

/// <summary>Threshold for icon set.</summary>
public sealed class IconThreshold
{
    /// <summary>Gets or sets the type of threshold value.</summary>
    public IconValueType Type { get; set; } = IconValueType.Percent;

    /// <summary>Gets or sets the threshold value.</summary>
    public double Value { get; set; }

    /// <summary>Gets or sets the formula (if Type is Formula).</summary>
    public string? Formula { get; set; }

    /// <summary>Gets or sets the operator for comparison.</summary>
    public IconOperator Operator { get; set; } = IconOperator.GreaterThanOrEqual;
}

/// <summary>Types of values for icon thresholds.</summary>
public enum IconValueType : byte
{
    /// <summary>Numeric value.</summary>
    Num = 0,

    /// <summary>Percentage of range.</summary>
    Percent = 1,

    /// <summary>Formula result.</summary>
    Formula = 2,

    /// <summary>Percentile value.</summary>
    Percentile = 3
}

/// <summary>Operators for icon threshold comparison.</summary>
public enum IconOperator : byte
{
    /// <summary>Greater than.</summary>
    GreaterThan = 0,

    /// <summary>Greater than or equal to.</summary>
    GreaterThanOrEqual = 1
}
