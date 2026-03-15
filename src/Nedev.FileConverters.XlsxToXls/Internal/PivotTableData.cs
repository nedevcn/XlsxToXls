namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Represents a PivotTable for BIFF8 format conversion.
/// Contains all information needed to create and configure a pivot table.
/// </summary>
public sealed class PivotTableData
{
    /// <summary>Gets or sets the unique identifier for this pivot table.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>Gets or sets the name of the pivot table.</summary>
    public string Name { get; set; } = "PivotTable1";

    /// <summary>Gets or sets the cache ID this pivot table references.</summary>
    public int CacheId { get; set; }

    /// <summary>Gets or sets the location (top-left cell) of the pivot table.</summary>
    public CellLocation Location { get; set; } = new();

    /// <summary>Gets or sets the row fields.</summary>
    public List<PivotField> RowFields { get; set; } = [];

    /// <summary>Gets or sets the column fields.</summary>
    public List<PivotField> ColumnFields { get; set; } = [];

    /// <summary>Gets or sets the data fields (values).</summary>
    public List<PivotDataField> DataFields { get; set; } = [];

    /// <summary>Gets or sets the page fields (filters).</summary>
    public List<PivotField> PageFields { get; set; } = [];

    /// <summary>Gets or sets the hidden fields.</summary>
    public List<PivotField> HiddenFields { get; set; } = [];

    /// <summary>Gets or sets whether to show grand totals for rows.</summary>
    public bool ShowRowGrandTotals { get; set; } = true;

    /// <summary>Gets or sets whether to show grand totals for columns.</summary>
    public bool ShowColumnGrandTotals { get; set; } = true;

    /// <summary>Gets or sets whether to show error messages in data cells.</summary>
    public bool ShowError { get; set; }

    /// <summary>Gets or sets the error message to display.</summary>
    public string? ErrorString { get; set; }

    /// <summary>Gets or sets whether to show empty cells as custom string.</summary>
    public bool ShowEmpty { get; set; }

    /// <summary>Gets or sets the string to display for empty cells.</summary>
    public string? EmptyString { get; set; }

    /// <summary>Gets or sets whether to auto format the pivot table.</summary>
    public bool AutoFormat { get; set; } = true;

    /// <summary>Gets or sets whether to preserve formatting.</summary>
    public bool PreserveFormatting { get; set; } = true;

    /// <summary>Gets or sets whether to use custom lists for sorting.</summary>
    public bool UseCustomLists { get; set; } = true;

    /// <summary>Gets or sets whether to show expand/collapse buttons.</summary>
    public bool ShowExpandCollapseButtons { get; set; } = true;

    /// <summary>Gets or sets whether to show field headers.</summary>
    public bool ShowFieldHeaders { get; set; } = true;

    /// <summary>Gets or sets whether to show item names in outline form.</summary>
    public bool OutlineForm { get; set; }

    /// <summary>Gets or sets whether to compact row axis.</summary>
    public bool CompactRowAxis { get; set; } = true;

    /// <summary>Gets or sets whether to compact column axis.</summary>
    public bool CompactColumnAxis { get; set; } = true;

    /// <summary>Gets or sets the outline indentation.</summary>
    public byte OutlineIndent { get; set; } = 1;

    /// <summary>Gets or sets the pivot table style name.</summary>
    public string? StyleName { get; set; }

    /// <summary>Gets or sets the merge labels option.</summary>
    public MergeLabels MergeLabels { get; set; } = MergeLabels.None;

    /// <summary>Gets or sets the page wrap count.</summary>
    public int PageWrap { get; set; }

    /// <summary>Gets or sets the page filter order.</summary>
    public PageFilterOrder PageFilterOrder { get; set; } = PageFilterOrder.DownThenOver;

    /// <summary>Gets or sets the source data range.</summary>
    public CellRange? SourceRange { get; set; }

    /// <summary>Gets or sets the target sheet name.</summary>
    public string? TargetSheet { get; set; }
}

/// <summary>
/// Represents a pivot field (row, column, or page field).
/// </summary>
public sealed class PivotField
{
    /// <summary>Gets or sets the field index in the cache.</summary>
    public int FieldIndex { get; set; }

    /// <summary>Gets or sets the axis this field belongs to.</summary>
    public PivotAxis Axis { get; set; } = PivotAxis.Row;

    /// <summary>Gets or sets the subtotal type.</summary>
    public SubtotalType Subtotal { get; set; } = SubtotalType.None;

    /// <summary>Gets or sets whether to show subtotals at the top.</summary>
    public bool SubtotalTop { get; set; } = true;

    /// <summary>Gets or sets whether to show items with no data.</summary>
    public bool ShowAllItems { get; set; }

    /// <summary>Gets or sets whether to insert blank rows after each item.</summary>
    public bool InsertBlankRows { get; set; }

    /// <summary>Gets or sets whether to insert page breaks between items.</summary>
    public bool InsertPageBreaks { get; set; }

    /// <summary>Gets or sets the sort order.</summary>
    public SortOrder SortOrder { get; set; } = SortOrder.Ascending;

    /// <summary>Gets or sets whether to auto sort.</summary>
    public bool AutoSort { get; set; }

    /// <summary>Gets or sets the auto sort field.</summary>
    public int? AutoSortField { get; set; }

    /// <summary>Gets or sets whether to auto show.</summary>
    public bool AutoShow { get; set; }

    /// <summary>Gets or sets the auto show count.</summary>
    public int AutoShowCount { get; set; } = 10;

    /// <summary>Gets or sets the auto show type.</summary>
    public AutoShowType AutoShowType { get; set; } = AutoShowType.Top;

    /// <summary>Gets or sets the auto show field.</summary>
    public int? AutoShowField { get; set; }

    /// <summary>Gets or sets the hidden items.</summary>
    public List<int> HiddenItems { get; set; } = [];

    /// <summary>Gets or sets the field name.</summary>
    public string? Name { get; set; }

    /// <summary>Gets or sets the number format index.</summary>
    public int? NumberFormat { get; set; }

    /// <summary>Gets or sets the outline level.</summary>
    public byte OutlineLevel { get; set; }

    /// <summary>Gets or sets whether this is a compact field.</summary>
    public bool Compact { get; set; } = true;
}

/// <summary>
/// Represents a pivot data field (value field).
/// </summary>
public sealed class PivotDataField
{
    /// <summary>Gets or sets the field index in the cache.</summary>
    public int FieldIndex { get; set; }

    /// <summary>Gets or sets the aggregation function.</summary>
    public AggregationFunction Function { get; set; } = AggregationFunction.Sum;

    /// <summary>Gets or sets the name to display.</summary>
    public string? Name { get; set; }

    /// <summary>Gets or sets the number format index.</summary>
    public int? NumberFormat { get; set; }

    /// <summary>Gets or sets the base field for calculations.</summary>
    public int? BaseField { get; set; }

    /// <summary>Gets or sets the base item for calculations.</summary>
    public int? BaseItem { get; set; }

    /// <summary>Gets or sets the show data as type.</summary>
    public ShowDataAs ShowDataAs { get; set; } = ShowDataAs.Normal;

    /// <summary>Gets or sets whether to show values as percentage.</summary>
    public bool ShowAsPercentage { get; set; }

    /// <summary>Gets or sets the position in the data fields list.</summary>
    public int Position { get; set; }
}

/// <summary>
/// Represents a pivot cache definition.
/// </summary>
public sealed class PivotCacheDefinition
{
    /// <summary>Gets or sets the cache ID.</summary>
    public int CacheId { get; set; }

    /// <summary>Gets or sets the source data range.</summary>
    public CellRange? SourceRange { get; set; }

    /// <summary>Gets or sets the source sheet name.</summary>
    public string? SourceSheet { get; set; }

    /// <summary>Gets or sets the cache fields.</summary>
    public List<PivotCacheField> Fields { get; set; } = [];

    /// <summary>Gets or sets the number of records.</summary>
    public int RecordCount { get; set; }

    /// <summary>Gets or sets whether the cache is refreshed on open.</summary>
    public bool RefreshOnLoad { get; set; }

    /// <summary>Gets or sets the created version.</summary>
    public int CreatedVersion { get; set; } = 3;

    /// <summary>Gets or sets the refreshed version.</summary>
    public int RefreshedVersion { get; set; } = 3;

    /// <summary>Gets or sets the minimum refresh version.</summary>
    public int MinRefreshableVersion { get; set; } = 3;
}

/// <summary>
/// Represents a field in the pivot cache.
/// </summary>
public sealed class PivotCacheField
{
    /// <summary>Gets or sets the field name.</summary>
    public string Name { get; set; } = "";

    /// <summary>Gets or sets the field type.</summary>
    public CacheFieldType Type { get; set; } = CacheFieldType.String;

    /// <summary>Gets or sets the shared items (unique values).</summary>
    public List<string> SharedItems { get; set; } = [];

    /// <summary>Gets or sets the number format index.</summary>
    public int? NumberFormat { get; set; }

    /// <summary>Gets or sets whether this field contains mixed types.</summary>
    public bool MixedTypes { get; set; }

    /// <summary>Gets or sets the number of items.</summary>
    public int ItemCount => SharedItems.Count;
}

/// <summary>
/// Represents a cell location.
/// </summary>
public sealed class CellLocation
{
    /// <summary>Gets or sets the row index.</summary>
    public int Row { get; set; }

    /// <summary>Gets or sets the column index.</summary>
    public int Column { get; set; }

    /// <summary>Creates a cell location from row and column.</summary>
    public static CellLocation FromCell(int row, int col) => new() { Row = row, Column = col };
}

/// <summary>Pivot table axis.</summary>
public enum PivotAxis : byte
{
    /// <summary>Row axis.</summary>
    Row = 0,

    /// <summary>Column axis.</summary>
    Column = 1,

    /// <summary>Page axis (filter).</summary>
    Page = 2,

    /// <summary>Data axis (values).</summary>
    Data = 3,

    /// <summary>Hidden axis.</summary>
    Hidden = 4
}

/// <summary>Aggregation functions for data fields.</summary>
public enum AggregationFunction : byte
{
    /// <summary>Average.</summary>
    Average = 0,

    /// <summary>Count.</summary>
    Count = 1,

    /// <summary>Count numbers only.</summary>
    CountNums = 2,

    /// <summary>Maximum.</summary>
    Max = 3,

    /// <summary>Minimum.</summary>
    Min = 4,

    /// <summary>Product.</summary>
    Product = 5,

    /// <summary>Standard deviation.</summary>
    StdDev = 6,

    /// <summary>Standard deviation of population.</summary>
    StdDevP = 7,

    /// <summary>Sum.</summary>
    Sum = 8,

    /// <summary>Variance.</summary>
    Var = 9,

    /// <summary>Variance of population.</summary>
    VarP = 10
}

/// <summary>Subtotal types.</summary>
public enum SubtotalType : ushort
{
    /// <summary>No subtotal.</summary>
    None = 0,

    /// <summary>Default subtotal.</summary>
    Default = 1,

    /// <summary>Sum.</summary>
    Sum = 2,

    /// <summary>Count.</summary>
    Count = 4,

    /// <summary>Average.</summary>
    Average = 8,

    /// <summary>Maximum.</summary>
    Max = 16,

    /// <summary>Minimum.</summary>
    Min = 32,

    /// <summary>Product.</summary>
    Product = 64,

    /// <summary>Count numbers.</summary>
    CountNums = 128,

    /// <summary>Standard deviation.</summary>
    StdDev = 256,

    /// <summary>Standard deviation of population.</summary>
    StdDevP = 512,

    /// <summary>Variance.</summary>
    Var = 1024,

    /// <summary>Variance of population.</summary>
    VarP = 2048
}

/// <summary>Show data as options.</summary>
public enum ShowDataAs : byte
{
    /// <summary>Normal values.</summary>
    Normal = 0,

    /// <summary>Difference from.</summary>
    Difference = 1,

    /// <summary>Percentage of.</summary>
    PercentOf = 2,

    /// <summary>Percentage difference from.</summary>
    PercentDiff = 3,

    /// <summary>Running total in.</summary>
    RunTotal = 4,

    /// <summary>Percentage of row.</summary>
    PercentOfRow = 5,

    /// <summary>Percentage of column.</summary>
    PercentOfCol = 6,

    /// <summary>Percentage of total.</summary>
    PercentOfTotal = 7,

    /// <summary>Index.</summary>
    Index = 8
}

/// <summary>Sort order.</summary>
public enum SortOrder : byte
{
    /// <summary>Ascending.</summary>
    Ascending = 0,

    /// <summary>Descending.</summary>
    Descending = 1
}

/// <summary>Auto show type.</summary>
public enum AutoShowType : byte
{
    /// <summary>Show top items.</summary>
    Top = 0,

    /// <summary>Show bottom items.</summary>
    Bottom = 1
}

/// <summary>Merge labels options.</summary>
public enum MergeLabels : byte
{
    /// <summary>Do not merge labels.</summary>
    None = 0,

    /// <summary>Merge row labels.</summary>
    Row = 1,

    /// <summary>Merge column labels.</summary>
    Column = 2,

    /// <summary>Merge both.</summary>
    Both = 3
}

/// <summary>Page filter order.</summary>
public enum PageFilterOrder : byte
{
    /// <summary>Down then over.</summary>
    DownThenOver = 0,

    /// <summary>Over then down.</summary>
    OverThenDown = 1
}

/// <summary>Cache field type.</summary>
public enum CacheFieldType : byte
{
    /// <summary>String type.</summary>
    String = 0,

    /// <summary>Numeric type.</summary>
    Numeric = 1,

    /// <summary>Integer type.</summary>
    Integer = 2,

    /// <summary>Boolean type.</summary>
    Boolean = 3,

    /// <summary>Date type.</summary>
    Date = 4
}
