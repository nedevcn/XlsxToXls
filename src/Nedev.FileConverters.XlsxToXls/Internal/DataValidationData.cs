namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Represents data validation rules for BIFF8 format conversion.
/// Contains all information needed to validate cell input.
/// </summary>
public sealed class DataValidationData
{
    /// <summary>Gets or sets the unique identifier for this validation rule.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>Gets or sets the cell ranges this validation applies to.</summary>
    public List<CellRange> Ranges { get; set; } = [];

    /// <summary>Gets or sets the type of data validation.</summary>
    public ValidationType Type { get; set; } = ValidationType.Any;

    /// <summary>Gets or sets the operator for value comparison.</summary>
    public ValidationOperator Operator { get; set; } = ValidationOperator.Between;

    /// <summary>Gets or sets the first formula/value for comparison.</summary>
    public string? Formula1 { get; set; }

    /// <summary>Gets or sets the second formula/value for comparison (for between operators).</summary>
    public string? Formula2 { get; set; }

    /// <summary>Gets or sets whether to allow empty cells. Default is true.</summary>
    public bool AllowBlank { get; set; } = true;

    /// <summary>Gets or sets whether to suppress the drop-down arrow for list validations.</summary>
    public bool SuppressDropDown { get; set; }

    /// <summary>Gets or sets whether to show the input message.</summary>
    public bool ShowInputMessage { get; set; } = true;

    /// <summary>Gets or sets whether to show the error alert.</summary>
    public bool ShowErrorMessage { get; set; } = true;

    /// <summary>Gets or sets the title of the input message dialog.</summary>
    public string? InputTitle { get; set; }

    /// <summary>Gets or sets the text of the input message.</summary>
    public string? InputMessage { get; set; }

    /// <summary>Gets or sets the title of the error alert dialog.</summary>
    public string? ErrorTitle { get; set; }

    /// <summary>Gets or sets the text of the error message.</summary>
    public string? ErrorMessage { get; set; }

    /// <summary>Gets or sets the type of error alert. Default is Stop.</summary>
    public ErrorAlertType ErrorAlertType { get; set; } = ErrorAlertType.Stop;

    /// <summary>Gets or sets the list of values for list validation.</summary>
    public List<string>? ListValues { get; set; }

    /// <summary>Gets or sets whether the list allows multiple selections.</summary>
    public bool AllowMultiSelect { get; set; }

    /// <summary>Gets or sets whether to show error alert for empty cells.</summary>
    public bool ShowErrorOnBlank { get; set; }
}

/// <summary>Types of data validation.</summary>
public enum ValidationType : byte
{
    /// <summary>Any value is allowed.</summary>
    Any = 0,

    /// <summary>Whole numbers only.</summary>
    Whole = 1,

    /// <summary>Decimal numbers only.</summary>
    Decimal = 2,

    /// <summary>Values from a list.</summary>
    List = 3,

    /// <summary>Date values only.</summary>
    Date = 4,

    /// <summary>Time values only.</summary>
    Time = 5,

    /// <summary>Text length restrictions.</summary>
    TextLength = 6,

    /// <summary>Custom formula validation.</summary>
    Custom = 7
}

/// <summary>Validation operators for value comparison.</summary>
public enum ValidationOperator : byte
{
    /// <summary>No comparison (used with Any type).</summary>
    None = 0,

    /// <summary>Value must be between two values.</summary>
    Between = 1,

    /// <summary>Value must not be between two values.</summary>
    NotBetween = 2,

    /// <summary>Value must be equal to specified value.</summary>
    Equal = 3,

    /// <summary>Value must not be equal to specified value.</summary>
    NotEqual = 4,

    /// <summary>Value must be greater than specified value.</summary>
    GreaterThan = 5,

    /// <summary>Value must be less than specified value.</summary>
    LessThan = 6,

    /// <summary>Value must be greater than or equal to specified value.</summary>
    GreaterThanOrEqual = 7,

    /// <summary>Value must be less than or equal to specified value.</summary>
    LessThanOrEqual = 8
}

/// <summary>Types of error alerts.</summary>
public enum ErrorAlertType : byte
{
    /// <summary>Stop alert (must correct to continue).</summary>
    Stop = 0,

    /// <summary>Warning alert (can continue with invalid data).</summary>
    Warning = 1,

    /// <summary>Information alert (for information only).</summary>
    Information = 2
}

/// <summary>
/// Helper methods for data validation.
/// </summary>
public static class DataValidationHelper
{
    /// <summary>
    /// Creates a whole number validation rule.
    /// </summary>
    public static DataValidationData CreateWholeNumberValidation(
        List<CellRange> ranges,
        ValidationOperator op,
        int? min = null,
        int? max = null)
    {
        var validation = new DataValidationData
        {
            Type = ValidationType.Whole,
            Operator = op,
            Ranges = ranges
        };

        if (min.HasValue)
        {
            validation.Formula1 = min.Value.ToString();
        }

        if (max.HasValue)
        {
            validation.Formula2 = max.Value.ToString();
        }

        return validation;
    }

    /// <summary>
    /// Creates a decimal number validation rule.
    /// </summary>
    public static DataValidationData CreateDecimalValidation(
        List<CellRange> ranges,
        ValidationOperator op,
        double? min = null,
        double? max = null)
    {
        var validation = new DataValidationData
        {
            Type = ValidationType.Decimal,
            Operator = op,
            Ranges = ranges
        };

        if (min.HasValue)
        {
            validation.Formula1 = min.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        if (max.HasValue)
        {
            validation.Formula2 = max.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        return validation;
    }

    /// <summary>
    /// Creates a list validation rule.
    /// </summary>
    public static DataValidationData CreateListValidation(
        List<CellRange> ranges,
        List<string> values,
        bool allowBlank = true,
        bool suppressDropDown = false)
    {
        return new DataValidationData
        {
            Type = ValidationType.List,
            Operator = ValidationOperator.None,
            Ranges = ranges,
            ListValues = values,
            AllowBlank = allowBlank,
            SuppressDropDown = suppressDropDown,
            Formula1 = string.Join("\0", values)
        };
    }

    /// <summary>
    /// Creates a date validation rule.
    /// </summary>
    public static DataValidationData CreateDateValidation(
        List<CellRange> ranges,
        ValidationOperator op,
        DateTime? startDate = null,
        DateTime? endDate = null)
    {
        var validation = new DataValidationData
        {
            Type = ValidationType.Date,
            Operator = op,
            Ranges = ranges
        };

        if (startDate.HasValue)
        {
            validation.Formula1 = startDate.Value.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        if (endDate.HasValue)
        {
            validation.Formula2 = endDate.Value.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        return validation;
    }

    /// <summary>
    /// Creates a time validation rule.
    /// </summary>
    public static DataValidationData CreateTimeValidation(
        List<CellRange> ranges,
        ValidationOperator op,
        TimeSpan? startTime = null,
        TimeSpan? endTime = null)
    {
        var validation = new DataValidationData
        {
            Type = ValidationType.Time,
            Operator = op,
            Ranges = ranges
        };

        if (startTime.HasValue)
        {
            // Convert TimeSpan to Excel time value (fraction of day)
            var excelTime = startTime.Value.TotalDays;
            validation.Formula1 = excelTime.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        if (endTime.HasValue)
        {
            var excelTime = endTime.Value.TotalDays;
            validation.Formula2 = excelTime.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        return validation;
    }

    /// <summary>
    /// Creates a text length validation rule.
    /// </summary>
    public static DataValidationData CreateTextLengthValidation(
        List<CellRange> ranges,
        ValidationOperator op,
        int? minLength = null,
        int? maxLength = null)
    {
        var validation = new DataValidationData
        {
            Type = ValidationType.TextLength,
            Operator = op,
            Ranges = ranges
        };

        if (minLength.HasValue)
        {
            validation.Formula1 = minLength.Value.ToString();
        }

        if (maxLength.HasValue)
        {
            validation.Formula2 = maxLength.Value.ToString();
        }

        return validation;
    }

    /// <summary>
    /// Creates a custom formula validation rule.
    /// </summary>
    public static DataValidationData CreateCustomValidation(
        List<CellRange> ranges,
        string formula)
    {
        return new DataValidationData
        {
            Type = ValidationType.Custom,
            Operator = ValidationOperator.None,
            Ranges = ranges,
            Formula1 = formula
        };
    }

    /// <summary>
    /// Sets the input message for a validation rule.
    /// </summary>
    public static DataValidationData WithInputMessage(
        this DataValidationData validation,
        string title,
        string message)
    {
        validation.InputTitle = title;
        validation.InputMessage = message;
        validation.ShowInputMessage = true;
        return validation;
    }

    /// <summary>
    /// Sets the error message for a validation rule.
    /// </summary>
    public static DataValidationData WithErrorMessage(
        this DataValidationData validation,
        string title,
        string message,
        ErrorAlertType alertType = ErrorAlertType.Stop)
    {
        validation.ErrorTitle = title;
        validation.ErrorMessage = message;
        validation.ErrorAlertType = alertType;
        validation.ShowErrorMessage = true;
        return validation;
    }

    /// <summary>
    /// Converts a cell reference to R1C1 style formula.
    /// </summary>
    public static string ToR1C1Reference(int row, int col)
    {
        return $"R{row + 1}C{col + 1}";
    }

    /// <summary>
    /// Converts a cell range to R1C1 style formula.
    /// </summary>
    public static string ToR1C1Range(CellRange range)
    {
        if (range.FirstRow == range.LastRow && range.FirstCol == range.LastCol)
        {
            return ToR1C1Reference(range.FirstRow, range.FirstCol);
        }

        var start = ToR1C1Reference(range.FirstRow, range.FirstCol);
        var end = ToR1C1Reference(range.LastRow, range.LastCol);
        return $"{start}:{end}";
    }
}
