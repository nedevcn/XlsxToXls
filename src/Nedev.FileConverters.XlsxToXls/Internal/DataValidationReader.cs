using System.IO.Compression;
using System.Xml;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Reads data validation rules from XLSX files.
/// Parses dataValidations and dataValidation elements.
/// </summary>
public static class DataValidationReader
{
    private static readonly XmlReaderSettings Settings = new()
    {
        IgnoreWhitespace = true,
        IgnoreComments = true
    };

    /// <summary>
    /// Reads all data validation rules from the worksheet XML.
    /// </summary>
    public static List<DataValidationData> ReadDataValidations(ZipArchive archive, string worksheetPath, Action<string>? log = null)
    {
        var validations = new List<DataValidationData>();

        try
        {
            var entry = archive.GetEntry(worksheetPath) ?? archive.GetEntry("xl/" + worksheetPath);
            if (entry == null)
            {
                log?.Invoke($"[DataValidationReader] Worksheet not found: {worksheetPath}");
                return validations;
            }

            log?.Invoke($"[DataValidationReader] Reading data validations from: {worksheetPath}");

            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, Settings);
            var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element &&
                    reader.LocalName == "dataValidations" &&
                    reader.NamespaceURI == ns)
                {
                    // Read all dataValidation elements
                    var depth = 1;
                    while (reader.Read() && depth > 0)
                    {
                        if (reader.NodeType == XmlNodeType.Element &&
                            reader.LocalName == "dataValidation" &&
                            reader.NamespaceURI == ns)
                        {
                            var validation = ReadDataValidation(reader, ns, log);
                            if (validation != null)
                            {
                                validations.Add(validation);
                                log?.Invoke($"[DataValidationReader] Found validation: {validation.Type}");
                            }
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement &&
                                 reader.LocalName == "dataValidations" &&
                                 reader.NamespaceURI == ns)
                        {
                            depth--;
                        }
                    }
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[DataValidationReader] Error reading data validations: {ex.Message}");
        }

        return validations;
    }

    private static DataValidationData? ReadDataValidation(XmlReader reader, string ns, Action<string>? log)
    {
        try
        {
            var validation = new DataValidationData();

            // Read attributes
            if (reader.GetAttribute("type") is { } typeStr)
            {
                validation.Type = typeStr.ToLowerInvariant() switch
                {
                    "whole" => ValidationType.Whole,
                    "decimal" => ValidationType.Decimal,
                    "list" => ValidationType.List,
                    "date" => ValidationType.Date,
                    "time" => ValidationType.Time,
                    "textLength" => ValidationType.TextLength,
                    "custom" => ValidationType.Custom,
                    _ => ValidationType.Any
                };
            }

            if (reader.GetAttribute("operator") is { } operatorStr)
            {
                validation.Operator = operatorStr.ToLowerInvariant() switch
                {
                    "between" => ValidationOperator.Between,
                    "notBetween" => ValidationOperator.NotBetween,
                    "equal" => ValidationOperator.Equal,
                    "notEqual" => ValidationOperator.NotEqual,
                    "greaterThan" => ValidationOperator.GreaterThan,
                    "lessThan" => ValidationOperator.LessThan,
                    "greaterThanOrEqual" => ValidationOperator.GreaterThanOrEqual,
                    "lessThanOrEqual" => ValidationOperator.LessThanOrEqual,
                    _ => ValidationOperator.Between
                };
            }

            if (reader.GetAttribute("allowBlank") is { } allowBlankStr &&
                bool.TryParse(allowBlankStr, out var allowBlank))
            {
                validation.AllowBlank = allowBlank;
            }

            if (reader.GetAttribute("showDropDown") is { } showDropDownStr &&
                bool.TryParse(showDropDownStr, out var showDropDown))
            {
                validation.SuppressDropDown = !showDropDown;
            }

            if (reader.GetAttribute("showInputMessage") is { } showInputStr &&
                bool.TryParse(showInputStr, out var showInput))
            {
                validation.ShowInputMessage = showInput;
            }

            if (reader.GetAttribute("showErrorMessage") is { } showErrorStr &&
                bool.TryParse(showErrorStr, out var showError))
            {
                validation.ShowErrorMessage = showError;
            }

            if (reader.GetAttribute("errorStyle") is { } errorStyleStr)
            {
                validation.ErrorAlertType = errorStyleStr.ToLowerInvariant() switch
                {
                    "warning" => ErrorAlertType.Warning,
                    "information" => ErrorAlertType.Information,
                    _ => ErrorAlertType.Stop
                };
            }

            // Read sqref (cell ranges)
            if (reader.GetAttribute("sqref") is { } sqref)
            {
                validation.Ranges = ParseSqref(sqref, log);
                log?.Invoke($"[DataValidationReader] Validation ranges: {sqref}");
            }

            // Read child elements
            if (!reader.IsEmptyElement)
            {
                var depth = 1;
                while (reader.Read() && depth > 0)
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
                    {
                        switch (reader.LocalName)
                        {
                            case "formula1":
                                validation.Formula1 = reader.ReadElementContentAsString();
                                depth--; // ReadElementContentAsString advances past the end element
                                break;

                            case "formula2":
                                validation.Formula2 = reader.ReadElementContentAsString();
                                depth--;
                                break;

                            case "prompt":
                                validation.InputMessage = reader.ReadElementContentAsString();
                                depth--;
                                break;

                            case "promptTitle":
                                validation.InputTitle = reader.ReadElementContentAsString();
                                depth--;
                                break;

                            case "error":
                                validation.ErrorMessage = reader.ReadElementContentAsString();
                                depth--;
                                break;

                            case "errorTitle":
                                validation.ErrorTitle = reader.ReadElementContentAsString();
                                depth--;
                                break;
                        }
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement)
                    {
                        if (reader.LocalName == "dataValidation" && reader.NamespaceURI == ns)
                        {
                            depth--;
                        }
                    }
                }
            }

            // Parse list values if it's a list validation
            if (validation.Type == ValidationType.List && !string.IsNullOrEmpty(validation.Formula1))
            {
                validation.ListValues = ParseListValues(validation.Formula1);
            }

            return validation;
        }
        catch (Exception ex)
        {
            log?.Invoke($"[DataValidationReader] Error reading data validation: {ex.Message}");
            return null;
        }
    }

    private static List<CellRange> ParseSqref(string sqref, Action<string>? log)
    {
        var ranges = new List<CellRange>();

        try
        {
            // sqref can contain multiple ranges separated by spaces
            var rangeStrings = sqref.Split([' '], StringSplitOptions.RemoveEmptyEntries);

            foreach (var rangeStr in rangeStrings)
            {
                var range = ParseRange(rangeStr.Trim());
                if (range != null)
                {
                    ranges.Add(range);
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[DataValidationReader] Error parsing sqref '{sqref}': {ex.Message}");
        }

        return ranges;
    }

    private static CellRange? ParseRange(string rangeStr)
    {
        // Parse range like "A1:B10" or single cell like "A1"
        try
        {
            if (rangeStr.Contains(':'))
            {
                var parts = rangeStr.Split(':', 2);
                if (parts.Length == 2)
                {
                    var (col1, row1) = ParseCellReference(parts[0]);
                    var (col2, row2) = ParseCellReference(parts[1]);

                    if (col1.HasValue && row1.HasValue && col2.HasValue && row2.HasValue)
                    {
                        return new CellRange
                        {
                            FirstRow = row1.Value,
                            FirstCol = col1.Value,
                            LastRow = row2.Value,
                            LastCol = col2.Value
                        };
                    }
                }
            }
            else
            {
                // Single cell
                var (col, row) = ParseCellReference(rangeStr);
                if (col.HasValue && row.HasValue)
                {
                    return new CellRange
                    {
                        FirstRow = row.Value,
                        FirstCol = col.Value,
                        LastRow = row.Value,
                        LastCol = col.Value
                    };
                }
            }
        }
        catch
        {
            // Ignore parsing errors
        }

        return null;
    }

    private static (int? col, int? row) ParseCellReference(string cellRef)
    {
        var colStr = "";
        var rowStr = "";

        foreach (var c in cellRef)
        {
            if (char.IsLetter(c))
            {
                colStr += c;
            }
            else if (char.IsDigit(c))
            {
                rowStr += c;
            }
        }

        int? col = null;
        int? row = null;

        if (!string.IsNullOrEmpty(colStr))
        {
            col = 0;
            foreach (var c in colStr.ToUpperInvariant())
            {
                col = col * 26 + (c - 'A' + 1);
            }
            col--; // Convert to 0-based
        }

        if (!string.IsNullOrEmpty(rowStr) && int.TryParse(rowStr, out var rowNum))
        {
            row = rowNum - 1; // Convert to 0-based
        }

        return (col, row);
    }

    private static List<string> ParseListValues(string formula)
    {
        var values = new List<string>();

        try
        {
            // List formula can be:
            // 1. Comma-separated values: "Value1,Value2,Value3"
            // 2. Cell reference: "Sheet1!$A$1:$A$5"
            // 3. Named range: "MyList"

            if (formula.StartsWith("\"") && formula.EndsWith("\""))
            {
                // Quoted list
                var content = formula.Trim('"');
                values.AddRange(content.Split(','));
            }
            else if (!formula.Contains('!') && !formula.Contains(':'))
            {
                // Direct comma-separated values
                values.AddRange(formula.Split(','));
            }
            else
            {
                // Cell reference or named range - store as single value
                values.Add(formula);
            }
        }
        catch
        {
            // If parsing fails, store the whole formula
            values.Add(formula);
        }

        return values;
    }
}
