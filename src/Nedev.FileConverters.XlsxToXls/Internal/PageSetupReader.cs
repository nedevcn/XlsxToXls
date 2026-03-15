using System.IO.Compression;
using System.Xml;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Reads page setup information from XLSX files.
/// Parses pageSetup, printOptions, and definedName elements.
/// </summary>
public static class PageSetupReader
{
    private static readonly XmlReaderSettings Settings = new()
    {
        IgnoreWhitespace = true,
        IgnoreComments = true
    };

    /// <summary>
    /// Reads page setup data from the worksheet XML.
    /// </summary>
    public static PageSetupData? ReadPageSetup(ZipArchive archive, string worksheetPath, Action<string>? log = null)
    {
        try
        {
            var entry = archive.GetEntry(worksheetPath) ?? archive.GetEntry("xl/" + worksheetPath);
            if (entry == null)
            {
                log?.Invoke($"[PageSetupReader] Worksheet not found: {worksheetPath}");
                return null;
            }

            log?.Invoke($"[PageSetupReader] Reading page setup from: {worksheetPath}");

            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, Settings);
            var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            var pageSetup = new PageSetupData();
            var foundPageSetup = false;

            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element) continue;

                switch (reader.LocalName)
                {
                    case "pageSetup":
                        if (reader.NamespaceURI == ns)
                        {
                            ReadPageSetupElement(reader, pageSetup, log);
                            foundPageSetup = true;
                        }
                        break;

                    case "printOptions":
                        if (reader.NamespaceURI == ns)
                        {
                            ReadPrintOptions(reader, pageSetup, log);
                        }
                        break;

                    case "pageMargins":
                        if (reader.NamespaceURI == ns)
                        {
                            ReadPageMargins(reader, pageSetup, log);
                        }
                        break;
                }
            }

            if (!foundPageSetup)
            {
                log?.Invoke("[PageSetupReader] No pageSetup element found, using defaults");
            }

            return pageSetup;
        }
        catch (Exception ex)
        {
            log?.Invoke($"[PageSetupReader] Error reading page setup: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Reads print area and print titles from workbook defined names.
    /// </summary>
    public static void ReadPrintRanges(ZipArchive archive, PageSetupData pageSetup, string sheetName, Action<string>? log = null)
    {
        try
        {
            var entry = archive.GetEntry("xl/workbook.xml");
            if (entry == null)
            {
                log?.Invoke("[PageSetupReader] Workbook.xml not found");
                return;
            }

            log?.Invoke($"[PageSetupReader] Reading print ranges for sheet: {sheetName}");

            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, Settings);
            var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element &&
                    reader.LocalName == "definedName" &&
                    reader.NamespaceURI == ns)
                {
                    var name = reader.GetAttribute("name");
                    var localSheetId = reader.GetAttribute("localSheetId");

                    if (string.IsNullOrEmpty(name)) continue;

                    // Read the value
                    var value = reader.ReadElementContentAsString();

                    if (name.Equals("_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase) ||
                        name.Equals("Print_Area", StringComparison.OrdinalIgnoreCase))
                    {
                        // Parse print area
                        var ranges = ParseDefinedNameValue(value, sheetName, log);
                        if (ranges.Count > 0)
                        {
                            pageSetup.PrintArea.AddRange(ranges);
                            log?.Invoke($"[PageSetupReader] Found print area: {value}");
                        }
                    }
                    else if (name.Equals("_xlnm.Print_Titles", StringComparison.OrdinalIgnoreCase) ||
                             name.Equals("Print_Titles", StringComparison.OrdinalIgnoreCase))
                    {
                        // Parse print titles
                        ParsePrintTitles(value, sheetName, pageSetup, log);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[PageSetupReader] Error reading print ranges: {ex.Message}");
        }
    }

    private static void ReadPageSetupElement(XmlReader reader, PageSetupData pageSetup, Action<string>? log)
    {
        // Paper size
        if (reader.GetAttribute("paperSize") is { } paperSizeStr &&
            ushort.TryParse(paperSizeStr, out var paperSize))
        {
            pageSetup.PaperSize = (PaperSize)paperSize;
            log?.Invoke($"[PageSetupReader] Paper size: {pageSetup.PaperSize}");
        }

        // Orientation
        if (reader.GetAttribute("orientation") is { } orientationStr)
        {
            pageSetup.Orientation = orientationStr.Equals("landscape", StringComparison.OrdinalIgnoreCase)
                ? PageOrientation.Landscape
                : PageOrientation.Portrait;
            log?.Invoke($"[PageSetupReader] Orientation: {pageSetup.Orientation}");
        }

        // Scale
        if (reader.GetAttribute("scale") is { } scaleStr &&
            int.TryParse(scaleStr, out var scale))
        {
            pageSetup.Scale = scale;
            log?.Invoke($"[PageSetupReader] Scale: {scale}%");
        }

        // Fit to width
        if (reader.GetAttribute("fitToWidth") is { } fitToWidthStr &&
            int.TryParse(fitToWidthStr, out var fitToWidth))
        {
            pageSetup.FitToWidth = fitToWidth;
            log?.Invoke($"[PageSetupReader] Fit to width: {fitToWidth}");
        }

        // Fit to height
        if (reader.GetAttribute("fitToHeight") is { } fitToHeightStr &&
            int.TryParse(fitToHeightStr, out var fitToHeight))
        {
            pageSetup.FitToHeight = fitToHeight;
            log?.Invoke($"[PageSetupReader] Fit to height: {fitToHeight}");
        }

        // First page number
        if (reader.GetAttribute("firstPageNumber") is { } firstPageStr &&
            int.TryParse(firstPageStr, out var firstPage))
        {
            pageSetup.FirstPageNumber = firstPage;
            log?.Invoke($"[PageSetupReader] First page number: {firstPage}");
        }

        // Page order
        if (reader.GetAttribute("pageOrder") is { } pageOrderStr)
        {
            pageSetup.PageOrder = pageOrderStr.Equals("overThenDown", StringComparison.OrdinalIgnoreCase)
                ? PageOrder.OverThenDown
                : PageOrder.DownThenOver;
            log?.Invoke($"[PageSetupReader] Page order: {pageSetup.PageOrder}");
        }

        // Black and white
        if (reader.GetAttribute("blackAndWhite") is { } bwStr &&
            bool.TryParse(bwStr, out var blackAndWhite))
        {
            pageSetup.BlackAndWhite = blackAndWhite;
            log?.Invoke($"[PageSetupReader] Black and white: {blackAndWhite}");
        }

        // Draft quality
        if (reader.GetAttribute("draft") is { } draftStr &&
            bool.TryParse(draftStr, out var draft))
        {
            pageSetup.DraftQuality = draft;
            log?.Invoke($"[PageSetupReader] Draft quality: {draft}");
        }

        // Print comments
        if (reader.GetAttribute("cellComments") is { } commentsStr)
        {
            pageSetup.PrintComments = commentsStr.ToLowerInvariant() switch
            {
                "asDisplayed" => PrintComments.AsDisplayed,
                "atEnd" => PrintComments.AtEnd,
                _ => PrintComments.None
            };
            log?.Invoke($"[PageSetupReader] Print comments: {pageSetup.PrintComments}");
        }

        // Cell errors
        if (reader.GetAttribute("errors") is { } errorsStr)
        {
            pageSetup.CellErrors = errorsStr.ToLowerInvariant() switch
            {
                "blank" => CellErrorPrint.Blank,
                "dash" => CellErrorPrint.DashDash,
                "NA" => CellErrorPrint.NA,
                _ => CellErrorPrint.Displayed
            };
            log?.Invoke($"[PageSetupReader] Cell errors: {pageSetup.CellErrors}");
        }

        // Use first page number
        if (reader.GetAttribute("useFirstPageNumber") is { } useFirstStr &&
            bool.TryParse(useFirstStr, out var useFirst))
        {
            if (!useFirst)
            {
                pageSetup.StartPageNumber = null;
            }
            log?.Invoke($"[PageSetupReader] Use first page number: {useFirst}");
        }

        // Horizontal center
        if (reader.GetAttribute("horizontalCentered") is { } hCenterStr &&
            bool.TryParse(hCenterStr, out var hCenter))
        {
            pageSetup.CenterHorizontally = hCenter;
            log?.Invoke($"[PageSetupReader] Center horizontally: {hCenter}");
        }

        // Vertical center
        if (reader.GetAttribute("verticalCentered") is { } vCenterStr &&
            bool.TryParse(vCenterStr, out var vCenter))
        {
            pageSetup.CenterVertically = vCenter;
            log?.Invoke($"[PageSetupReader] Center vertically: {vCenter}");
        }
    }

    private static void ReadPrintOptions(XmlReader reader, PageSetupData pageSetup, Action<string>? log)
    {
        // Print gridlines
        if (reader.GetAttribute("gridLines") is { } gridLinesStr &&
            bool.TryParse(gridLinesStr, out var gridLines))
        {
            pageSetup.PrintGridlines = gridLines;
            log?.Invoke($"[PageSetupReader] Print gridlines: {gridLines}");
        }

        // Print headings
        if (reader.GetAttribute("headings") is { } headingsStr &&
            bool.TryParse(headingsStr, out var headings))
        {
            pageSetup.PrintHeadings = headings;
            log?.Invoke($"[PageSetupReader] Print headings: {headings}");
        }

        // Horizontal center
        if (reader.GetAttribute("horizontalCentered") is { } hCenterStr &&
            bool.TryParse(hCenterStr, out var hCenter))
        {
            pageSetup.CenterHorizontally = hCenter;
            log?.Invoke($"[PageSetupReader] Center horizontally: {hCenter}");
        }

        // Vertical center
        if (reader.GetAttribute("verticalCentered") is { } vCenterStr &&
            bool.TryParse(vCenterStr, out var vCenter))
        {
            pageSetup.CenterVertically = vCenter;
            log?.Invoke($"[PageSetupReader] Center vertically: {vCenter}");
        }
    }

    private static void ReadPageMargins(XmlReader reader, PageSetupData pageSetup, Action<string>? log)
    {
        if (reader.GetAttribute("left") is { } leftStr &&
            double.TryParse(leftStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var left))
        {
            pageSetup.Margins.Left = left;
        }

        if (reader.GetAttribute("right") is { } rightStr &&
            double.TryParse(rightStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var right))
        {
            pageSetup.Margins.Right = right;
        }

        if (reader.GetAttribute("top") is { } topStr &&
            double.TryParse(topStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var top))
        {
            pageSetup.Margins.Top = top;
        }

        if (reader.GetAttribute("bottom") is { } bottomStr &&
            double.TryParse(bottomStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var bottom))
        {
            pageSetup.Margins.Bottom = bottom;
        }

        if (reader.GetAttribute("header") is { } headerStr &&
            double.TryParse(headerStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var header))
        {
            pageSetup.Margins.Header = header;
            pageSetup.HeaderMargin = header;
        }

        if (reader.GetAttribute("footer") is { } footerStr &&
            double.TryParse(footerStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var footer))
        {
            pageSetup.Margins.Footer = footer;
            pageSetup.FooterMargin = footer;
        }

        log?.Invoke($"[PageSetupReader] Margins: L={pageSetup.Margins.Left}, R={pageSetup.Margins.Right}, T={pageSetup.Margins.Top}, B={pageSetup.Margins.Bottom}");
    }

    private static List<CellRange> ParseDefinedNameValue(string value, string sheetName, Action<string>? log)
    {
        var ranges = new List<CellRange>();

        try
        {
            // Remove sheet name prefix if present (e.g., "Sheet1!$A$1:$B$10")
            var rangePart = value;
            if (value.Contains('!'))
            {
                var parts = value.Split('!', 2);
                if (parts.Length == 2)
                {
                    rangePart = parts[1];
                }
            }

            // Handle multiple ranges separated by comma
            var rangeStrings = rangePart.Split(',');

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
            log?.Invoke($"[PageSetupReader] Error parsing range '{value}': {ex.Message}");
        }

        return ranges;
    }

    private static CellRange? ParseRange(string rangeStr)
    {
        // Parse range like "$A$1:$B$10" or "$A:$B" or "$1:$10"
        rangeStr = rangeStr.Replace("$", "");

        if (rangeStr.Contains(':'))
        {
            var parts = rangeStr.Split(':', 2);
            if (parts.Length == 2)
            {
                var (col1, row1) = ParseCellReference(parts[0]);
                var (col2, row2) = ParseCellReference(parts[1]);

                return new CellRange
                {
                    FirstRow = row1 ?? 0,
                    FirstCol = col1 ?? 0,
                    LastRow = row2 ?? 65535,
                    LastCol = col2 ?? 255
                };
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

    private static void ParsePrintTitles(string value, string sheetName, PageSetupData pageSetup, Action<string>? log)
    {
        try
        {
            // Remove sheet name prefix if present
            var rangePart = value;
            if (value.Contains('!'))
            {
                var parts = value.Split('!', 2);
                if (parts.Length == 2)
                {
                    rangePart = parts[1];
                }
            }

            // Print titles can be rows (e.g., "$1:$2") or columns (e.g., "$A:$B")
            // or both separated by comma
            var ranges = rangePart.Split(',');

            foreach (var range in ranges)
            {
                var trimmed = range.Trim().Replace("$", "");

                if (trimmed.Contains(':'))
                {
                    var parts = trimmed.Split(':', 2);

                    // Check if it's rows (both parts are numbers) or columns (both parts are letters)
                    var isRow1 = int.TryParse(parts[0], out _);
                    var isRow2 = int.TryParse(parts[1], out _);

                    if (isRow1 && isRow2)
                    {
                        // Rows
                        var row1 = int.Parse(parts[0]) - 1;
                        var row2 = int.Parse(parts[1]) - 1;
                        pageSetup.PrintTitleRows = new CellRange
                        {
                            FirstRow = row1,
                            LastRow = row2,
                            FirstCol = 0,
                            LastCol = 255
                        };
                        log?.Invoke($"[PageSetupReader] Print title rows: {row1 + 1} to {row2 + 1}");
                    }
                    else if (!isRow1 && !isRow2)
                    {
                        // Columns
                        var col1 = ParseColumn(parts[0]);
                        var col2 = ParseColumn(parts[1]);
                        pageSetup.PrintTitleColumns = new CellRange
                        {
                            FirstRow = 0,
                            LastRow = 65535,
                            FirstCol = col1,
                            LastCol = col2
                        };
                        log?.Invoke($"[PageSetupReader] Print title columns: {col1} to {col2}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[PageSetupReader] Error parsing print titles '{value}': {ex.Message}");
        }
    }

    private static int ParseColumn(string colStr)
    {
        var col = 0;
        foreach (var c in colStr.ToUpperInvariant())
        {
            col = col * 26 + (c - 'A' + 1);
        }
        return col - 1; // Convert to 0-based
    }
}
