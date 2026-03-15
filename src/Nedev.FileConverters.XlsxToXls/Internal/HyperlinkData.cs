namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Represents a hyperlink in a cell for BIFF8 format conversion.
/// Supports URL links, file links, email links, and document links.
/// </summary>
public sealed class HyperlinkData
{
    /// <summary>
    /// Gets or sets the row index of the cell containing the hyperlink.
    /// </summary>
    public int Row { get; set; }

    /// <summary>
    /// Gets or sets the column index of the cell containing the hyperlink.
    /// </summary>
    public int Column { get; set; }

    /// <summary>
    /// Gets or sets the hyperlink type.
    /// </summary>
    public HyperlinkType Type { get; set; } = HyperlinkType.Url;

    /// <summary>
    /// Gets or sets the target URL or file path.
    /// </summary>
    public string Target { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the display text shown in the cell.
    /// If null, the target URL is displayed.
    /// </summary>
    public string? DisplayText { get; set; }

    /// <summary>
    /// Gets or sets the tooltip text shown on hover.
    /// </summary>
    public string? Tooltip { get; set; }

    /// <summary>
    /// Gets or sets the screen tip text (alternative to tooltip).
    /// </summary>
    public string? ScreenTip { get; set; }

    /// <summary>
    /// Gets or sets the location within a document (for document links).
    /// </summary>
    public string? Location { get; set; }

    /// <summary>
    /// Gets or sets the email subject (for mailto links).
    /// </summary>
    public string? EmailSubject { get; set; }

    /// <summary>
    /// Gets or sets whether this is a one-click hyperlink.
    /// </summary>
    public bool IsOneClick { get; set; }

    /// <summary>
    /// Gets or sets the hyperlink reference index for shared hyperlinks.
    /// </summary>
    public int RefIndex { get; set; } = -1;

    /// <summary>
    /// Creates a URL hyperlink.
    /// </summary>
    public static HyperlinkData CreateUrl(int row, int col, string url, string? displayText = null, string? tooltip = null)
    {
        return new HyperlinkData
        {
            Row = row,
            Column = col,
            Type = HyperlinkType.Url,
            Target = url,
            DisplayText = displayText,
            Tooltip = tooltip
        };
    }

    /// <summary>
    /// Creates a file hyperlink.
    /// </summary>
    public static HyperlinkData CreateFile(int row, int col, string filePath, string? displayText = null, string? tooltip = null)
    {
        return new HyperlinkData
        {
            Row = row,
            Column = col,
            Type = HyperlinkType.File,
            Target = filePath,
            DisplayText = displayText,
            Tooltip = tooltip
        };
    }

    /// <summary>
    /// Creates an email hyperlink.
    /// </summary>
    public static HyperlinkData CreateEmail(int row, int col, string email, string? subject = null, string? displayText = null)
    {
        var target = $"mailto:{email}";
        if (!string.IsNullOrEmpty(subject))
        {
            target += $"?subject={Uri.EscapeDataString(subject)}";
        }

        return new HyperlinkData
        {
            Row = row,
            Column = col,
            Type = HyperlinkType.Email,
            Target = target,
            EmailSubject = subject,
            DisplayText = displayText ?? email
        };
    }

    /// <summary>
    /// Creates a document hyperlink (link to a specific location in a document).
    /// </summary>
    public static HyperlinkData CreateDocument(int row, int col, string documentPath, string location, string? displayText = null)
    {
        return new HyperlinkData
        {
            Row = row,
            Column = col,
            Type = HyperlinkType.Document,
            Target = documentPath,
            Location = location,
            DisplayText = displayText
        };
    }

    /// <summary>
    /// Creates an internal hyperlink (link to a cell or named range in the same workbook).
    /// </summary>
    public static HyperlinkData CreateInternal(int row, int col, string reference, string? displayText = null)
    {
        return new HyperlinkData
        {
            Row = row,
            Column = col,
            Type = HyperlinkType.Internal,
            Target = reference,
            DisplayText = displayText
        };
    }

    /// <summary>
    /// Gets the cell reference string (e.g., "A1").
    /// </summary>
    public string GetCellReference()
    {
        return CellReferenceConverter.ToReference(Row, Column);
    }

    /// <summary>
    /// Validates the hyperlink data.
    /// </summary>
    public bool IsValid()
    {
        if (Row < 0 || Column < 0)
            return false;

        if (string.IsNullOrWhiteSpace(Target))
            return false;

        return Type switch
        {
            HyperlinkType.Url => Uri.TryCreate(Target, UriKind.Absolute, out var uri) &&
                                (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps),
            HyperlinkType.File => !string.IsNullOrWhiteSpace(Target),
            HyperlinkType.Email => Target.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase),
            HyperlinkType.Document => !string.IsNullOrWhiteSpace(Target) && !string.IsNullOrWhiteSpace(Location),
            HyperlinkType.Internal => !string.IsNullOrWhiteSpace(Target),
            _ => false
        };
    }
}

/// <summary>
/// Types of hyperlinks supported in Excel.
/// </summary>
public enum HyperlinkType : byte
{
    /// <summary>URL hyperlink (http:// or https://)</summary>
    Url = 0,

    /// <summary>File or folder link</summary>
    File = 1,

    /// <summary>Email link (mailto:)</summary>
    Email = 2,

    /// <summary>Link to a specific location in a document</summary>
    Document = 3,

    /// <summary>Internal link within the same workbook</summary>
    Internal = 4,

    /// <summary>UNC path (network share)</summary>
    Unc = 5,

    /// <summary>Worksheet reference</summary>
    Worksheet = 6,

    /// <summary>Named range reference</summary>
    NamedRange = 7
}

/// <summary>
/// Collection of hyperlinks for a worksheet.
/// </summary>
public sealed class HyperlinkCollection
{
    private readonly List<HyperlinkData> _hyperlinks = new();

    /// <summary>
    /// Gets all hyperlinks in the collection.
    /// </summary>
    public IReadOnlyList<HyperlinkData> Hyperlinks => _hyperlinks;

    /// <summary>
    /// Gets the number of hyperlinks.
    /// </summary>
    public int Count => _hyperlinks.Count;

    /// <summary>
    /// Adds a hyperlink to the collection.
    /// </summary>
    public void Add(HyperlinkData hyperlink)
    {
        if (hyperlink == null)
            throw new ArgumentNullException(nameof(hyperlink));

        // Remove any existing hyperlink at the same cell
        RemoveAt(hyperlink.Row, hyperlink.Column);

        _hyperlinks.Add(hyperlink);
    }

    /// <summary>
    /// Removes the hyperlink at the specified cell.
    /// </summary>
    public bool RemoveAt(int row, int col)
    {
        var existing = _hyperlinks.FindIndex(h => h.Row == row && h.Column == col);
        if (existing >= 0)
        {
            _hyperlinks.RemoveAt(existing);
            return true;
        }
        return false;
    }

    /// <summary>
    /// Gets the hyperlink at the specified cell, if any.
    /// </summary>
    public HyperlinkData? GetAt(int row, int col)
    {
        return _hyperlinks.Find(h => h.Row == row && h.Column == col);
    }

    /// <summary>
    /// Checks if a cell has a hyperlink.
    /// </summary>
    public bool HasHyperlink(int row, int col)
    {
        return _hyperlinks.Exists(h => h.Row == row && h.Column == col);
    }

    /// <summary>
    /// Clears all hyperlinks.
    /// </summary>
    public void Clear()
    {
        _hyperlinks.Clear();
    }

    /// <summary>
    /// Gets all hyperlinks of a specific type.
    /// </summary>
    public IEnumerable<HyperlinkData> GetByType(HyperlinkType type)
    {
        return _hyperlinks.Where(h => h.Type == type);
    }

    /// <summary>
    /// Gets all hyperlinks in a specific row.
    /// </summary>
    public IEnumerable<HyperlinkData> GetByRow(int row)
    {
        return _hyperlinks.Where(h => h.Row == row);
    }

    /// <summary>
    /// Gets all hyperlinks in a specific column.
    /// </summary>
    public IEnumerable<HyperlinkData> GetByColumn(int col)
    {
        return _hyperlinks.Where(h => h.Column == col);
    }
}

/// <summary>
/// Helper class for converting between cell references and row/column indices.
/// </summary>
public static class CellReferenceConverter
{
    /// <summary>
    /// Converts row and column indices to a cell reference (e.g., "A1", "AB100").
    /// </summary>
    public static string ToReference(int row, int col)
    {
        if (row < 0 || col < 0)
            throw new ArgumentOutOfRangeException("Row and column must be non-negative");

        var colLetters = ColumnToLetters(col);
        return $"{colLetters}{row + 1}";
    }

    /// <summary>
    /// Converts a cell reference to row and column indices.
    /// </summary>
    public static (int Row, int Col) FromReference(string reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
            throw new ArgumentException("Reference cannot be null or empty", nameof(reference));

        reference = reference.ToUpperInvariant();

        var colPart = string.Empty;
        var rowPart = string.Empty;

        foreach (var c in reference)
        {
            if (char.IsLetter(c))
            {
                if (rowPart.Length > 0)
                    throw new ArgumentException("Invalid cell reference format", nameof(reference));
                colPart += c;
            }
            else if (char.IsDigit(c))
            {
                rowPart += c;
            }
            else
            {
                throw new ArgumentException("Invalid character in cell reference", nameof(reference));
            }
        }

        if (colPart.Length == 0 || rowPart.Length == 0)
            throw new ArgumentException("Invalid cell reference format", nameof(reference));

        var col = LettersToColumn(colPart);
        var row = int.Parse(rowPart) - 1;

        return (row, col);
    }

    /// <summary>
    /// Converts a column index to letters (0 = A, 25 = Z, 26 = AA, etc.).
    /// </summary>
    public static string ColumnToLetters(int col)
    {
        if (col < 0)
            throw new ArgumentOutOfRangeException(nameof(col), "Column must be non-negative");

        var result = string.Empty;
        col++;

        while (col > 0)
        {
            col--;
            result = (char)('A' + (col % 26)) + result;
            col /= 26;
        }

        return result;
    }

    /// <summary>
    /// Converts column letters to an index (A = 0, Z = 25, AA = 26, etc.).
    /// </summary>
    public static int LettersToColumn(string letters)
    {
        if (string.IsNullOrWhiteSpace(letters))
            throw new ArgumentException("Letters cannot be null or empty", nameof(letters));

        letters = letters.ToUpperInvariant();
        var result = 0;

        foreach (var c in letters)
        {
            if (c < 'A' || c > 'Z')
                throw new ArgumentException("Invalid column letters", nameof(letters));

            result = result * 26 + (c - 'A' + 1);
        }

        return result - 1;
    }
}
