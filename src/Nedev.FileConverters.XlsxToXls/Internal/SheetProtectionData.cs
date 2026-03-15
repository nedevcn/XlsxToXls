namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Worksheet protection settings.
/// </summary>
public sealed class SheetProtection
{
    /// <summary>
    /// Gets or sets whether the sheet is protected.
    /// </summary>
    public bool IsProtected { get; set; }

    /// <summary>
    /// Gets or sets the password hash (BIFF8 format).
    /// </summary>
    public ushort PasswordHash { get; set; }

    /// <summary>
    /// Gets or sets the protection options flags.
    /// </summary>
    public ProtectionOptions Options { get; set; } = new();

    /// <summary>
    /// Gets or sets whether to allow users to edit objects.
    /// </summary>
    public bool AllowEditObjects { get; set; }

    /// <summary>
    /// Gets or sets whether to allow users to edit scenarios.
    /// </summary>
    public bool AllowEditScenarios { get; set; }
}

/// <summary>
/// Worksheet protection options.
/// </summary>
public sealed class ProtectionOptions
{
    /// <summary>
    /// Gets or sets whether cells are locked by default.
    /// </summary>
    public bool LockCells { get; set; } = true;

    /// <summary>
    /// Gets or sets whether cells are hidden by default.
    /// </summary>
    public bool HideCells { get; set; }

    /// <summary>
    /// Gets or sets whether to allow formatting cells.
    /// </summary>
    public bool AllowFormatCells { get; set; }

    /// <summary>
    /// Gets or sets whether to allow formatting columns.
    /// </summary>
    public bool AllowFormatColumns { get; set; }

    /// <summary>
    /// Gets or sets whether to allow formatting rows.
    /// </summary>
    public bool AllowFormatRows { get; set; }

    /// <summary>
    /// Gets or sets whether to allow inserting columns.
    /// </summary>
    public bool AllowInsertColumns { get; set; }

    /// <summary>
    /// Gets or sets whether to allow inserting rows.
    /// </summary>
    public bool AllowInsertRows { get; set; }

    /// <summary>
    /// Gets or sets whether to allow inserting hyperlinks.
    /// </summary>
    public bool AllowInsertHyperlinks { get; set; }

    /// <summary>
    /// Gets or sets whether to allow deleting columns.
    /// </summary>
    public bool AllowDeleteColumns { get; set; }

    /// <summary>
    /// Gets or sets whether to allow deleting rows.
    /// </summary>
    public bool AllowDeleteRows { get; set; }

    /// <summary>
    /// Gets or sets whether to allow selecting locked cells.
    /// </summary>
    public bool AllowSelectLockedCells { get; set; } = true;

    /// <summary>
    /// Gets or sets whether to allow sorting.
    /// </summary>
    public bool AllowSort { get; set; }

    /// <summary>
    /// Gets or sets whether to allow using AutoFilter.
    /// </summary>
    public bool AllowAutoFilter { get; set; }

    /// <summary>
    /// Gets or sets whether to allow using PivotTables.
    /// </summary>
    public bool AllowPivotTables { get; set; }

    /// <summary>
    /// Gets or sets whether to allow selecting unlocked cells.
    /// </summary>
    public bool AllowSelectUnlockedCells { get; set; } = true;
}

/// <summary>
/// Workbook protection settings.
/// </summary>
public sealed class WorkbookProtection
{
    /// <summary>
    /// Gets or sets whether the workbook structure is protected.
    /// </summary>
    public bool ProtectStructure { get; set; }

    /// <summary>
    /// Gets or sets whether the workbook windows are protected.
    /// </summary>
    public bool ProtectWindows { get; set; }

    /// <summary>
    /// Gets or sets the password hash for workbook protection.
    /// </summary>
    public ushort PasswordHash { get; set; }
}

/// <summary>
/// Utility class for password hashing in Excel BIFF8 format.
/// </summary>
internal static class PasswordHasher
{
    /// <summary>
    /// Computes the BIFF8 password hash for a given password.
    /// </summary>
    /// <param name="password">The password to hash.</param>
    /// <returns>The 16-bit password hash.</returns>
    public static ushort ComputePasswordHash(string password)
    {
        if (string.IsNullOrEmpty(password))
            return 0;

        var passwordBytes = System.Text.Encoding.Unicode.GetBytes(password);
        ushort hash = 0;

        for (var i = password.Length - 1; i >= 0; i--)
        {
            hash = (ushort)(((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF));
            hash ^= (ushort)password[i];
        }

        hash = (ushort)(((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF));
        hash ^= (ushort)password.Length;
        hash ^= 0xCE4B;

        return hash;
    }
}
