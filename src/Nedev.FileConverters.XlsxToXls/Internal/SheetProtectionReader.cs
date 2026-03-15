using System.Xml;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Reads sheet protection settings from XLSX files.
/// </summary>
internal static class SheetProtectionReader
{
    /// <summary>
    /// Reads sheet protection from the worksheet XML.
    /// </summary>
    public static SheetProtectionInfo? ReadSheetProtection(XmlReader reader, string ns)
    {
        var isProtected = false;
        ushort passwordHash = 0;
        var allowEditObjects = true;
        var allowEditScenarios = true;
        var allowFormatCells = true;
        var allowFormatColumns = true;
        var allowFormatRows = true;
        var allowInsertColumns = true;
        var allowInsertRows = true;
        var allowInsertHyperlinks = true;
        var allowDeleteColumns = true;
        var allowDeleteRows = true;
        var allowSelectLockedCells = true;
        var allowSort = true;
        var allowAutoFilter = true;
        var allowPivotTables = true;
        var allowSelectUnlockedCells = true;

        // Check if sheet protection element exists
        if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "sheetProtection")
            return null;

        isProtected = true;

        // Read password hash if present
        var passwordAttr = reader.GetAttribute("password");
        if (!string.IsNullOrEmpty(passwordAttr) && ushort.TryParse(passwordAttr, out var pwdHash))
        {
            passwordHash = pwdHash;
        }

        // Read algorithm name (for newer Excel versions, we may need to handle this differently)
        var algorithm = reader.GetAttribute("algorithmName");
        var hashValue = reader.GetAttribute("hashValue");
        var saltValue = reader.GetAttribute("saltValue");
        var spinCount = reader.GetAttribute("spinCount");

        // If using modern encryption, we can't convert the password hash
        // but we still mark the sheet as protected
        if (!string.IsNullOrEmpty(algorithm))
        {
            // Modern encryption - set a default hash that indicates protection
            // The actual password won't work in XLS format
            passwordHash = 0x0000;
        }

        // Read protection options (attributes starting with "allow" or similar)
        // In XLSX, the default is true (allowed), and attributes are present when set to false
        allowEditObjects = reader.GetAttribute("objects") != "1";
        allowEditScenarios = reader.GetAttribute("scenarios") != "1";
        allowFormatCells = reader.GetAttribute("formatCells") != "0";
        allowFormatColumns = reader.GetAttribute("formatColumns") != "0";
        allowFormatRows = reader.GetAttribute("formatRows") != "0";
        allowInsertColumns = reader.GetAttribute("insertColumns") != "0";
        allowInsertRows = reader.GetAttribute("insertRows") != "0";
        allowInsertHyperlinks = reader.GetAttribute("insertHyperlinks") != "0";
        allowDeleteColumns = reader.GetAttribute("deleteColumns") != "0";
        allowDeleteRows = reader.GetAttribute("deleteRows") != "0";
        allowSelectLockedCells = reader.GetAttribute("selectLockedCells") != "0";
        allowSort = reader.GetAttribute("sort") != "0";
        allowAutoFilter = reader.GetAttribute("autoFilter") != "0";
        allowPivotTables = reader.GetAttribute("pivotTables") != "0";
        allowSelectUnlockedCells = reader.GetAttribute("selectUnlockedCells") != "0";

        return new SheetProtectionInfo(
            isProtected,
            passwordHash,
            allowEditObjects,
            allowEditScenarios,
            allowFormatCells,
            allowFormatColumns,
            allowFormatRows,
            allowInsertColumns,
            allowInsertRows,
            allowInsertHyperlinks,
            allowDeleteColumns,
            allowDeleteRows,
            allowSelectLockedCells,
            allowSort,
            allowAutoFilter,
            allowPivotTables,
            allowSelectUnlockedCells);
    }

    /// <summary>
    /// Reads workbook protection from the workbook XML.
    /// </summary>
    public static WorkbookProtectionInfo? ReadWorkbookProtection(XmlReader reader, string ns)
    {
        var protectStructure = false;
        var protectWindows = false;
        ushort passwordHash = 0;

        // Check if workbook protection element exists
        if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "workbookProtection")
            return null;

        // Read lock structure
        var lockStructureAttr = reader.GetAttribute("lockStructure");
        if (lockStructureAttr == "1" || lockStructureAttr?.ToLowerInvariant() == "true")
        {
            protectStructure = true;
        }

        // Read lock windows
        var lockWindowsAttr = reader.GetAttribute("lockWindows");
        if (lockWindowsAttr == "1" || lockWindowsAttr?.ToLowerInvariant() == "true")
        {
            protectWindows = true;
        }

        // Read password hash if present
        var passwordAttr = reader.GetAttribute("workbookPassword");
        if (!string.IsNullOrEmpty(passwordAttr) && ushort.TryParse(passwordAttr, out var pwdHash))
        {
            passwordHash = pwdHash;
        }

        // Check for modern encryption
        var algorithm = reader.GetAttribute("algorithmName");
        if (!string.IsNullOrEmpty(algorithm))
        {
            // Modern encryption - can't convert password hash
            passwordHash = 0x0000;
        }

        return new WorkbookProtectionInfo(protectStructure, protectWindows, passwordHash);
    }
}

/// <summary>
/// Workbook protection information for internal use.
/// </summary>
internal record struct WorkbookProtectionInfo(
    bool ProtectStructure,
    bool ProtectWindows,
    ushort PasswordHash);
