namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Represents a VBA project for BIFF8 format conversion.
/// Contains all VBA modules, references, and project properties.
/// </summary>
public sealed class VbaProjectData
{
    /// <summary>
    /// Gets or sets the project name.
    /// </summary>
    public string Name { get; set; } = "VBAProject";

    /// <summary>
    /// Gets or sets the project description.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Gets or sets the project help file.
    /// </summary>
    public string? HelpFile { get; set; }

    /// <summary>
    /// Gets or sets the project help context ID.
    /// </summary>
    public int HelpContextId { get; set; }

    /// <summary>
    /// Gets or sets the project protection state.
    /// </summary>
    public VbaProtection Protection { get; set; } = VbaProtection.None;

    /// <summary>
    /// Gets or sets whether the project is locked for viewing.
    /// </summary>
    public bool IsLocked { get; set; }

    /// <summary>
    /// Gets or sets the project password hash (if protected).
    /// </summary>
    public byte[]? PasswordHash { get; set; }

    /// <summary>
    /// Gets the collection of VBA modules.
    /// </summary>
    public List<VbaModule> Modules { get; set; } = new();

    /// <summary>
    /// Gets the collection of project references.
    /// </summary>
    public List<VbaReference> References { get; set; } = new();

    /// <summary>
    /// Gets or sets the raw project stream data.
    /// </summary>
    public byte[]? ProjectStream { get; set; }

    /// <summary>
    /// Gets or sets the raw project storage data.
    /// </summary>
    public byte[]? ProjectStorage { get; set; }

    /// <summary>
    /// Gets or sets the project version (major).
    /// </summary>
    public int VersionMajor { get; set; } = 7;

    /// <summary>
    /// Gets or sets the project version (minor).
    /// </summary>
    public int VersionMinor { get; set; } = 1;

    /// <summary>
    /// Gets or sets the code page for the project.
    /// </summary>
    public int CodePage { get; set; } = 1252;

    /// <summary>
    /// Gets or sets the LCID (locale ID) for the project.
    /// </summary>
    public int Lcid { get; set; } = 1033;

    /// <summary>
    /// Gets or sets the LCID for the module.
    /// </summary>
    public int LcidModule { get; set; } = 1033;

    /// <summary>
    /// Gets the number of modules in the project.
    /// </summary>
    public int ModuleCount => Modules.Count;

    /// <summary>
    /// Gets whether the project contains any modules.
    /// </summary>
    public bool HasModules => Modules.Count > 0;

    /// <summary>
    /// Adds a standard module to the project.
    /// </summary>
    public VbaModule AddStandardModule(string name, string code)
    {
        var module = new VbaModule
        {
            Name = name,
            Type = VbaModuleType.Standard,
            Code = code
        };
        Modules.Add(module);
        return module;
    }

    /// <summary>
    /// Adds a class module to the project.
    /// </summary>
    public VbaModule AddClassModule(string name, string code, bool isGlobal = false)
    {
        var module = new VbaModule
        {
            Name = name,
            Type = VbaModuleType.Class,
            Code = code,
            IsGlobal = isGlobal
        };
        Modules.Add(module);
        return module;
    }

    /// <summary>
    /// Adds a worksheet module to the project.
    /// </summary>
    public VbaModule AddWorksheetModule(string name, int sheetIndex, string code)
    {
        var module = new VbaModule
        {
            Name = name,
            Type = VbaModuleType.Worksheet,
            SheetIndex = sheetIndex,
            Code = code
        };
        Modules.Add(module);
        return module;
    }

    /// <summary>
    /// Adds a workbook module to the project.
    /// </summary>
    public VbaModule AddWorkbookModule(string code)
    {
        var module = new VbaModule
        {
            Name = "ThisWorkbook",
            Type = VbaModuleType.Workbook,
            Code = code
        };
        Modules.Add(module);
        return module;
    }

    /// <summary>
    /// Adds a reference to the project.
    /// </summary>
    public void AddReference(string name, string guid, int majorVersion, int minorVersion, string? path = null)
    {
        References.Add(new VbaReference
        {
            Name = name,
            Guid = guid,
            MajorVersion = majorVersion,
            MinorVersion = minorVersion,
            Path = path
        });
    }

    /// <summary>
    /// Gets a module by name.
    /// </summary>
    public VbaModule? GetModule(string name)
    {
        return Modules.Find(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Removes a module by name.
    /// </summary>
    public bool RemoveModule(string name)
    {
        var module = GetModule(name);
        if (module != null)
        {
            return Modules.Remove(module);
        }
        return false;
    }

    /// <summary>
    /// Validates the project data.
    /// </summary>
    public bool IsValid()
    {
        if (string.IsNullOrWhiteSpace(Name))
            return false;

        // Check for duplicate module names
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var module in Modules)
        {
            if (!names.Add(module.Name))
                return false; // Duplicate name

            if (!module.IsValid())
                return false;
        }

        return true;
    }
}

/// <summary>
/// Represents a single VBA module.
/// </summary>
public sealed class VbaModule
{
    /// <summary>
    /// Gets or sets the module name.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the module type.
    /// </summary>
    public VbaModuleType Type { get; set; } = VbaModuleType.Standard;

    /// <summary>
    /// Gets or sets the module code.
    /// </summary>
    public string Code { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the module stream name.
    /// </summary>
    public string? StreamName { get; set; }

    /// <summary>
    /// Gets or sets the module description.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Gets or sets whether this is a global module.
    /// </summary>
    public bool IsGlobal { get; set; }

    /// <summary>
    /// Gets or sets whether this is a private module.
    /// </summary>
    public bool IsPrivate { get; set; }

    /// <summary>
    /// Gets or sets the worksheet index (for worksheet modules).
    /// </summary>
    public int SheetIndex { get; set; } = -1;

    /// <summary>
    /// Gets or sets the module offset in the stream.
    /// </summary>
    public int Offset { get; set; }

    /// <summary>
    /// Gets or sets the module length.
    /// </summary>
    public int Length { get; set; }

    /// <summary>
    /// Gets the number of lines in the code.
    /// </summary>
    public int LineCount => Code.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Length;

    /// <summary>
    /// Gets whether the module has any code.
    /// </summary>
    public bool HasCode => !string.IsNullOrWhiteSpace(Code);

    /// <summary>
    /// Validates the module.
    /// </summary>
    public bool IsValid()
    {
        return !string.IsNullOrWhiteSpace(Name) && Name.Length <= 31;
    }
}

/// <summary>
/// Represents a VBA project reference.
/// </summary>
public sealed class VbaReference
{
    /// <summary>
    /// Gets or sets the reference name.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the reference GUID (for type libraries).
    /// </summary>
    public string? Guid { get; set; }

    /// <summary>
    /// Gets or sets the major version.
    /// </summary>
    public int MajorVersion { get; set; }

    /// <summary>
    /// Gets or sets the minor version.
    /// </summary>
    public int MinorVersion { get; set; }

    /// <summary>
    /// Gets or sets the reference path (for project references).
    /// </summary>
    public string? Path { get; set; }

    /// <summary>
    /// Gets or sets whether this is a control reference.
    /// </summary>
    public bool IsControl { get; set; }

    /// <summary>
    /// Gets or sets the control GUID.
    /// </summary>
    public string? ControlGuid { get; set; }

    /// <summary>
    /// Gets or sets the control cookie.
    /// </summary>
    public int ControlCookie { get; set; }

    /// <summary>
    /// Gets or sets the control type library GUID.
    /// </summary>
    public string? ControlTypeLibGuid { get; set; }

    /// <summary>
    /// Gets whether this is a type library reference.
    /// </summary>
    public bool IsTypeLibrary => !string.IsNullOrEmpty(Guid);

    /// <summary>
    /// Gets whether this is a project reference.
    /// </summary>
    public bool IsProjectReference => !string.IsNullOrEmpty(Path);
}

/// <summary>
/// Types of VBA modules.
/// </summary>
public enum VbaModuleType : byte
{
    /// <summary>Standard code module (.bas)</summary>
    Standard = 0,

    /// <summary>Class module (.cls)</summary>
    Class = 1,

    /// <summary>Worksheet module</summary>
    Worksheet = 2,

    /// <summary>Workbook module</summary>
    Workbook = 3,

    /// <summary>UserForm module (.frm)</summary>
    UserForm = 4,

    /// <summary>Designer module</summary>
    Designer = 5
}

/// <summary>
/// VBA project protection levels.
/// </summary>
public enum VbaProtection : byte
{
    /// <summary>No protection</summary>
    None = 0,

    /// <summary>Project is locked for viewing</summary>
    Locked = 1,

    /// <summary>Project is password protected</summary>
    PasswordProtected = 2
}

/// <summary>
/// Represents the raw VBA project binary data from XLSX.
/// </summary>
public sealed class VbaBinaryData
{
    /// <summary>
    /// Gets or sets the vbaProject.bin data.
    /// </summary>
    public byte[]? ProjectData { get; set; }

    /// <summary>
    /// Gets or sets the project signature (if signed).
    /// </summary>
    public byte[]? SignatureData { get; set; }

    /// <summary>
    /// Gets whether the project is signed.
    /// </summary>
    public bool IsSigned => SignatureData != null && SignatureData.Length > 0;

    /// <summary>
    /// Gets or sets the project creation date.
    /// </summary>
    public DateTime? CreationDate { get; set; }

    /// <summary>
    /// Gets or sets the project last modified date.
    /// </summary>
    public DateTime? ModifiedDate { get; set; }

    /// <summary>
    /// Gets whether the data is valid.
    /// </summary>
    public bool IsValid => ProjectData != null && ProjectData.Length > 0;
}
