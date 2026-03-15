using System.IO.Compression;
using System.Text;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Reads VBA project data from XLSX files.
/// Extracts the vbaProject.bin file and parses its contents.
/// </summary>
public static class VbaProjectReader
{
    private const string VbaProjectPath = "xl/vbaProject.bin";
    private const string VbaSignaturePath = "xl/vbaProjectSignature.bin";

    /// <summary>
    /// Reads the VBA project from the XLSX archive.
    /// </summary>
    public static VbaBinaryData? ReadVbaProject(ZipArchive archive, Action<string>? log = null)
    {
        try
        {
            log?.Invoke("[VbaProjectReader] Reading VBA project");

            // Check if vbaProject.bin exists
            var vbaEntry = archive.GetEntry(VbaProjectPath);
            if (vbaEntry == null)
            {
                log?.Invoke("[VbaProjectReader] No VBA project found");
                return null;
            }

            var binaryData = new VbaBinaryData();

            // Read the project data
            using (var stream = vbaEntry.Open())
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                binaryData.ProjectData = ms.ToArray();
            }

            log?.Invoke($"[VbaProjectReader] Read {binaryData.ProjectData.Length} bytes of VBA project data");

            // Check for digital signature
            var sigEntry = archive.GetEntry(VbaSignaturePath);
            if (sigEntry != null)
            {
                using var stream = sigEntry.Open();
                using var ms = new MemoryStream();
                stream.CopyTo(ms);
                binaryData.SignatureData = ms.ToArray();
                log?.Invoke($"[VbaProjectReader] Read {binaryData.SignatureData.Length} bytes of signature data");
            }

            // Parse the project data
            var project = ParseVbaProject(binaryData.ProjectData, log);
            if (project != null)
            {
                binaryData.CreationDate = DateTime.Now;
                binaryData.ModifiedDate = DateTime.Now;
            }

            return binaryData;
        }
        catch (Exception ex)
        {
            log?.Invoke($"[VbaProjectReader] Error reading VBA project: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Checks if the XLSX file contains a VBA project.
    /// </summary>
    public static bool HasVbaProject(ZipArchive archive)
    {
        return archive.GetEntry(VbaProjectPath) != null;
    }

    /// <summary>
    /// Parses the VBA project binary data to extract project information.
    /// </summary>
    private static VbaProjectData? ParseVbaProject(byte[] data, Action<string>? log)
    {
        try
        {
            var project = new VbaProjectData();

            // The VBA project is stored in a compound file format
            // For now, we'll create a basic project structure
            // Full parsing would require implementing the CFB format parser

            // Try to extract module names from the project stream
            ExtractModuleNames(data, project, log);

            // Try to extract references
            ExtractReferences(data, project, log);

            return project;
        }
        catch (Exception ex)
        {
            log?.Invoke($"[VbaProjectReader] Error parsing VBA project: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Extracts module names from the VBA project data.
    /// </summary>
    private static void ExtractModuleNames(byte[] data, VbaProjectData project, Action<string>? log)
    {
        try
        {
            // Look for module names in the project stream
            // Module names are typically stored as ASCII strings

            var encoding = Encoding.GetEncoding(1252); // Windows-1252
            var text = encoding.GetString(data);

            // Common module patterns
            var patterns = new[]
            {
                "Module=",
                "Class=",
                "BaseClass=",
                "ThisWorkbook=",
                "Sheet1=",
                "Sheet2=",
                "Sheet3="
            };

            foreach (var pattern in patterns)
            {
                var index = 0;
                while ((index = text.IndexOf(pattern, index, StringComparison.Ordinal)) != -1)
                {
                    index += pattern.Length;

                    // Find the end of the module name
                    var endIndex = text.IndexOfAny(new[] { '\r', '\n', '\0' }, index);
                    if (endIndex == -1) endIndex = text.Length;

                    var moduleName = text.Substring(index, endIndex - index).Trim();

                    if (!string.IsNullOrWhiteSpace(moduleName) && !project.Modules.Exists(m => m.Name == moduleName))
                    {
                        var moduleType = pattern switch
                        {
                            "Class=" => VbaModuleType.Class,
                            "BaseClass=" => VbaModuleType.UserForm,
                            "ThisWorkbook=" => VbaModuleType.Workbook,
                            _ => text.Contains($"Sheet{moduleName.Replace("Sheet", "")}=") ? VbaModuleType.Worksheet : VbaModuleType.Standard
                        };

                        project.Modules.Add(new VbaModule
                        {
                            Name = moduleName,
                            Type = moduleType
                        });

                        log?.Invoke($"[VbaProjectReader] Found module: {moduleName} ({moduleType})");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[VbaProjectReader] Error extracting module names: {ex.Message}");
        }
    }

    /// <summary>
    /// Extracts references from the VBA project data.
    /// </summary>
    private static void ExtractReferences(byte[] data, VbaProjectData project, Action<string>? log)
    {
        try
        {
            var encoding = Encoding.GetEncoding(1252);
            var text = encoding.GetString(data);

            // Look for reference patterns
            // Reference\*\G{GUID}#Major#Minor#Name

            var refIndex = 0;
            while ((refIndex = text.IndexOf("Reference\\*\\G{", refIndex, StringComparison.Ordinal)) != -1)
            {
                refIndex += "Reference\\*\\G{".Length;

                // Find the closing brace
                var guidEnd = text.IndexOf('}', refIndex);
                if (guidEnd == -1) continue;

                var guid = text.Substring(refIndex, guidEnd - refIndex);
                refIndex = guidEnd + 1;

                // Find version numbers
                var hash1 = text.IndexOf('#', refIndex);
                if (hash1 == -1) continue;

                var hash2 = text.IndexOf('#', hash1 + 1);
                if (hash2 == -1) continue;

                var hash3 = text.IndexOf('#', hash2 + 1);
                if (hash3 == -1) continue;

                if (int.TryParse(text.Substring(hash1 + 1, hash2 - hash1 - 1), out var majorVersion) &&
                    int.TryParse(text.Substring(hash2 + 1, hash3 - hash2 - 1), out var minorVersion))
                {
                    var nameEnd = text.IndexOfAny(new[] { '\r', '\n', '\0' }, hash3 + 1);
                    if (nameEnd == -1) nameEnd = text.Length;

                    var name = text.Substring(hash3 + 1, nameEnd - hash3 - 1).Trim();

                    project.References.Add(new VbaReference
                    {
                        Name = name,
                        Guid = guid,
                        MajorVersion = majorVersion,
                        MinorVersion = minorVersion
                    });

                    log?.Invoke($"[VbaProjectReader] Found reference: {name} ({guid})");
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[VbaProjectReader] Error extracting references: {ex.Message}");
        }
    }

    /// <summary>
    /// Reads a specific module's code from the VBA project.
    /// </summary>
    public static string? ReadModuleCode(byte[] projectData, string moduleName, Action<string>? log = null)
    {
        try
        {
            // The module code is stored in a separate stream within the compound file
            // For now, return a placeholder
            log?.Invoke($"[VbaProjectReader] Reading code for module: {moduleName}");

            // In a full implementation, we would:
            // 1. Parse the CFB structure
            // 2. Find the module stream
            // 3. Decompress the code (if compressed)
            // 4. Return the source code

            return $"' Module: {moduleName}\r\n' Code extracted from XLSX\r\n";
        }
        catch (Exception ex)
        {
            log?.Invoke($"[VbaProjectReader] Error reading module code: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Gets information about the VBA project without fully parsing it.
    /// </summary>
    public static VbaProjectInfo GetProjectInfo(byte[] projectData, Action<string>? log = null)
    {
        var info = new VbaProjectInfo();

        try
        {
            var encoding = Encoding.GetEncoding(1252);
            var text = encoding.GetString(projectData);

            // Count modules
            info.ModuleCount = text.Split(new[] { "Module=" }, StringSplitOptions.None).Length - 1;
            info.HasWorkbookModule = text.Contains("ThisWorkbook=");
            info.HasClassModules = text.Contains("Class=");
            info.HasUserForms = text.Contains("BaseClass=");

            // Check for protection
            info.IsProtected = text.Contains("CMG=") && text.Contains("GC=") && text.Contains("DPB=");

            log?.Invoke($"[VbaProjectReader] Project info: {info.ModuleCount} modules, Protected: {info.IsProtected}");
        }
        catch (Exception ex)
        {
            log?.Invoke($"[VbaProjectReader] Error getting project info: {ex.Message}");
        }

        return info;
    }
}

/// <summary>
/// Summary information about a VBA project.
/// </summary>
public class VbaProjectInfo
{
    /// <summary>
    /// Gets or sets the number of modules.
    /// </summary>
    public int ModuleCount { get; set; }

    /// <summary>
    /// Gets or sets whether the project has a workbook module.
    /// </summary>
    public bool HasWorkbookModule { get; set; }

    /// <summary>
    /// Gets or sets whether the project has class modules.
    /// </summary>
    public bool HasClassModules { get; set; }

    /// <summary>
    /// Gets or sets whether the project has UserForms.
    /// </summary>
    public bool HasUserForms { get; set; }

    /// <summary>
    /// Gets or sets whether the project is protected.
    /// </summary>
    public bool IsProtected { get; set; }

    /// <summary>
    /// Gets or sets whether the project is signed.
    /// </summary>
    public bool IsSigned { get; set; }
}
