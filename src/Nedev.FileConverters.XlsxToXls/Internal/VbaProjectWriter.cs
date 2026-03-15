using System.Buffers;
using System.Buffers.Binary;
using System.Text;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Writes VBA project data to BIFF8 format.
/// Creates the _VBA_PROJECT_CUR stream and related storage.
/// </summary>
public sealed class VbaProjectWriter : IDisposable
{
    private const int DefaultBufferSize = 262144; // 256KB for VBA projects

    private byte[] _buffer;
    private Memory<byte> _memory;
    private int _position;
    private bool _isRented;

    /// <summary>
    /// Creates a new VbaProjectWriter with the specified buffer.
    /// </summary>
    public VbaProjectWriter(byte[] buffer)
    {
        _buffer = buffer;
        _memory = buffer;
        _position = 0;
        _isRented = false;
    }

    /// <summary>
    /// Creates a VbaProjectWriter using a pooled buffer.
    /// </summary>
    public static VbaProjectWriter CreatePooled(out byte[] buffer, int minimumLength = DefaultBufferSize)
    {
        buffer = ArrayPool<byte>.Shared.Rent(minimumLength);
        return new VbaProjectWriter(buffer) { _isRented = true };
    }

    /// <summary>
    /// Disposes the writer, returning the buffer to the pool if rented.
    /// </summary>
    public void Dispose()
    {
        if (_isRented)
        {
            ArrayPool<byte>.Shared.Return(_buffer);
            _isRented = false;
        }
    }

    /// <summary>
    /// Gets the written data as a span.
    /// </summary>
    public ReadOnlySpan<byte> GetData() => _memory.Span.Slice(0, _position);

    /// <summary>
    /// Writes a VBA project to the BIFF8 format.
    /// </summary>
    public int WriteVbaProject(VbaProjectData project)
    {
        if (!project.IsValid())
        {
            return 0;
        }

        // Write the VBA project header
        WriteProjectHeader(project);

        // Write the project information record
        WriteProjectInfo(project);

        // Write the project references
        WriteProjectReferences(project);

        // Write the module records
        WriteModuleRecords(project);

        // Write the project compatibility record
        WriteProjectCompatibility();

        // Write the project code page
        WriteProjectCodePage(project.CodePage);

        // Write the project name
        WriteProjectName(project.Name);

        // Write the project constants
        WriteProjectConstants(project);

        return _position;
    }

    /// <summary>
    /// Writes raw VBA binary data (passthrough from XLSX).
    /// </summary>
    public int WriteRawVbaData(byte[] vbaData)
    {
        if (vbaData == null || vbaData.Length == 0)
        {
            return 0;
        }

        // Copy the raw VBA data
        vbaData.CopyTo(_buffer.AsSpan(_position));
        _position += vbaData.Length;

        return _position;
    }

    private void WriteProjectHeader(VbaProjectData project)
    {
        // VBA project signature
        WriteBytes(new byte[] { 0xCC, 0x61 });

        // Version (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0x0062);
        _position += 2;

        // Reserved (2 bytes)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 0x0000);
        _position += 2;
    }

    private void WriteProjectInfo(VbaProjectData project)
    {
        // PROJECTINFO record (0x0001)
        const ushort recordId = 0x0001;

        var startPosition = _position;
        _position += 4; // Reserve for header

        // Version major
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)project.VersionMajor);
        _position += 2;

        // Version minor
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)project.VersionMinor);
        _position += 2;

        // LCID
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)project.Lcid);
        _position += 4;

        // LCID module
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)project.LcidModule);
        _position += 4;

        // Module count
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)project.ModuleCount);
        _position += 2;

        // Project protection state
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)project.Protection);
        _position += 4;

        // Password hash (if protected)
        if (project.PasswordHash != null && project.PasswordHash.Length > 0)
        {
            WriteBytes(project.PasswordHash);
        }
        else
        {
            // Write empty password hash
            for (var i = 0; i < 4; i++)
            {
                _buffer[_position++] = 0;
            }
        }

        // Write record header
        var recordLength = (ushort)(_position - startPosition - 4);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(startPosition), recordId);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(startPosition + 2), recordLength);
    }

    private void WriteProjectReferences(VbaProjectData project)
    {
        foreach (var reference in project.References)
        {
            WriteReference(reference);
        }
    }

    private void WriteReference(VbaReference reference)
    {
        // REFERENCENAME record (0x0016)
        const ushort refNameId = 0x0016;

        var nameBytes = Encoding.UTF8.GetBytes(reference.Name);
        var nameLength = (ushort)(nameBytes.Length + 1);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), refNameId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(nameLength + 3));
        _position += 2;

        // Name length
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), nameLength);
        _position += 2;

        // Reserved
        _buffer[_position++] = 0x48;

        // Name
        nameBytes.CopyTo(_buffer.AsSpan(_position));
        _position += nameBytes.Length;
        _buffer[_position++] = 0; // Null terminator

        // If it's a type library reference, write the GUID
        if (reference.IsTypeLibrary)
        {
            WriteReferenceControl(reference);
        }
        else if (reference.IsProjectReference)
        {
            WriteReferenceProject(reference);
        }
    }

    private void WriteReferenceControl(VbaReference reference)
    {
        // REFERENCECONTROL record (0x002F)
        const ushort recordId = 0x002F;

        var startPosition = _position;
        _position += 4; // Reserve for header

        // Original type lib GUID
        var guidBytes = Guid.Parse(reference.Guid!).ToByteArray();
        WriteBytes(guidBytes);

        // Reserved (4 bytes)
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), 0x00000000);
        _position += 4;

        // Cookie
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)reference.ControlCookie);
        _position += 4;

        // Major version
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)reference.MajorVersion);
        _position += 4;

        // Minor version
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)reference.MinorVersion);
        _position += 2;

        // Write record header
        var recordLength = (ushort)(_position - startPosition - 4);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(startPosition), recordId);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(startPosition + 2), recordLength);
    }

    private void WriteReferenceProject(VbaReference reference)
    {
        // REFERENCEPROJECT record (0x000E)
        const ushort recordId = 0x000E;

        var startPosition = _position;
        _position += 4; // Reserve for header

        // Project path
        var pathBytes = Encoding.UTF8.GetBytes(reference.Path!);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)pathBytes.Length);
        _position += 4;
        pathBytes.CopyTo(_buffer.AsSpan(_position));
        _position += pathBytes.Length;

        // Reserved (4 bytes)
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), 0x00000000);
        _position += 4;

        // Major version
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)reference.MajorVersion);
        _position += 4;

        // Minor version
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)reference.MinorVersion);
        _position += 2;

        // Write record header
        var recordLength = (ushort)(_position - startPosition - 4);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(startPosition), recordId);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(startPosition + 2), recordLength);
    }

    private void WriteModuleRecords(VbaProjectData project)
    {
        foreach (var module in project.Modules)
        {
            WriteModuleRecord(module);
        }
    }

    private void WriteModuleRecord(VbaModule module)
    {
        // MODULENAME record (0x0019)
        WriteModuleName(module.Name);

        // MODULESTREAMNAME record (0x001A)
        WriteModuleStreamName(module.StreamName ?? module.Name);

        // MODuledocSTRING record (0x001C)
        if (!string.IsNullOrEmpty(module.Description))
        {
            WriteModuleDocString(module.Description);
        }

        // MODULEOFFSET record (0x0031)
        WriteModuleOffset(module.Offset);

        // MODULEOPTIONS record (0x001F)
        WriteModuleOptions(module);

        // If the module has code, write it
        if (!string.IsNullOrEmpty(module.Code))
        {
            WriteModuleCode(module.Code);
        }
    }

    private void WriteModuleName(string name)
    {
        const ushort recordId = 0x0019;

        var nameBytes = Encoding.UTF8.GetBytes(name);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)nameBytes.Length);
        _position += 2;

        nameBytes.CopyTo(_buffer.AsSpan(_position));
        _position += nameBytes.Length;
    }

    private void WriteModuleStreamName(string streamName)
    {
        const ushort recordId = 0x001A;

        var nameBytes = Encoding.UTF8.GetBytes(streamName);
        var unicodeBytes = Encoding.Unicode.GetBytes(streamName);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(nameBytes.Length + unicodeBytes.Length + 4));
        _position += 2;

        // Length of ASCII name
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)nameBytes.Length);
        _position += 2;

        // Reserved
        _buffer[_position++] = 0x48;

        // ASCII name
        nameBytes.CopyTo(_buffer.AsSpan(_position));
        _position += nameBytes.Length;

        // Unicode name
        unicodeBytes.CopyTo(_buffer.AsSpan(_position));
        _position += unicodeBytes.Length;
    }

    private void WriteModuleDocString(string description)
    {
        const ushort recordId = 0x001C;

        var descBytes = Encoding.UTF8.GetBytes(description);
        var unicodeBytes = Encoding.Unicode.GetBytes(description);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(descBytes.Length + unicodeBytes.Length + 4));
        _position += 2;

        // Length of ASCII description
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)descBytes.Length);
        _position += 2;

        // Reserved
        _buffer[_position++] = 0x48;

        // ASCII description
        descBytes.CopyTo(_buffer.AsSpan(_position));
        _position += descBytes.Length;

        // Unicode description
        unicodeBytes.CopyTo(_buffer.AsSpan(_position));
        _position += unicodeBytes.Length;
    }

    private void WriteModuleOffset(int offset)
    {
        const ushort recordId = 0x0031;

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 4);
        _position += 2;

        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), (uint)offset);
        _position += 4;
    }

    private void WriteModuleOptions(VbaModule module)
    {
        const ushort recordId = 0x001F;

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 4);
        _position += 2;

        // Module options flags
        uint flags = 0;
        if (module.IsPrivate) flags |= 0x00000001;
        if (module.IsGlobal) flags |= 0x00000002;

        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), flags);
        _position += 4;
    }

    private void WriteModuleCode(string code)
    {
        // MODULECODE record (0x001D)
        const ushort recordId = 0x001D;

        var codeBytes = Encoding.UTF8.GetBytes(code);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)codeBytes.Length);
        _position += 2;

        codeBytes.CopyTo(_buffer.AsSpan(_position));
        _position += codeBytes.Length;
    }

    private void WriteProjectCompatibility()
    {
        // PROJECTCOMPATVERSION record (0x004A)
        const ushort recordId = 0x004A;

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 4);
        _position += 2;

        // Compatibility version
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_position), 0x0000006E);
        _position += 4;
    }

    private void WriteProjectCodePage(int codePage)
    {
        // PROJECTCODEPAGE record (0x0003)
        const ushort recordId = 0x0003;

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), 2);
        _position += 2;

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)codePage);
        _position += 2;
    }

    private void WriteProjectName(string name)
    {
        // PROJECTNAME record (0x0004)
        const ushort recordId = 0x0004;

        var nameBytes = Encoding.UTF8.GetBytes(name);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)nameBytes.Length);
        _position += 2;

        nameBytes.CopyTo(_buffer.AsSpan(_position));
        _position += nameBytes.Length;
    }

    private void WriteProjectConstants(VbaProjectData project)
    {
        // PROJECTCONSTANTS record (0x000C)
        const ushort recordId = 0x000C;

        // Build constants string
        var constants = $"vbExl=\"{project.VersionMajor}.{project.VersionMinor}\"";
        var constBytes = Encoding.UTF8.GetBytes(constants);
        var unicodeBytes = Encoding.Unicode.GetBytes(constants);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), recordId);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)(constBytes.Length + unicodeBytes.Length + 4));
        _position += 2;

        // Length of ASCII constants
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_position), (ushort)constBytes.Length);
        _position += 2;

        // Reserved
        _buffer[_position++] = 0x48;

        // ASCII constants
        constBytes.CopyTo(_buffer.AsSpan(_position));
        _position += constBytes.Length;

        // Unicode constants
        unicodeBytes.CopyTo(_buffer.AsSpan(_position));
        _position += unicodeBytes.Length;
    }

    private void WriteBytes(byte[] data)
    {
        data.CopyTo(_buffer.AsSpan(_position));
        _position += data.Length;
    }
}

/// <summary>
/// Helper class for writing VBA project to OLE compound file storage.
/// </summary>
public static class VbaStorageWriter
{
    /// <summary>
    /// Creates the _VBA_PROJECT_CUR storage structure for BIFF8.
    /// </summary>
    public static byte[] CreateVbaStorage(VbaBinaryData vbaData)
    {
        // In a full implementation, this would create the complete
        // OLE compound file structure for the VBA project
        // For now, return the raw data
        return vbaData.ProjectData ?? Array.Empty<byte>();
    }

    /// <summary>
    /// Creates a minimal VBA project for testing.
    /// </summary>
    public static byte[] CreateMinimalProject(string projectName)
    {
        using var writer = VbaProjectWriter.CreatePooled(out var buffer, 65536);

        var project = new VbaProjectData
        {
            Name = projectName,
            Modules =
            {
                new VbaModule
                {
                    Name = "Module1",
                    Type = VbaModuleType.Standard,
                    Code = $"Attribute VB_Name = \"Module1\"\r\n"
                }
            }
        };

        writer.WriteVbaProject(project);
        return writer.GetData().ToArray();
    }
}
