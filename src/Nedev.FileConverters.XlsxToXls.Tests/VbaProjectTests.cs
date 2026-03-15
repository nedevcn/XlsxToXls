using System;
using System.Collections.Generic;
using System.Linq;
using Nedev.FileConverters.XlsxToXls.Internal;
using Xunit;

namespace Nedev.FileConverters.XlsxToXls.Tests
{
    public class VbaProjectTests
    {
        #region VbaProjectData Tests

        [Fact]
        public void VbaProjectData_DefaultValues()
        {
            var project = new VbaProjectData();

            Assert.Equal("VBAProject", project.Name);
            Assert.Null(project.Description);
            Assert.Null(project.HelpFile);
            Assert.Equal(0, project.HelpContextId);
            Assert.Equal(VbaProtection.None, project.Protection);
            Assert.False(project.IsLocked);
            Assert.Null(project.PasswordHash);
            Assert.NotNull(project.Modules);
            Assert.Empty(project.Modules);
            Assert.NotNull(project.References);
            Assert.Empty(project.References);
            Assert.Null(project.ProjectStream);
            Assert.Null(project.ProjectStorage);
            Assert.Equal(7, project.VersionMajor);
            Assert.Equal(1, project.VersionMinor);
            Assert.Equal(1252, project.CodePage);
            Assert.Equal(1033, project.Lcid);
            Assert.Equal(1033, project.LcidModule);
            Assert.Equal(0, project.ModuleCount);
            Assert.False(project.HasModules);
        }

        [Fact]
        public void VbaProjectData_AddStandardModule()
        {
            var project = new VbaProjectData();
            var module = project.AddStandardModule("TestModule", "Sub Test(): End Sub");

            Assert.Single(project.Modules);
            Assert.Equal("TestModule", module.Name);
            Assert.Equal(VbaModuleType.Standard, module.Type);
            Assert.Equal("Sub Test(): End Sub", module.Code);
        }

        [Fact]
        public void VbaProjectData_AddClassModule()
        {
            var project = new VbaProjectData();
            var module = project.AddClassModule("TestClass", "Public Value As String", true);

            Assert.Single(project.Modules);
            Assert.Equal("TestClass", module.Name);
            Assert.Equal(VbaModuleType.Class, module.Type);
            Assert.True(module.IsGlobal);
        }

        [Fact]
        public void VbaProjectData_AddWorksheetModule()
        {
            var project = new VbaProjectData();
            var module = project.AddWorksheetModule("Sheet1", 0, "Private Sub Worksheet_Activate(): End Sub");

            Assert.Single(project.Modules);
            Assert.Equal("Sheet1", module.Name);
            Assert.Equal(VbaModuleType.Worksheet, module.Type);
            Assert.Equal(0, module.SheetIndex);
        }

        [Fact]
        public void VbaProjectData_AddWorkbookModule()
        {
            var project = new VbaProjectData();
            var module = project.AddWorkbookModule("Private Sub Workbook_Open(): End Sub");

            Assert.Single(project.Modules);
            Assert.Equal("ThisWorkbook", module.Name);
            Assert.Equal(VbaModuleType.Workbook, module.Type);
        }

        [Fact]
        public void VbaProjectData_AddReference()
        {
            var project = new VbaProjectData();
            project.AddReference("Excel", "{00020813-0000-0000-C000-000000000046}", 1, 9);

            Assert.Single(project.References);
            Assert.Equal("Excel", project.References[0].Name);
            Assert.True(project.References[0].IsTypeLibrary);
        }

        [Fact]
        public void VbaProjectData_GetModule()
        {
            var project = new VbaProjectData();
            project.AddStandardModule("Module1", "");
            project.AddStandardModule("Module2", "");

            var module = project.GetModule("Module1");

            Assert.NotNull(module);
            Assert.Equal("Module1", module.Name);
        }

        [Fact]
        public void VbaProjectData_GetModule_NotFound()
        {
            var project = new VbaProjectData();

            var module = project.GetModule("NonExistent");

            Assert.Null(module);
        }

        [Fact]
        public void VbaProjectData_RemoveModule()
        {
            var project = new VbaProjectData();
            project.AddStandardModule("Module1", "");

            var removed = project.RemoveModule("Module1");

            Assert.True(removed);
            Assert.Empty(project.Modules);
        }

        [Fact]
        public void VbaProjectData_RemoveModule_NotFound()
        {
            var project = new VbaProjectData();

            var removed = project.RemoveModule("NonExistent");

            Assert.False(removed);
        }

        [Fact]
        public void VbaProjectData_IsValid_Valid()
        {
            var project = new VbaProjectData();
            project.AddStandardModule("Module1", "");

            Assert.True(project.IsValid());
        }

        [Fact]
        public void VbaProjectData_IsValid_EmptyName()
        {
            var project = new VbaProjectData { Name = "" };

            Assert.False(project.IsValid());
        }

        [Fact]
        public void VbaProjectData_IsValid_DuplicateModuleNames()
        {
            var project = new VbaProjectData();
            project.AddStandardModule("Module1", "");
            project.AddStandardModule("Module1", "");

            Assert.False(project.IsValid());
        }

        #endregion

        #region VbaModule Tests

        [Fact]
        public void VbaModule_DefaultValues()
        {
            var module = new VbaModule();

            Assert.Equal(string.Empty, module.Name);
            Assert.Equal(VbaModuleType.Standard, module.Type);
            Assert.Equal(string.Empty, module.Code);
            Assert.Null(module.StreamName);
            Assert.Null(module.Description);
            Assert.False(module.IsGlobal);
            Assert.False(module.IsPrivate);
            Assert.Equal(-1, module.SheetIndex);
            Assert.Equal(0, module.Offset);
            Assert.Equal(0, module.Length);
        }

        [Fact]
        public void VbaModule_LineCount()
        {
            var module = new VbaModule { Code = "Line1\r\nLine2\r\nLine3" };

            Assert.Equal(3, module.LineCount);
        }

        [Fact]
        public void VbaModule_LineCount_Empty()
        {
            var module = new VbaModule { Code = "" };

            Assert.Equal(0, module.LineCount);
        }

        [Fact]
        public void VbaModule_HasCode()
        {
            var module = new VbaModule { Code = "Sub Test(): End Sub" };

            Assert.True(module.HasCode);
        }

        [Fact]
        public void VbaModule_HasCode_Empty()
        {
            var module = new VbaModule { Code = "" };

            Assert.False(module.HasCode);
        }

        [Fact]
        public void VbaModule_HasCode_Whitespace()
        {
            var module = new VbaModule { Code = "   \r\n   " };

            Assert.False(module.HasCode);
        }

        [Fact]
        public void VbaModule_IsValid()
        {
            var module = new VbaModule { Name = "TestModule" };

            Assert.True(module.IsValid());
        }

        [Fact]
        public void VbaModule_IsValid_EmptyName()
        {
            var module = new VbaModule { Name = "" };

            Assert.False(module.IsValid());
        }

        [Fact]
        public void VbaModule_IsValid_NameTooLong()
        {
            var module = new VbaModule { Name = new string('A', 32) };

            Assert.False(module.IsValid());
        }

        #endregion

        #region VbaReference Tests

        [Fact]
        public void VbaReference_DefaultValues()
        {
            var reference = new VbaReference();

            Assert.Equal(string.Empty, reference.Name);
            Assert.Null(reference.Guid);
            Assert.Equal(0, reference.MajorVersion);
            Assert.Equal(0, reference.MinorVersion);
            Assert.Null(reference.Path);
            Assert.False(reference.IsControl);
            Assert.Null(reference.ControlGuid);
            Assert.Equal(0, reference.ControlCookie);
            Assert.Null(reference.ControlTypeLibGuid);
            Assert.False(reference.IsTypeLibrary);
            Assert.False(reference.IsProjectReference);
        }

        [Fact]
        public void VbaReference_IsTypeLibrary()
        {
            var reference = new VbaReference
            {
                Name = "Excel",
                Guid = "{00020813-0000-0000-C000-000000000046}"
            };

            Assert.True(reference.IsTypeLibrary);
            Assert.False(reference.IsProjectReference);
        }

        [Fact]
        public void VbaReference_IsProjectReference()
        {
            var reference = new VbaReference
            {
                Name = "OtherProject",
                Path = "C:\\Projects\\Other.xls"
            };

            Assert.False(reference.IsTypeLibrary);
            Assert.True(reference.IsProjectReference);
        }

        #endregion

        #region VbaBinaryData Tests

        [Fact]
        public void VbaBinaryData_DefaultValues()
        {
            var data = new VbaBinaryData();

            Assert.Null(data.ProjectData);
            Assert.Null(data.SignatureData);
            Assert.False(data.IsSigned);
            Assert.Null(data.CreationDate);
            Assert.Null(data.ModifiedDate);
            Assert.False(data.IsValid);
        }

        [Fact]
        public void VbaBinaryData_IsSigned()
        {
            var data = new VbaBinaryData
            {
                ProjectData = new byte[] { 1, 2, 3 },
                SignatureData = new byte[] { 4, 5, 6 }
            };

            Assert.True(data.IsSigned);
            Assert.True(data.IsValid);
        }

        [Fact]
        public void VbaBinaryData_IsValid()
        {
            var data = new VbaBinaryData { ProjectData = new byte[] { 1, 2, 3 } };

            Assert.True(data.IsValid);
        }

        [Fact]
        public void VbaBinaryData_IsValid_Empty()
        {
            var data = new VbaBinaryData { ProjectData = Array.Empty<byte>() };

            Assert.False(data.IsValid);
        }

        #endregion

        #region VbaProjectInfo Tests

        [Fact]
        public void VbaProjectInfo_DefaultValues()
        {
            var info = new VbaProjectInfo();

            Assert.Equal(0, info.ModuleCount);
            Assert.False(info.HasWorkbookModule);
            Assert.False(info.HasClassModules);
            Assert.False(info.HasUserForms);
            Assert.False(info.IsProtected);
            Assert.False(info.IsSigned);
        }

        #endregion

        #region Enum Tests

        [Theory]
        [InlineData(VbaModuleType.Standard, 0)]
        [InlineData(VbaModuleType.Class, 1)]
        [InlineData(VbaModuleType.Worksheet, 2)]
        [InlineData(VbaModuleType.Workbook, 3)]
        [InlineData(VbaModuleType.UserForm, 4)]
        [InlineData(VbaModuleType.Designer, 5)]
        public void VbaModuleType_HasCorrectValues(VbaModuleType type, byte expected)
        {
            Assert.Equal(expected, (byte)type);
        }

        [Theory]
        [InlineData(VbaProtection.None, 0)]
        [InlineData(VbaProtection.Locked, 1)]
        [InlineData(VbaProtection.PasswordProtected, 2)]
        public void VbaProtection_HasCorrectValues(VbaProtection protection, byte expected)
        {
            Assert.Equal(expected, (byte)protection);
        }

        #endregion

        #region VbaProjectWriter Tests

        [Fact]
        public void VbaProjectWriter_CreatePooled()
        {
            var writer = VbaProjectWriter.CreatePooled(out var buffer, 65536);
            try
            {
                Assert.True(buffer.Length >= 65536);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void VbaProjectWriter_WriteVbaProject()
        {
            var writer = VbaProjectWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var project = new VbaProjectData
                {
                    Name = "TestProject",
                    Modules =
                    {
                        new VbaModule
                        {
                            Name = "Module1",
                            Type = VbaModuleType.Standard,
                            Code = "Sub Test(): End Sub"
                        }
                    }
                };

                var bytesWritten = writer.WriteVbaProject(project);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void VbaProjectWriter_WriteVbaProject_Invalid()
        {
            var writer = VbaProjectWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var project = new VbaProjectData { Name = "" };

                var bytesWritten = writer.WriteVbaProject(project);
                Assert.Equal(0, bytesWritten);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void VbaProjectWriter_WriteRawVbaData()
        {
            var writer = VbaProjectWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var data = new byte[] { 0xCC, 0x61, 0x00, 0x00 };
                var bytesWritten = writer.WriteRawVbaData(data);

                Assert.Equal(4, bytesWritten);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void VbaProjectWriter_WriteRawVbaData_Null()
        {
            var writer = VbaProjectWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var bytesWritten = writer.WriteRawVbaData(null!);
                Assert.Equal(0, bytesWritten);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void VbaProjectWriter_WriteRawVbaData_Empty()
        {
            var writer = VbaProjectWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var bytesWritten = writer.WriteRawVbaData(Array.Empty<byte>());
                Assert.Equal(0, bytesWritten);
            }
            finally
            {
                writer.Dispose();
            }
        }

        #endregion

        #region VbaStorageWriter Tests

        [Fact]
        public void VbaStorageWriter_CreateVbaStorage()
        {
            var vbaData = new VbaBinaryData
            {
                ProjectData = new byte[] { 1, 2, 3, 4, 5 }
            };

            var storage = VbaStorageWriter.CreateVbaStorage(vbaData);

            Assert.Equal(vbaData.ProjectData, storage);
        }

        [Fact]
        public void VbaStorageWriter_CreateVbaStorage_NullData()
        {
            var vbaData = new VbaBinaryData();

            var storage = VbaStorageWriter.CreateVbaStorage(vbaData);

            Assert.Empty(storage);
        }

        [Fact]
        public void VbaStorageWriter_CreateMinimalProject()
        {
            var data = VbaStorageWriter.CreateMinimalProject("MinimalProject");

            Assert.NotNull(data);
            Assert.True(data.Length > 0);
        }

        #endregion

        #region Complex Project Tests

        [Fact]
        public void VbaProject_ComplexProject()
        {
            var project = new VbaProjectData
            {
                Name = "ComplexProject",
                Description = "A complex VBA project",
                VersionMajor = 7,
                VersionMinor = 1,
                CodePage = 1252,
                Lcid = 1033
            };

            // Add references
            project.AddReference("Excel", "{00020813-0000-0000-C000-000000000046}", 1, 9);
            project.AddReference("Office", "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", 2, 8);
            project.AddReference("Scripting", "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0);

            // Add modules
            project.AddStandardModule("Utils", @"
                Public Function Add(a As Integer, b As Integer) As Integer
                    Add = a + b
                End Function
            ");

            project.AddClassModule("Person", @"
                Private m_Name As String
                
                Public Property Get Name() As String
                    Name = m_Name
                End Property
                
                Public Property Let Name(value As String)
                    m_Name = value
                End Property
            ", true);

            project.AddWorkbookModule(@"
                Private Sub Workbook_Open()
                    MsgBox ""Welcome!""
                End Sub
            ");

            project.AddWorksheetModule("Sheet1", 0, @"
                Private Sub Worksheet_Activate()
                    Range(""A1"").Select
                End Sub
            ");

            // Verify
            Assert.Equal("ComplexProject", project.Name);
            Assert.Equal(3, project.References.Count);
            Assert.Equal(4, project.ModuleCount);
            Assert.True(project.HasModules);
            Assert.True(project.IsValid());
        }

        [Fact]
        public void VbaProject_ProtectedProject()
        {
            var project = new VbaProjectData
            {
                Name = "ProtectedProject",
                Protection = VbaProtection.PasswordProtected,
                PasswordHash = new byte[] { 0x12, 0x34, 0x56, 0x78 }
            };

            project.AddStandardModule("Module1", "");

            Assert.Equal(VbaProtection.PasswordProtected, project.Protection);
            Assert.NotNull(project.PasswordHash);
            Assert.True(project.IsValid());
        }

        #endregion
    }
}
