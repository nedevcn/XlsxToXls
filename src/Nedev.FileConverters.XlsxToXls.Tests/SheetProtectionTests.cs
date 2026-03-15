using System;
using System.Collections.Generic;
using Nedev.FileConverters.XlsxToXls.Internal;
using Xunit;

namespace Nedev.FileConverters.XlsxToXls.Tests
{
    public class SheetProtectionTests
    {
        [Fact]
        public void SheetProtection_DefaultValues()
        {
            var protection = new SheetProtection();

            Assert.False(protection.IsProtected);
            Assert.Equal(0, protection.PasswordHash);
            Assert.NotNull(protection.Options);
            Assert.True(protection.Options.LockCells);
            Assert.False(protection.Options.HideCells);
            // AllowEditObjects and AllowEditScenarios default to false in the class,
            // but in XLSX they are true when not specified (attribute not present means allowed)
            Assert.False(protection.AllowEditObjects);
            Assert.False(protection.AllowEditScenarios);
        }

        [Fact]
        public void SheetProtection_WithPassword()
        {
            var protection = new SheetProtection
            {
                IsProtected = true,
                PasswordHash = 12345
            };

            Assert.True(protection.IsProtected);
            Assert.Equal(12345, protection.PasswordHash);
        }

        [Fact]
        public void SheetProtection_AllowOptions()
        {
            var protection = new SheetProtection
            {
                IsProtected = true,
                AllowEditObjects = false,
                AllowEditScenarios = false,
                Options = new ProtectionOptions
                {
                    AllowFormatCells = false,
                    AllowFormatColumns = false,
                    AllowFormatRows = false,
                    AllowInsertColumns = false,
                    AllowInsertRows = false,
                    AllowInsertHyperlinks = false,
                    AllowDeleteColumns = false,
                    AllowDeleteRows = false,
                    AllowSelectLockedCells = true,
                    AllowSort = false,
                    AllowAutoFilter = false,
                    AllowPivotTables = false,
                    AllowSelectUnlockedCells = true
                }
            };

            Assert.False(protection.AllowEditObjects);
            Assert.False(protection.AllowEditScenarios);
            Assert.False(protection.Options.AllowFormatCells);
            Assert.False(protection.Options.AllowSort);
            Assert.True(protection.Options.AllowSelectLockedCells);
            Assert.True(protection.Options.AllowSelectUnlockedCells);
        }

        [Fact]
        public void WorkbookProtection_DefaultValues()
        {
            var protection = new WorkbookProtection();

            Assert.False(protection.ProtectStructure);
            Assert.False(protection.ProtectWindows);
            Assert.Equal(0, protection.PasswordHash);
        }

        [Fact]
        public void WorkbookProtection_WithSettings()
        {
            var protection = new WorkbookProtection
            {
                ProtectStructure = true,
                ProtectWindows = true,
                PasswordHash = 54321
            };

            Assert.True(protection.ProtectStructure);
            Assert.True(protection.ProtectWindows);
            Assert.Equal(54321, protection.PasswordHash);
        }

        [Theory]
        [InlineData("")]
        [InlineData("password")]
        [InlineData("123456")]
        [InlineData("ExcelPassword")]
        public void PasswordHasher_ComputesHash(string password)
        {
            var hash = PasswordHasher.ComputePasswordHash(password);

            // Empty password should return 0
            if (string.IsNullOrEmpty(password))
            {
                Assert.Equal(0, hash);
            }
            else
            {
                // Non-empty password should return non-zero hash
                Assert.NotEqual(0, hash);
            }
        }

        [Fact]
        public void PasswordHasher_SamePassword_SameHash()
        {
            var hash1 = PasswordHasher.ComputePasswordHash("test123");
            var hash2 = PasswordHasher.ComputePasswordHash("test123");

            Assert.Equal(hash1, hash2);
        }

        [Fact]
        public void PasswordHasher_DifferentPassword_DifferentHash()
        {
            var hash1 = PasswordHasher.ComputePasswordHash("password1");
            var hash2 = PasswordHasher.ComputePasswordHash("password2");

            Assert.NotEqual(hash1, hash2);
        }

        [Fact]
        public void SheetProtectionInfo_DefaultValues()
        {
            var info = new SheetProtectionInfo(
                false, 0, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true);

            Assert.False(info.IsProtected);
            Assert.Equal(0, info.PasswordHash);
            Assert.True(info.AllowEditObjects);
            Assert.True(info.AllowEditScenarios);
            Assert.True(info.AllowFormatCells);
            Assert.True(info.AllowSelectLockedCells);
            Assert.True(info.AllowSelectUnlockedCells);
        }

        [Fact]
        public void SheetProtectionInfo_WithProtection()
        {
            var info = new SheetProtectionInfo(
                true, 12345, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false);

            Assert.True(info.IsProtected);
            Assert.Equal(12345, info.PasswordHash);
            Assert.False(info.AllowEditObjects);
            Assert.False(info.AllowEditScenarios);
            Assert.False(info.AllowFormatCells);
            Assert.False(info.AllowSelectLockedCells);
            Assert.False(info.AllowSelectUnlockedCells);
        }

        [Fact]
        public void WorkbookProtectionInfo_DefaultValues()
        {
            var info = new WorkbookProtectionInfo(false, false, 0);

            Assert.False(info.ProtectStructure);
            Assert.False(info.ProtectWindows);
            Assert.Equal(0, info.PasswordHash);
        }

        [Fact]
        public void WorkbookProtectionInfo_WithProtection()
        {
            var info = new WorkbookProtectionInfo(true, true, 9999);

            Assert.True(info.ProtectStructure);
            Assert.True(info.ProtectWindows);
            Assert.Equal(9999, info.PasswordHash);
        }

        [Fact]
        public void SheetProtectionWriter_WritesProtectionRecords()
        {
            var buffer = new byte[1024];
            var writer = new SheetProtectionWriter(buffer.AsSpan());

            var protection = new SheetProtectionInfo(
                true, 12345, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true);

            var position = writer.WriteAllProtectionRecords(protection);

            Assert.True(position > 0);
        }

        [Fact]
        public void SheetProtectionWriter_NoProtection_WritesNothing()
        {
            var buffer = new byte[1024];
            var writer = new SheetProtectionWriter(buffer.AsSpan());

            var protection = new SheetProtectionInfo(
                false, 0, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true);

            var position = writer.WriteAllProtectionRecords(protection);

            Assert.Equal(0, position);
        }

        [Fact]
        public void SheetProtectionWriter_NullProtection_WritesNothing()
        {
            var buffer = new byte[1024];
            var writer = new SheetProtectionWriter(buffer.AsSpan());

            var position = writer.WriteAllProtectionRecords(null);

            Assert.Equal(0, position);
        }

        [Fact]
        public void WorkbookProtectionWriter_WritesProtectionRecords()
        {
            var buffer = new byte[1024];
            var writer = new WorkbookProtectionWriter(buffer.AsSpan());

            var protection = new WorkbookProtectionInfo(true, true, 12345);

            var position = writer.WriteAllProtectionRecords(protection);

            Assert.True(position > 0);
        }

        [Fact]
        public void WorkbookProtectionWriter_NoProtection_WritesNothing()
        {
            var buffer = new byte[1024];
            var writer = new WorkbookProtectionWriter(buffer.AsSpan());

            var protection = new WorkbookProtectionInfo(false, false, 0);

            var position = writer.WriteAllProtectionRecords(protection);

            Assert.Equal(0, position);
        }

        [Fact]
        public void WorkbookProtectionWriter_NullProtection_WritesNothing()
        {
            var buffer = new byte[1024];
            var writer = new WorkbookProtectionWriter(buffer.AsSpan());

            var position = writer.WriteAllProtectionRecords(null);

            Assert.Equal(0, position);
        }

        [Fact]
        public void ProtectionOptions_AllProperties()
        {
            var options = new ProtectionOptions
            {
                LockCells = true,
                HideCells = true,
                AllowFormatCells = false,
                AllowFormatColumns = false,
                AllowFormatRows = false,
                AllowInsertColumns = false,
                AllowInsertRows = false,
                AllowInsertHyperlinks = false,
                AllowDeleteColumns = false,
                AllowDeleteRows = false,
                AllowSelectLockedCells = true,
                AllowSort = false,
                AllowAutoFilter = false,
                AllowPivotTables = false,
                AllowSelectUnlockedCells = true
            };

            Assert.True(options.LockCells);
            Assert.True(options.HideCells);
            Assert.False(options.AllowFormatCells);
            Assert.False(options.AllowFormatColumns);
            Assert.False(options.AllowFormatRows);
            Assert.False(options.AllowInsertColumns);
            Assert.False(options.AllowInsertRows);
            Assert.False(options.AllowInsertHyperlinks);
            Assert.False(options.AllowDeleteColumns);
            Assert.False(options.AllowDeleteRows);
            Assert.True(options.AllowSelectLockedCells);
            Assert.False(options.AllowSort);
            Assert.False(options.AllowAutoFilter);
            Assert.False(options.AllowPivotTables);
            Assert.True(options.AllowSelectUnlockedCells);
        }

        [Fact]
        public void SheetProtection_CompleteScenario()
        {
            // Simulate creating a protected sheet with password
            var password = "MySecretPassword";
            var passwordHash = PasswordHasher.ComputePasswordHash(password);

            var protection = new SheetProtection
            {
                IsProtected = true,
                PasswordHash = passwordHash,
                AllowEditObjects = false,
                AllowEditScenarios = false,
                Options = new ProtectionOptions
                {
                    LockCells = true,
                    HideCells = false,
                    AllowFormatCells = false,
                    AllowFormatColumns = false,
                    AllowFormatRows = false,
                    AllowInsertColumns = false,
                    AllowInsertRows = false,
                    AllowInsertHyperlinks = false,
                    AllowDeleteColumns = false,
                    AllowDeleteRows = false,
                    AllowSelectLockedCells = true,
                    AllowSort = false,
                    AllowAutoFilter = false,
                    AllowPivotTables = false,
                    AllowSelectUnlockedCells = true
                }
            };

            Assert.True(protection.IsProtected);
            Assert.NotEqual(0, protection.PasswordHash);
            Assert.False(protection.AllowEditObjects);
            Assert.False(protection.AllowEditScenarios);
            Assert.True(protection.Options.LockCells);
            Assert.True(protection.Options.AllowSelectLockedCells);
            Assert.True(protection.Options.AllowSelectUnlockedCells);
        }
    }
}
