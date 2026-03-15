using System;
using System.Collections.Generic;
using System.Linq;
using Nedev.FileConverters.XlsxToXls.Internal;
using Xunit;

namespace Nedev.FileConverters.XlsxToXls.Tests
{
    public class HyperlinkTests
    {
        #region HyperlinkData Tests

        [Fact]
        public void HyperlinkData_DefaultValues()
        {
            var hyperlink = new HyperlinkData();

            Assert.Equal(0, hyperlink.Row);
            Assert.Equal(0, hyperlink.Column);
            Assert.Equal(HyperlinkType.Url, hyperlink.Type);
            Assert.Equal(string.Empty, hyperlink.Target);
            Assert.Null(hyperlink.DisplayText);
            Assert.Null(hyperlink.Tooltip);
            Assert.Null(hyperlink.ScreenTip);
            Assert.Null(hyperlink.Location);
            Assert.Null(hyperlink.EmailSubject);
            Assert.False(hyperlink.IsOneClick);
            Assert.Equal(-1, hyperlink.RefIndex);
        }

        [Fact]
        public void HyperlinkData_CreateUrl()
        {
            var hyperlink = HyperlinkData.CreateUrl(5, 10, "https://example.com", "Example Site", "Click to visit");

            Assert.Equal(5, hyperlink.Row);
            Assert.Equal(10, hyperlink.Column);
            Assert.Equal(HyperlinkType.Url, hyperlink.Type);
            Assert.Equal("https://example.com", hyperlink.Target);
            Assert.Equal("Example Site", hyperlink.DisplayText);
            Assert.Equal("Click to visit", hyperlink.Tooltip);
        }

        [Fact]
        public void HyperlinkData_CreateFile()
        {
            var hyperlink = HyperlinkData.CreateFile(3, 7, @"C:\Documents\file.pdf", "Open PDF", "Click to open");

            Assert.Equal(3, hyperlink.Row);
            Assert.Equal(7, hyperlink.Column);
            Assert.Equal(HyperlinkType.File, hyperlink.Type);
            Assert.Equal(@"C:\Documents\file.pdf", hyperlink.Target);
            Assert.Equal("Open PDF", hyperlink.DisplayText);
        }

        [Fact]
        public void HyperlinkData_CreateEmail()
        {
            var hyperlink = HyperlinkData.CreateEmail(2, 5, "test@example.com", "Hello", "Send Email");

            Assert.Equal(2, hyperlink.Row);
            Assert.Equal(5, hyperlink.Column);
            Assert.Equal(HyperlinkType.Email, hyperlink.Type);
            Assert.Equal("mailto:test@example.com?subject=Hello", hyperlink.Target);
            Assert.Equal("Hello", hyperlink.EmailSubject);
            Assert.Equal("Send Email", hyperlink.DisplayText);
        }

        [Fact]
        public void HyperlinkData_CreateEmailWithoutSubject()
        {
            var hyperlink = HyperlinkData.CreateEmail(0, 0, "contact@company.com");

            Assert.Equal(HyperlinkType.Email, hyperlink.Type);
            Assert.Equal("mailto:contact@company.com", hyperlink.Target);
            Assert.Null(hyperlink.EmailSubject);
            Assert.Equal("contact@company.com", hyperlink.DisplayText);
        }

        [Fact]
        public void HyperlinkData_CreateDocument()
        {
            var hyperlink = HyperlinkData.CreateDocument(1, 2, @"C:\Docs\report.docx", "Section1", "Go to Section");

            Assert.Equal(1, hyperlink.Row);
            Assert.Equal(2, hyperlink.Column);
            Assert.Equal(HyperlinkType.Document, hyperlink.Type);
            Assert.Equal(@"C:\Docs\report.docx", hyperlink.Target);
            Assert.Equal("Section1", hyperlink.Location);
            Assert.Equal("Go to Section", hyperlink.DisplayText);
        }

        [Fact]
        public void HyperlinkData_CreateInternal()
        {
            var hyperlink = HyperlinkData.CreateInternal(0, 0, "Sheet2!A1", "Go to Sheet2");

            Assert.Equal(0, hyperlink.Row);
            Assert.Equal(0, hyperlink.Column);
            Assert.Equal(HyperlinkType.Internal, hyperlink.Type);
            Assert.Equal("Sheet2!A1", hyperlink.Target);
            Assert.Equal("Go to Sheet2", hyperlink.DisplayText);
        }

        [Fact]
        public void HyperlinkData_GetCellReference()
        {
            var hyperlink = HyperlinkData.CreateUrl(0, 0, "https://example.com");
            Assert.Equal("A1", hyperlink.GetCellReference());

            hyperlink = HyperlinkData.CreateUrl(9, 25, "https://example.com");
            Assert.Equal("Z10", hyperlink.GetCellReference());

            hyperlink = HyperlinkData.CreateUrl(99, 26, "https://example.com");
            Assert.Equal("AA100", hyperlink.GetCellReference());

            hyperlink = HyperlinkData.CreateUrl(0, 701, "https://example.com");
            Assert.Equal("ZZ1", hyperlink.GetCellReference());
        }

        [Theory]
        [InlineData("https://example.com", true)]
        [InlineData("http://example.com", true)]
        [InlineData("ftp://example.com", false)]
        [InlineData("not-a-url", false)]
        [InlineData("", false)]
        public void HyperlinkData_IsValid_Url(string url, bool expectedValid)
        {
            var hyperlink = HyperlinkData.CreateUrl(0, 0, url);
            Assert.Equal(expectedValid, hyperlink.IsValid());
        }

        [Fact]
        public void HyperlinkData_IsValid_InvalidRow()
        {
            var hyperlink = new HyperlinkData { Row = -1, Column = 0, Target = "https://example.com", Type = HyperlinkType.Url };
            Assert.False(hyperlink.IsValid());
        }

        [Fact]
        public void HyperlinkData_IsValid_InvalidColumn()
        {
            var hyperlink = new HyperlinkData { Row = 0, Column = -1, Target = "https://example.com", Type = HyperlinkType.Url };
            Assert.False(hyperlink.IsValid());
        }

        [Fact]
        public void HyperlinkData_IsValid_EmptyTarget()
        {
            var hyperlink = new HyperlinkData { Row = 0, Column = 0, Target = "", Type = HyperlinkType.Url };
            Assert.False(hyperlink.IsValid());
        }

        #endregion

        #region HyperlinkCollection Tests

        [Fact]
        public void HyperlinkCollection_DefaultState()
        {
            var collection = new HyperlinkCollection();

            Assert.Empty(collection.Hyperlinks);
            Assert.Equal(0, collection.Count);
        }

        [Fact]
        public void HyperlinkCollection_Add()
        {
            var collection = new HyperlinkCollection();
            var hyperlink = HyperlinkData.CreateUrl(0, 0, "https://example.com");

            collection.Add(hyperlink);

            Assert.Single(collection.Hyperlinks);
            Assert.Equal(1, collection.Count);
        }

        [Fact]
        public void HyperlinkCollection_Add_ReplaceExisting()
        {
            var collection = new HyperlinkCollection();
            collection.Add(HyperlinkData.CreateUrl(0, 0, "https://old.com"));
            collection.Add(HyperlinkData.CreateUrl(0, 0, "https://new.com"));

            Assert.Single(collection.Hyperlinks);
            Assert.Equal("https://new.com", collection.Hyperlinks[0].Target);
        }

        [Fact]
        public void HyperlinkCollection_Add_NullThrows()
        {
            var collection = new HyperlinkCollection();
            Assert.Throws<ArgumentNullException>(() => collection.Add(null!));
        }

        [Fact]
        public void HyperlinkCollection_RemoveAt()
        {
            var collection = new HyperlinkCollection();
            collection.Add(HyperlinkData.CreateUrl(0, 0, "https://example.com"));

            var removed = collection.RemoveAt(0, 0);

            Assert.True(removed);
            Assert.Empty(collection.Hyperlinks);
        }

        [Fact]
        public void HyperlinkCollection_RemoveAt_NotFound()
        {
            var collection = new HyperlinkCollection();

            var removed = collection.RemoveAt(0, 0);

            Assert.False(removed);
        }

        [Fact]
        public void HyperlinkCollection_GetAt()
        {
            var collection = new HyperlinkCollection();
            var hyperlink = HyperlinkData.CreateUrl(5, 10, "https://example.com");
            collection.Add(hyperlink);

            var found = collection.GetAt(5, 10);

            Assert.NotNull(found);
            Assert.Equal("https://example.com", found.Target);
        }

        [Fact]
        public void HyperlinkCollection_GetAt_NotFound()
        {
            var collection = new HyperlinkCollection();

            var found = collection.GetAt(0, 0);

            Assert.Null(found);
        }

        [Fact]
        public void HyperlinkCollection_HasHyperlink()
        {
            var collection = new HyperlinkCollection();
            collection.Add(HyperlinkData.CreateUrl(3, 7, "https://example.com"));

            Assert.True(collection.HasHyperlink(3, 7));
            Assert.False(collection.HasHyperlink(0, 0));
        }

        [Fact]
        public void HyperlinkCollection_Clear()
        {
            var collection = new HyperlinkCollection();
            collection.Add(HyperlinkData.CreateUrl(0, 0, "https://example.com"));
            collection.Add(HyperlinkData.CreateUrl(1, 1, "https://example2.com"));

            collection.Clear();

            Assert.Empty(collection.Hyperlinks);
            Assert.Equal(0, collection.Count);
        }

        [Fact]
        public void HyperlinkCollection_GetByType()
        {
            var collection = new HyperlinkCollection();
            collection.Add(HyperlinkData.CreateUrl(0, 0, "https://example.com"));
            collection.Add(HyperlinkData.CreateFile(1, 0, @"C:\file.txt"));
            collection.Add(HyperlinkData.CreateUrl(2, 0, "https://example2.com"));

            var urlLinks = collection.GetByType(HyperlinkType.Url).ToList();
            var fileLinks = collection.GetByType(HyperlinkType.File).ToList();

            Assert.Equal(2, urlLinks.Count);
            Assert.Single(fileLinks);
        }

        [Fact]
        public void HyperlinkCollection_GetByRow()
        {
            var collection = new HyperlinkCollection();
            collection.Add(HyperlinkData.CreateUrl(0, 0, "https://example.com"));
            collection.Add(HyperlinkData.CreateUrl(0, 1, "https://example2.com"));
            collection.Add(HyperlinkData.CreateUrl(1, 0, "https://example3.com"));

            var row0Links = collection.GetByRow(0).ToList();

            Assert.Equal(2, row0Links.Count);
        }

        [Fact]
        public void HyperlinkCollection_GetByColumn()
        {
            var collection = new HyperlinkCollection();
            collection.Add(HyperlinkData.CreateUrl(0, 0, "https://example.com"));
            collection.Add(HyperlinkData.CreateUrl(1, 0, "https://example2.com"));
            collection.Add(HyperlinkData.CreateUrl(0, 1, "https://example3.com"));

            var col0Links = collection.GetByColumn(0).ToList();

            Assert.Equal(2, col0Links.Count);
        }

        #endregion

        #region CellReferenceConverter Tests

        [Theory]
        [InlineData(0, 0, "A1")]
        [InlineData(0, 1, "B1")]
        [InlineData(0, 25, "Z1")]
        [InlineData(0, 26, "AA1")]
        [InlineData(0, 51, "AZ1")]
        [InlineData(0, 52, "BA1")]
        [InlineData(0, 701, "ZZ1")]
        [InlineData(0, 702, "AAA1")]
        [InlineData(9, 0, "A10")]
        [InlineData(99, 25, "Z100")]
        [InlineData(16383, 255, "IV16384")]
        public void CellReferenceConverter_ToReference(int row, int col, string expected)
        {
            Assert.Equal(expected, CellReferenceConverter.ToReference(row, col));
        }

        [Theory]
        [InlineData("A1", 0, 0)]
        [InlineData("B1", 0, 1)]
        [InlineData("Z1", 0, 25)]
        [InlineData("AA1", 0, 26)]
        [InlineData("AZ1", 0, 51)]
        [InlineData("BA1", 0, 52)]
        [InlineData("ZZ1", 0, 701)]
        [InlineData("AAA1", 0, 702)]
        [InlineData("A10", 9, 0)]
        [InlineData("Z100", 99, 25)]
        [InlineData("a1", 0, 0)]
        [InlineData("A10", 9, 0)]
        public void CellReferenceConverter_FromReference(string reference, int expectedRow, int expectedCol)
        {
            var (row, col) = CellReferenceConverter.FromReference(reference);
            Assert.Equal(expectedRow, row);
            Assert.Equal(expectedCol, col);
        }

        [Theory]
        [InlineData(0, "A")]
        [InlineData(1, "B")]
        [InlineData(25, "Z")]
        [InlineData(26, "AA")]
        [InlineData(51, "AZ")]
        [InlineData(52, "BA")]
        [InlineData(701, "ZZ")]
        [InlineData(702, "AAA")]
        public void CellReferenceConverter_ColumnToLetters(int col, string expected)
        {
            Assert.Equal(expected, CellReferenceConverter.ColumnToLetters(col));
        }

        [Theory]
        [InlineData("A", 0)]
        [InlineData("B", 1)]
        [InlineData("Z", 25)]
        [InlineData("AA", 26)]
        [InlineData("AZ", 51)]
        [InlineData("BA", 52)]
        [InlineData("ZZ", 701)]
        [InlineData("AAA", 702)]
        [InlineData("a", 0)]
        [InlineData("z", 25)]
        public void CellReferenceConverter_LettersToColumn(string letters, int expected)
        {
            Assert.Equal(expected, CellReferenceConverter.LettersToColumn(letters));
        }

        [Fact]
        public void CellReferenceConverter_ToReference_NegativeRow()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => CellReferenceConverter.ToReference(-1, 0));
        }

        [Fact]
        public void CellReferenceConverter_ToReference_NegativeColumn()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => CellReferenceConverter.ToReference(0, -1));
        }

        [Fact]
        public void CellReferenceConverter_FromReference_Null()
        {
            Assert.Throws<ArgumentException>(() => CellReferenceConverter.FromReference(null!));
        }

        [Fact]
        public void CellReferenceConverter_FromReference_Empty()
        {
            Assert.Throws<ArgumentException>(() => CellReferenceConverter.FromReference(""));
        }

        [Fact]
        public void CellReferenceConverter_FromReference_InvalidFormat()
        {
            Assert.Throws<ArgumentException>(() => CellReferenceConverter.FromReference("123"));
        }

        [Fact]
        public void CellReferenceConverter_FromReference_InvalidCharacters()
        {
            Assert.Throws<ArgumentException>(() => CellReferenceConverter.FromReference("A1B"));
        }

        [Fact]
        public void CellReferenceConverter_LettersToColumn_Null()
        {
            Assert.Throws<ArgumentException>(() => CellReferenceConverter.LettersToColumn(null!));
        }

        [Fact]
        public void CellReferenceConverter_LettersToColumn_InvalidCharacters()
        {
            Assert.Throws<ArgumentException>(() => CellReferenceConverter.LettersToColumn("A1"));
        }

        #endregion

        #region HyperlinkType Enum Tests

        [Theory]
        [InlineData(HyperlinkType.Url, 0)]
        [InlineData(HyperlinkType.File, 1)]
        [InlineData(HyperlinkType.Email, 2)]
        [InlineData(HyperlinkType.Document, 3)]
        [InlineData(HyperlinkType.Internal, 4)]
        [InlineData(HyperlinkType.Unc, 5)]
        [InlineData(HyperlinkType.Worksheet, 6)]
        [InlineData(HyperlinkType.NamedRange, 7)]
        public void HyperlinkType_HasCorrectValues(HyperlinkType type, byte expected)
        {
            Assert.Equal(expected, (byte)type);
        }

        #endregion

        #region HyperlinkWriter Tests

        [Fact]
        public void HyperlinkWriter_CreatePooled()
        {
            var writer = HyperlinkWriter.CreatePooled(out var buffer, 65536);
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
        public void HyperlinkWriter_WriteHyperlink_ValidUrl()
        {
            var writer = HyperlinkWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var hyperlink = HyperlinkData.CreateUrl(0, 0, "https://example.com", "Example");
                var bytesWritten = writer.WriteHyperlink(hyperlink);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void HyperlinkWriter_WriteHyperlink_InvalidHyperlink()
        {
            var writer = HyperlinkWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var hyperlink = new HyperlinkData { Row = -1, Column = 0, Target = "test" };
                var bytesWritten = writer.WriteHyperlink(hyperlink);
                // Should return without writing for invalid hyperlink
                Assert.Equal(0, bytesWritten);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void HyperlinkWriter_WriteHyperlinks()
        {
            var writer = HyperlinkWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var collection = new HyperlinkCollection();
                collection.Add(HyperlinkData.CreateUrl(0, 0, "https://example.com"));
                collection.Add(HyperlinkData.CreateUrl(1, 1, "https://example2.com"));
                collection.Add(HyperlinkData.CreateFile(2, 2, @"C:\test.txt"));

                var bytesWritten = writer.WriteHyperlinks(collection);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        #endregion

        #region SimpleHyperlinkWriter Tests

        [Fact]
        public void SimpleHyperlinkWriter_CreatePooled()
        {
            var writer = SimpleHyperlinkWriter.CreatePooled(out var buffer, 65536);
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
        public void SimpleHyperlinkWriter_WriteUrlHyperlink()
        {
            var writer = SimpleHyperlinkWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var bytesWritten = writer.WriteUrlHyperlink(5, 10, "https://example.com", "Example Site");
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void SimpleHyperlinkWriter_WriteUrlHyperlink_WithoutDisplayText()
        {
            var writer = SimpleHyperlinkWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var bytesWritten = writer.WriteUrlHyperlink(0, 0, "https://example.com");
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void SimpleHyperlinkWriter_MultipleHyperlinks()
        {
            var writer = SimpleHyperlinkWriter.CreatePooled(out var buffer, 65536);
            try
            {
                writer.WriteUrlHyperlink(0, 0, "https://example.com", "Site 1");
                writer.WriteUrlHyperlink(1, 0, "https://example2.com", "Site 2");
                writer.WriteUrlHyperlink(2, 0, "https://example3.com");

                var data = writer.GetData();
                Assert.True(data.Length > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        #endregion
    }
}
