using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Nedev.FileConverters.XlsxToXls;
using Xunit;

namespace Nedev.FileConverters.XlsxToXls.Tests
{
    public class AsyncApiTests
    {
        private static string CreateTestXlsx()
        {
            var tempPath = Path.GetTempFileName() + ".xlsx";
            using (var fs = File.Create(tempPath))
            using (var archive = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                // [Content_Types].xml
                var contentTypes = archive.CreateEntry("[Content_Types].xml");
                using (var s = contentTypes.Open())
                {
                    var xml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
  <Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
</Types>";
                    s.Write(Encoding.UTF8.GetBytes(xml));
                }

                // _rels/.rels
                var rels = archive.CreateEntry("_rels/.rels");
                using (var s = rels.Open())
                {
                    var xml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
</Relationships>";
                    s.Write(Encoding.UTF8.GetBytes(xml));
                }

                // xl/_rels/workbook.xml.rels
                var wbRels = archive.CreateEntry("xl/_rels/workbook.xml.rels");
                using (var s = wbRels.Open())
                {
                    var xml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
</Relationships>";
                    s.Write(Encoding.UTF8.GetBytes(xml));
                }

                // xl/workbook.xml
                var wb = archive.CreateEntry("xl/workbook.xml");
                using (var s = wb.Open())
                {
                    var xml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <sheets>
    <sheet name=""Sheet1"" sheetId=""1"" r:id=""rId1"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""/>
  </sheets>
</workbook>";
                    s.Write(Encoding.UTF8.GetBytes(xml));
                }

                // xl/worksheets/sheet1.xml
                var sheet = archive.CreateEntry("xl/worksheets/sheet1.xml");
                using (var s = sheet.Open())
                {
                    var xml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <sheetData>
    <row r=""1"">
      <c r=""A1"" t=""inlineStr""><is><t>Test</t></is></c>
      <c r=""B1""><v>123</v></c>
    </row>
  </sheetData>
</worksheet>";
                    s.Write(Encoding.UTF8.GetBytes(xml));
                }
            }
            return tempPath;
        }

        [Fact]
        public async Task ConvertFileAsync_BasicConversion()
        {
            var inputPath = CreateTestXlsx();
            var outputPath = Path.GetTempFileName() + ".xls";

            try
            {
                await XlsxToXlsConverter.ConvertFileAsync(inputPath, outputPath);

                Assert.True(File.Exists(outputPath));
                var fileInfo = new FileInfo(outputPath);
                Assert.True(fileInfo.Length > 0);

                // Verify it's a valid OLE compound file
                using var fs = File.OpenRead(outputPath);
                var header = new byte[8];
                fs.Read(header, 0, 8);
                Assert.Equal(0xD0, header[0]);
                Assert.Equal(0xCF, header[1]);
                Assert.Equal(0x11, header[2]);
                Assert.Equal(0xE0, header[3]);
            }
            finally
            {
                if (File.Exists(inputPath)) File.Delete(inputPath);
                if (File.Exists(outputPath)) File.Delete(outputPath);
            }
        }

        [Fact]
        public async Task ConvertFileAsync_WithLogging()
        {
            var inputPath = CreateTestXlsx();
            var outputPath = Path.GetTempFileName() + ".xls";
            var logs = new List<string>();

            try
            {
                await XlsxToXlsConverter.ConvertFileAsync(inputPath, outputPath, log => logs.Add(log));

                Assert.True(File.Exists(outputPath));
                Assert.NotEmpty(logs);
            }
            finally
            {
                if (File.Exists(inputPath)) File.Delete(inputPath);
                if (File.Exists(outputPath)) File.Delete(outputPath);
            }
        }

        [Fact]
        public async Task ConvertFileAsync_Cancellation_ThrowsOperationCanceledException()
        {
            var inputPath = CreateTestXlsx();
            var outputPath = Path.GetTempFileName() + ".xls";
            var cts = new CancellationTokenSource();
            cts.Cancel();

            try
            {
                await Assert.ThrowsAsync<OperationCanceledException>(() =>
                    XlsxToXlsConverter.ConvertFileAsync(inputPath, outputPath, cts.Token));
            }
            finally
            {
                if (File.Exists(inputPath)) File.Delete(inputPath);
                if (File.Exists(outputPath)) File.Delete(outputPath);
            }
        }

        [Fact]
        public async Task ConvertFileAsync_NullPath_ThrowsArgumentException()
        {
            await Assert.ThrowsAsync<ArgumentException>(() =>
                XlsxToXlsConverter.ConvertFileAsync(null!, "output.xls"));

            await Assert.ThrowsAsync<ArgumentException>(() =>
                XlsxToXlsConverter.ConvertFileAsync("input.xlsx", null!));
        }
    }
}
