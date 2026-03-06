using System;
using System.Collections.Generic;
using Nedev.FileConverters.XlsxToXls;
using Nedev.FileConverters.XlsxToXls.Internal;
using Xunit;

namespace Nedev.FileConverters.XlsxToXls.Tests
{
    public class FormulaCompilerTests
    {
        [Fact]
        public void CompileSimpleArithmeticProducesBytes()
        {
            var tokens = XlsxToXlsConverter.CompileFormulaTokens("=1+2", 0, new Dictionary<string, int>());
            Assert.NotNull(tokens);
            Assert.NotEmpty(tokens);
        }

        [Fact]
        public void CompileUnsupportedFunctionThrowsNotSupported()
        {
            Assert.Throws<NotSupportedException>(() => FormulaCompiler.Compile("FOO(1)", 0, new Dictionary<string, int>()));
        }

        [Theory]
        [InlineData("=TRUE", true)]
        [InlineData("=FALSE", false)]
        public void CompileBooleanLiteral(string formula, bool expected)
        {
            var bytes = XlsxToXlsConverter.CompileFormulaTokens(formula, 0, new Dictionary<string, int>());
            // bool tokens start with 0x1D then 0 or 1
            Assert.Equal(2, bytes.Length);
            Assert.Equal(0x1D, bytes[0]);
            Assert.Equal(expected ? 1 : 0, bytes[1]);
        }

        [Fact]
        public void CompileCellReferenceIncludesRefOpcode()
        {
            var bytes = XlsxToXlsConverter.CompileFormulaTokens("=A1", 0, new Dictionary<string, int>());
            // reference opcode is 0x24 for relative A1 on same sheet
            Assert.True(Array.Exists(bytes, b => b == 0x24));
        }

        [Theory]
        [InlineData("SUMIF(A1:A5,\">0\",B1:B5)")]
        [InlineData("COUNTIF(C1:C10,\"foo\")")]
        [InlineData("INDEX(A1:B2,2,1)")]
        [InlineData("MATCH(3,A1:A5,0)")]
        [InlineData("CONCAT(\"a\",\"b\")")]
        [InlineData("TEXT(123,\"0\")")]
        [InlineData("LEN(\"hello\")")]
        [InlineData("TODAY()")]
        [InlineData("NOW()")]
        [InlineData("IFERROR(1/0,\"err\")")]
        public void NewFunctionsCompile(string formula)
        {
            var bytes = XlsxToXlsConverter.CompileFormulaTokens("=" + formula, 0, new Dictionary<string, int>());
            Assert.NotNull(bytes);
            Assert.NotEmpty(bytes);
        }

        [Fact]
        public void ExplicitListTokenBuilderProducesExpectedFormat()
        {
            var tok = FormulaCompiler.BuildExplicitListToken("a,b,c");
            Assert.NotEmpty(tok);
            // first byte should be 0x17 (tStr)
            Assert.Equal(0x17, tok[0]);
        }

        [Fact]
        public void WriteSheetLogsOnCompileException()
        {
            // create sheet with one formula cell that will throw (unsupported function)
            var rows = new List<Internal.RowData>
            {
                new Internal.RowData(
                    RowIndex: 0,
                    Cells: new[]
                    {
                        new Internal.CellData(
                            Row: 0,
                            Col: 0,
                            Kind: CellKind.Formula,
                            CachedKind: CellKind.Empty,
                            Value: null,
                            SstIndex: -1,
                            StyleIndex: 0,
                            Formula: "FOO(1)")
                    },
                    Height: 0,
                    Hidden: false)
            };
            var sheet = new Internal.SheetData("Sheet1", rows, new List<Internal.ColInfo>(), new List<Internal.MergeRange>(), null,
                new List<int>(), new List<int>(), null, null, null, null, new List<Internal.HyperlinkInfo>(), new List<Internal.CommentInfo>(),
                new List<Internal.DataValidationInfo>(), new List<Internal.ConditionalFormatInfo>(), 0);
            var shared = new List<ReadOnlyMemory<char>>();
            var styles = new Internal.StylesData();
            styles.EnsureMinFonts();

            var logMessages = new List<string>();
            var buf = new byte[1024];
            var sheetIndexByName = new Dictionary<string,int>();
            int written = XlsxToXlsConverter.WriteSheet(buf, sheet, shared, styles, 0, sheetIndexByName, msg => logMessages.Add(msg));
            Assert.NotEmpty(logMessages);
            Assert.Contains("Formula compilation failed", logMessages[0]);
            Assert.True(written > 0, "some bytes should still be written even if formula failed");
        }
    }
}