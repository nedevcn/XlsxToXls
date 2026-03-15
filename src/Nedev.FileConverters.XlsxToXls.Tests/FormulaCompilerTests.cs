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
        // Statistical functions
        [InlineData("SUMIF(A1:A5,\">0\",B1:B5)")]
        [InlineData("COUNTIF(C1:C10,\"foo\")")]
        [InlineData("AVERAGEIF(A1:A10,\">0\")")]
        [InlineData("SUMIFS(B1:B10,A1:A10,\">0\")")]
        [InlineData("COUNTIFS(A1:A10,\">0\",B1:B10,\"<100\")")]
        [InlineData("AVERAGEIFS(B1:B10,A1:A10,\">0\")")]
        [InlineData("MAXIFS(B1:B10,A1:A10,\">0\")")]
        [InlineData("MINIFS(B1:B10,A1:A10,\">0\")")]
        [InlineData("STDEV(A1:A10)")]
        [InlineData("STDEVP(A1:A10)")]
        [InlineData("VAR(A1:A10)")]
        [InlineData("VARP(A1:A10)")]
        [InlineData("MEDIAN(A1:A10)")]
        [InlineData("MODE(A1:A10)")]
        [InlineData("RANK(A1,A1:A10)")]
        [InlineData("PERCENTILE(A1:A10,0.5)")]
        // Lookup functions
        [InlineData("INDEX(A1:B2,2,1)")]
        [InlineData("MATCH(3,A1:A5,0)")]
        [InlineData("HLOOKUP(\"x\",A1:C3,2,FALSE)")]
        [InlineData("INDIRECT(\"A1\")")]
        [InlineData("OFFSET(A1,1,1)")]
        [InlineData("CHOOSE(1,A1,B1,C1)")]
        [InlineData("ROW()")]
        [InlineData("COLUMN()")]
        [InlineData("ROWS(A1:A10)")]
        [InlineData("COLUMNS(A1:J1)")]
        // Text functions
        [InlineData("CONCAT(\"a\",\"b\")")]
        [InlineData("TEXT(123,\"0\")")]
        [InlineData("LEN(\"hello\")")]
        [InlineData("LEFT(\"hello\",3)")]
        [InlineData("RIGHT(\"hello\",3)")]
        [InlineData("MID(\"hello\",2,3)")]
        [InlineData("FIND(\"l\",\"hello\")")]
        [InlineData("SUBSTITUTE(\"hello\",\"l\",\"x\")")]
        [InlineData("TRIM(\"  hello  \")")]
        [InlineData("UPPER(\"hello\")")]
        [InlineData("LOWER(\"HELLO\")")]
        [InlineData("PROPER(\"hello world\")")]
        [InlineData("VALUE(\"123\")")]
        [InlineData("REPT(\"*\",5)")]
        [InlineData("TEXTJOIN(\",\",TRUE,A1:A5)")]
        // Date & Time functions
        [InlineData("TODAY()")]
        [InlineData("NOW()")]
        [InlineData("DATE(2024,1,1)")]
        [InlineData("TIME(12,30,0)")]
        [InlineData("DAY(A1)")]
        [InlineData("MONTH(A1)")]
        [InlineData("YEAR(A1)")]
        [InlineData("WEEKDAY(A1)")]
        [InlineData("HOUR(A1)")]
        [InlineData("MINUTE(A1)")]
        [InlineData("SECOND(A1)")]
        [InlineData("DATEDIF(A1,A2,\"D\")")]
        [InlineData("EDATE(A1,1)")]
        [InlineData("EOMONTH(A1,0)")]
        [InlineData("NETWORKDAYS(A1,A2)")]
        [InlineData("WORKDAY(A1,5)")]
        // Math functions
        [InlineData("MOD(10,3)")]
        [InlineData("POWER(2,3)")]
        [InlineData("SQRT(16)")]
        [InlineData("EXP(1)")]
        [InlineData("LN(10)")]
        [InlineData("LOG10(100)")]
        // Financial functions
        [InlineData("PMT(0.05/12,360,200000)")]
        [InlineData("FV(0.05/12,360,-1000)")]
        [InlineData("PV(0.05/12,360,-1000)")]
        [InlineData("NPV(0.1,A1:A5)")]
        [InlineData("IRR(A1:A5)")]
        [InlineData("RATE(360,-1000,200000)")]
        [InlineData("NPER(0.05/12,-1000,200000)")]
        [InlineData("IPMT(0.05/12,1,360,200000)")]
        [InlineData("PPMT(0.05/12,1,360,200000)")]
        [InlineData("CUMIPMT(0.05/12,360,200000,1,12,0)")]
        [InlineData("CUMPRINC(0.05/12,360,200000,1,12,0)")]
        [InlineData("DB(10000,1000,5,1)")]
        [InlineData("DDB(10000,1000,5,1)")]
        [InlineData("SLN(10000,1000,5)")]
        [InlineData("SYD(10000,1000,5,1)")]
        [InlineData("VDB(10000,1000,5,1,2)")]
        // Information functions
        [InlineData("ISBLANK(A1)")]
        [InlineData("ISNUMBER(A1)")]
        [InlineData("ISTEXT(A1)")]
        [InlineData("ISLOGICAL(A1)")]
        [InlineData("ISERROR(A1)")]
        [InlineData("ISNA(A1)")]
        // Logical functions
        [InlineData("TRUE")]
        [InlineData("FALSE")]
        [InlineData("IFERROR(1/0,\"err\")")]
        [InlineData("IFNA(A1,0)")]
        [InlineData("ANDIFS(A1:A10,\">0\")")]
        [InlineData("ORIFS(A1:A10,\">0\")")]
        [InlineData("SWITCH(A1,1,\"One\",2,\"Two\",\"Other\")")]
        [InlineData("IFS(A1>0,\"Positive\",A1<0,\"Negative\",TRUE,\"Zero\")")]
        // Database functions
        [InlineData("DSUM(A1:C10,\"Sales\",E1:F2)")]
        [InlineData("DCOUNT(A1:C10,\"Sales\",E1:F2)")]
        [InlineData("DCOUNTA(A1:C10,\"Sales\",E1:F2)")]
        [InlineData("DAVERAGE(A1:C10,\"Sales\",E1:F2)")]
        [InlineData("DMIN(A1:C10,\"Sales\",E1:F2)")]
        [InlineData("DMAX(A1:C10,\"Sales\",E1:F2)")]
        [InlineData("DSTDEV(A1:C10,\"Sales\",E1:F2)")]
        [InlineData("DSTDEVP(A1:C10,\"Sales\",E1:F2)")]
        [InlineData("DVAR(A1:C10,\"Sales\",E1:F2)")]
        [InlineData("DVARP(A1:C10,\"Sales\",E1:F2)")]
        [InlineData("DGET(A1:C10,\"Sales\",E1:F2)")]
        // Additional aggregation
        [InlineData("COUNTA(A1:A10)")]
        [InlineData("PRODUCT(A1:A10)")]
        public void ExtendedFunctionsCompile(string formula)
        {
            var bytes = XlsxToXlsConverter.CompileFormulaTokens("=" + formula, 0, new Dictionary<string, int>());
            Assert.NotNull(bytes);
            Assert.NotEmpty(bytes);
        }

        [Fact]
        public void TokenToAstConversionProducesFunctionNode()
        {
            var tokens = typeof(Nedev.FileConverters.XlsxToXls.Internal.FormulaCompiler)
                .GetMethod("Tokenize", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
                ?.Invoke(null, new object[] { "=SUM(1,2)", 0, new Dictionary<string, int>() });
            // call AST builder via reflection
            var ast = typeof(Nedev.FileConverters.XlsxToXls.Internal.FormulaCompiler)
                .GetMethod("AstFromTokens", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
                ?.Invoke(null, new object[] { tokens, 0 });
            Assert.NotNull(ast);
            Assert.Contains("FunctionNode", ast.ToString());
        }

        [Fact]
        public void CustomParserCanBeInjected()
        {
            var original = Nedev.FileConverters.XlsxToXls.Internal.FormulaCompiler.Parser;
            try
            {
                var called = false;
                Nedev.FileConverters.XlsxToXls.Internal.FormulaCompiler.Parser =
                    new TestParser(() => called = true);
                var bytes = Nedev.FileConverters.XlsxToXls.Internal.FormulaCompiler.Compile("=1+1", 0, new Dictionary<string, int>());
                Assert.True(called);
            }
            finally
            {
                Nedev.FileConverters.XlsxToXls.Internal.FormulaCompiler.Parser = original;
            }
        }

        private class TestParser : Nedev.FileConverters.XlsxToXls.Internal.FormulaCompiler.IFormulaParser
        {
            private readonly System.Action _onCall;
            public TestParser(System.Action onCall) => _onCall = onCall;
            public AstNode Parse(string formula, int sheetIndex, IReadOnlyDictionary<string, int> sheetIndexByName)
            {
                _onCall();
                return new Nedev.FileConverters.XlsxToXls.Internal.NumberNode(42);
            }
        }

        [Fact]
        public void DefaultParserIsSimpleFormulaParser()
        {
            // ensure default parser is the new simple one, not legacy
            var parserType = Nedev.FileConverters.XlsxToXls.Internal.FormulaCompiler.Parser.GetType();
            Assert.Contains("SimpleFormulaParser", parserType.Name);
        }

        [Theory]
        [InlineData("=3*(2+1)")]
        [InlineData("=SUM(1,2,3)")]
        [InlineData("=A1+B2")]
        [InlineData("=IF(TRUE,10,20)")]
        public void SimpleParserProducesBytes(string formula)
        {
            var bytes = Nedev.FileConverters.XlsxToXls.Internal.FormulaCompiler.Compile(formula, 0, new Dictionary<string, int>());
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
                new List<Internal.DataValidationInfo>(), new List<Internal.ConditionalFormatInfo>(), 0, new List<Internal.ChartData>(), null, new List<Internal.ShapeInfo>());
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