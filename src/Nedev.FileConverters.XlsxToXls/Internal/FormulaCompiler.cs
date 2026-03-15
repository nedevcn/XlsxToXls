using System.Buffers.Binary;
using System.Globalization;
using System.Text;

namespace Nedev.FileConverters.XlsxToXls.Internal;

// FormulaCompiler currently provides two layers:
// 1. a legacy tokenizer (Tokenize) that produces a flat token list from a formula string.
// 2. an AST builder and byte emitter that operate over the token sequence.
//
// Stage 2 refactoring aims to make the tokenizer pluggable and eventually
// replaceable by a third‑party parser; the AST types defined in FormulaAst.cs
// serve as the universal intermediate representation.
//
// Clients should call Compile(formula, ...) which uses the AST path.
// Helpers such as Tokenize are still used internally and exposed for tests but
// may be removed once a full parser is adopted.
internal static class FormulaCompiler
{
    private readonly record struct Tok(TokKind Kind, string Text, double Number, bool Bool, int Row, int Col, int Row2, int Col2, int SheetIndex, bool HasSheet);

    private enum TokKind
    {
        Number,
        String,
        Bool,
        Ref,
        Area,
        Func,
        Op,
        LParen,
        RParen,
        Comma
    }

    private enum OpKind
    {
        Add, Sub, Mul, Div, Pow, Concat,
        Eq, Ne, Lt, Le, Gt, Ge,
        UMinus
    }

    private static readonly Dictionary<string, ushort> FuncTable = new(StringComparer.OrdinalIgnoreCase)
    {
        // Logical / basic aggregation
        ["COUNT"] = 0x0000,
        ["IF"] = 0x0001,
        ["SUM"] = 0x0004,
        ["AVERAGE"] = 0x0005,
        ["MIN"] = 0x0006,
        ["MAX"] = 0x0007,
        ["COUNTA"] = 0x0008,
        ["PRODUCT"] = 0x000B,

        // Math
        ["ABS"] = 0x0018,
        ["INT"] = 0x0019,
        ["ROUND"] = 0x001B,
        ["MOD"] = 0x001D,
        ["POWER"] = 0x001E,
        ["SQRT"] = 0x0020,
        ["EXP"] = 0x0021,
        ["LN"] = 0x0022,
        ["LOG10"] = 0x0023,

        // Logical
        ["AND"] = 0x0024,
        ["OR"] = 0x0025,
        ["NOT"] = 0x0026,
        ["TRUE"] = 0x0027,
        ["FALSE"] = 0x0028,

        // Text
        ["LEFT"] = 0x0030,
        ["RIGHT"] = 0x0031,
        ["MID"] = 0x0032,
        ["FIND"] = 0x0033,
        ["SUBSTITUTE"] = 0x0034,
        ["TRIM"] = 0x0035,
        ["UPPER"] = 0x0036,
        ["LOWER"] = 0x0037,
        ["PROPER"] = 0x0038,
        ["VALUE"] = 0x0039,
        ["REPT"] = 0x001C,

        // Lookup & Reference
        ["VLOOKUP"] = 0x0066,
        ["HLOOKUP"] = 0x0067,
        ["INDIRECT"] = 0x0068,
        ["OFFSET"] = 0x0069,
        ["CHOOSE"] = 0x006A,
        ["ROW"] = 0x006B,
        ["COLUMN"] = 0x006C,
        ["ROWS"] = 0x006D,
        ["COLUMNS"] = 0x006E,

        // Statistical
        ["COUNTIF"] = 0x0100,
        ["SUMIF"] = 0x0101,
        ["AVERAGEIF"] = 0x0102,
        ["COUNTIFS"] = 0x0103,
        ["SUMIFS"] = 0x0104,
        ["AVERAGEIFS"] = 0x0105,
        ["MAXIFS"] = 0x0106,
        ["MINIFS"] = 0x0107,
        ["STDEV"] = 0x0108,
        ["STDEVP"] = 0x0109,
        ["VAR"] = 0x010A,
        ["VARP"] = 0x010B,
        ["MEDIAN"] = 0x010C,
        ["MODE"] = 0x010D,
        ["RANK"] = 0x010E,
        ["PERCENTILE"] = 0x010F,

        // Information
        ["ISBLANK"] = 0x0200,
        ["ISNUMBER"] = 0x0201,
        ["ISTEXT"] = 0x0202,
        ["ISLOGICAL"] = 0x0203,
        ["ISERROR"] = 0x0204,
        ["ISNA"] = 0x0205,
        ["INFO"] = 0x0206,
        ["CELL"] = 0x0207,

        // Date & Time
        ["DATE"] = 0x0300,
        ["TIME"] = 0x0301,
        ["DAY"] = 0x0302,
        ["MONTH"] = 0x0303,
        ["YEAR"] = 0x0304,
        ["WEEKDAY"] = 0x0305,
        ["HOUR"] = 0x0306,
        ["MINUTE"] = 0x0307,
        ["SECOND"] = 0x0308,
        ["DATEDIF"] = 0x0309,
        ["EDATE"] = 0x030A,
        ["EOMONTH"] = 0x030B,
        ["NETWORKDAYS"] = 0x030C,
        ["WORKDAY"] = 0x030D,
        ["TODAY"] = 0x030E,
        ["NOW"] = 0x030F,

        // Financial
        ["PMT"] = 0x0400,
        ["FV"] = 0x0401,
        ["PV"] = 0x0402,
        ["NPV"] = 0x0403,
        ["IRR"] = 0x0404,
        ["RATE"] = 0x0405,
        ["NPER"] = 0x0406,
        ["IPMT"] = 0x0407,
        ["PPMT"] = 0x0408,
        ["CUMIPMT"] = 0x0409,
        ["CUMPRINC"] = 0x040A,
        ["DB"] = 0x040B,
        ["DDB"] = 0x040C,
        ["SLN"] = 0x040D,
        ["SYD"] = 0x040E,
        ["VDB"] = 0x040F,

        // Database
        ["DSUM"] = 0x0500,
        ["DCOUNT"] = 0x0501,
        ["DCOUNTA"] = 0x0502,
        ["DAVERAGE"] = 0x0503,
        ["DMIN"] = 0x0504,
        ["DMAX"] = 0x0505,
        ["DSTDEV"] = 0x0506,
        ["DSTDEVP"] = 0x0507,
        ["DVAR"] = 0x0508,
        ["DVARP"] = 0x0509,
        ["DGET"] = 0x050A,

        // Engineering
        ["CONVERT"] = 0x0600,
        ["COMPLEX"] = 0x0601,
        ["IMREAL"] = 0x0602,
        ["IMAGINARY"] = 0x0603,
        ["IMABS"] = 0x0604,

        // Web
        ["ENCODEURL"] = 0x0700,
        ["WEBSERVICE"] = 0x0701,

        // Extended functions (Stage-1 additions)
        ["INDEX"] = 0x0800,
        ["MATCH"] = 0x0801,
        ["CONCAT"] = 0x0802,
        ["TEXT"] = 0x0803,
        ["LEN"] = 0x0804,
        ["IFERROR"] = 0x0805,
        ["IFNA"] = 0x0806,
        ["ANDIFS"] = 0x0807,
        ["ORIFS"] = 0x0808,
        ["SWITCH"] = 0x0809,
        ["IFS"] = 0x080A,
        ["MAXIFS"] = 0x080B,
        ["MINIFS"] = 0x080C,
        ["TEXTJOIN"] = 0x080D,
        ["UNICHAR"] = 0x080E,
        ["UNICODE"] = 0x080F
    };

    public static byte[] BuildExplicitListToken(string commaSeparated)
    {
        if (string.IsNullOrEmpty(commaSeparated)) return Array.Empty<byte>();
        var items = commaSeparated.Split(',');
        var sb = StringBuilderPool.Rent(256);
        try
        {
            for (var i = 0; i < items.Length; i++)
            {
                if (i > 0) sb.Append('\0');
                sb.Append(items[i].Trim());
            }
            var s = StringBuilderPool.ToStringAndReturn(sb);
            if (s.Length > 255) return Array.Empty<byte>();
            var need16 = HasHighChar(s.AsSpan());
            var list = ListPool<byte>.Rent(64);
            try
            {
                list.Add(0x17);
                list.Add((byte)(need16 ? 1 : 0));
                list.Add((byte)s.Length);
                if (need16)
                    foreach (var c in s)
                    {
                        list.Add((byte)(c & 0xFF));
                        list.Add((byte)((c >> 8) & 0xFF));
                    }
                else
                    foreach (var c in s)
                        list.Add((byte)c);
                return list.ToArray();
            }
            finally
            {
                ListPool<byte>.Return(list);
            }
        }
        catch
        {
            StringBuilderPool.Return(sb);
            throw;
        }
    }

    // parser abstraction (Stage2): callers can replace this with a richer parser.
    public static IFormulaParser Parser { get; set; } = new SimpleFormulaParser();

    // public compilation entry point now simply delegates to the configured parser.
    public static byte[] Compile(string formula, int currentSheetIndex, IReadOnlyDictionary<string, int> sheetIndexByName)
    {
        var ast = Parser.Parse(formula, currentSheetIndex, sheetIndexByName);
        var outBytes = new List<byte>(64);
        EmitFromAst(ast, outBytes);
        return outBytes.ToArray();
    }

    // legacy parser that wraps the existing tokenize + AST builder logic
    private class LegacyFormulaParser : IFormulaParser
    {
        public AstNode Parse(string formula, int sheetIndex, IReadOnlyDictionary<string, int> sheetIndexByName)
        {
            var tokens = Tokenize(formula, sheetIndex, sheetIndexByName);
            try
            {
                return AstFromTokens(tokens, sheetIndex);
            }
            finally
            {
                ListPool<Tok>.Return(tokens);
            }
        }
    }

    // simple hand-coded recursive-descent parser for formulas (Stage3 demonstration)
    internal class SimpleFormulaParser : IFormulaParser
    {
        private string _input = string.Empty;
        private int _pos;
        private int _sheetIndex;
        private IReadOnlyDictionary<string, int> _sheetIndexByName = new Dictionary<string,int>();

        public AstNode Parse(string formula, int sheetIndex, IReadOnlyDictionary<string, int> sheetIndexByName)
        {
            _sheetIndex = sheetIndex;
            _sheetIndexByName = sheetIndexByName;
            _input = formula.StartsWith("=", System.StringComparison.Ordinal) ? formula.Substring(1) : formula;
            _pos = 0;
            var expr = ParseExpression();
            SkipWhitespace();
            if (_pos < _input.Length)
                throw new FormatException($"Unexpected char '{_input[_pos]}' in formula");
            return expr;
        }

        private void SkipWhitespace()
        {
            while (_pos < _input.Length && char.IsWhiteSpace(_input[_pos])) _pos++;
        }

        private char Peek() => _pos < _input.Length ? _input[_pos] : '\0';
        private char Next() => _pos < _input.Length ? _input[_pos++] : '\0';

        private AstNode ParseExpression(int minPrec = 0)
        {
            var left = ParseUnary();
            while (true)
            {
                SkipWhitespace();
                var op = PeekOperator();
                if (op == null) break;
                var prec = GetPrecedence(op);
                if (prec < minPrec) break;
                bool rightAssoc = IsRightAssoc(op);
                // consume operator
                for (int i = 0; i < op.Length; i++) _pos++;
                var nextMin = prec + (rightAssoc ? 0 : 1);
                var right = ParseExpression(nextMin);
                left = new OperatorNode(op, left, right);
            }
            return left;
        }

        private AstNode ParseUnary()
        {
            SkipWhitespace();
            if (Peek() == '+' || Peek() == '-')
            {
                var op = Next().ToString();
                var operand = ParseUnary();
                return new UnaryOperatorNode(op, operand);
            }
            return ParsePrimary();
        }

        private AstNode ParsePrimary()
        {
            SkipWhitespace();
            var ch = Peek();
            if (ch == '(')
            {
                Next();
                var inner = ParseExpression();
                SkipWhitespace();
                if (Next() != ')') throw new FormatException("Missing closing parenthesis");
                return inner;
            }
            if (ch == '"')
            {
                return new StringNode(ParseString());
            }
            if (char.IsDigit(ch) || ch == '.')
            {
                return new NumberNode(ParseNumber());
            }
            if (char.IsLetter(ch) || ch == '_' || ch == '\'')
            {
                // could be boolean, function, or reference
                var id = ParseIdentifier();
                if (id.Equals("TRUE", System.StringComparison.OrdinalIgnoreCase))
                    return new BoolNode(true);
                if (id.Equals("FALSE", System.StringComparison.OrdinalIgnoreCase))
                    return new BoolNode(false);
                SkipWhitespace();
                if (Peek() == '(')
                {
                    // function call
                    Next();
                    var args = new List<AstNode>();
                    SkipWhitespace();
                    if (Peek() != ')')
                    {
                        while (true)
                        {
                            args.Add(ParseExpression());
                            SkipWhitespace();
                            if (Peek() == ',') { Next(); continue; }
                            break;
                        }
                    }
                    if (Next() != ')') throw new FormatException("Missing ')' after function args");
                    return new FunctionNode(id, args);
                }
                // otherwise treat as reference/area
                var refStr = id;
                // consume following reference characters
                while (char.IsLetterOrDigit(Peek()) || Peek() == '$' || Peek() == ':' || Peek() == '!' )
                {
                    refStr += Next();
                }
                if (TryParseA1OrArea(refStr, out int r1, out int c1, out int r2, out int c2, out bool isArea))
                {
                    if (isArea)
                        return new AreaNode(r1, c1, r2, c2, _sheetIndex, false);
                    else
                        return new RefNode(r1, c1, _sheetIndex, false);
                }
                // fallback to identifier as string
                return new StringNode(id);
            }
            throw new FormatException($"Unexpected character '{ch}'");
        }

        private string? PeekOperator()
        {
            var ops = new[] { ">=", "<=", "<>", ">", "<", "=", "+", "-", "*", "/", "^", "&" };
            foreach (var op in ops)
            {
                if (_input.Substring(_pos).StartsWith(op, System.StringComparison.Ordinal))
                    return op;
            }
            return null;
        }

        private int GetPrecedence(string op) => op switch
        {
            "^" => 4,
            "*" or "/" => 3,
            "+" or "-" => 2,
            "&" => 1,
            "=" or "<>" or ">" or "<" or ">=" or "<=" => 0,
            _ => -1
        };

        private bool IsRightAssoc(string op) => op == "^";

        private double ParseNumber()
        {
            var start = _pos;
            while (_pos < _input.Length && (char.IsDigit(_input[_pos]) || _input[_pos] == '.' || _input[_pos] == 'E' || _input[_pos] == 'e' || _input[_pos] == '+' || _input[_pos] == '-'))
            {
                if ((_input[_pos] == '+' || _input[_pos] == '-') && _pos > start && !(_input[_pos-1] == 'e' || _input[_pos-1] == 'E'))
                    break;
                _pos++;
            }
            var text = _input[start.._pos];
            if (!double.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                throw new FormatException($"Invalid number: {text}");
            return v;
        }

        private string ParseString()
        {
            // assumes leading '"'
            Next();
            var sb = new System.Text.StringBuilder();
            while (_pos < _input.Length)
            {
                if (Peek() == '"')
                {
                    Next();
                    if (Peek() == '"') { sb.Append('"'); Next(); continue; }
                    break;
                }
                sb.Append(Next());
            }
            return sb.ToString();
        }

        private string ParseIdentifier()
        {
            var start = _pos;
            if (Peek() == '\'')
            {
                // quoted sheet name or identifier
                Next();
                while (_pos < _input.Length && Peek() != '\'')
                {
                    if (Peek() == '\'' && _pos+1 < _input.Length && _input[_pos+1] == '\'')
                    {
                        _pos += 2;
                        continue;
                    }
                    _pos++;
                }
                if (Peek() == '\'') Next();
                // consume optional !
                if (Peek() == '!') { Next(); }
                return _input[start.._pos];
            }
            while (_pos < _input.Length && (char.IsLetterOrDigit(Peek()) || Peek() == '_')) _pos++;
            return _input[start.._pos];
        }
    }

    // parser interface definition
    public interface IFormulaParser
    {
        AstNode Parse(string formula, int sheetIndex, IReadOnlyDictionary<string, int> sheetIndexByName);
    }


    private static AstNode AstFromTokens(List<Tok> tokens, int currentSheet)
    {
        // shunting-yard variant that constructs AST nodes instead of raw tokens
        var output = new Stack<AstNode>();
        var ops = new Stack<Tok>();

        Tok? prev = null;
        foreach (var t in tokens)
        {
            switch (t.Kind)
            {
                case TokKind.Number:
                    output.Push(new NumberNode(t.Number));
                    break;
                case TokKind.String:
                    output.Push(new StringNode(t.Text));
                    break;
                case TokKind.Bool:
                    output.Push(new BoolNode(t.Bool));
                    break;
                case TokKind.Ref:
                    output.Push(new RefNode(t.Row, t.Col, t.SheetIndex, t.HasSheet));
                    break;
                case TokKind.Area:
                    output.Push(new AreaNode(t.Row, t.Col, t.Row2, t.Col2, t.SheetIndex, t.HasSheet));
                    break;
                case TokKind.Func:
                    ops.Push(t);
                    break;
                case TokKind.Op:
                {
                    while (ops.Count > 0 && ops.Peek().Kind == TokKind.Op)
                    {
                        var topOp = ParseOp(ops.Peek().Text, prev);
                        var thisOp = ParseOp(t.Text, prev);
                        if ((IsRightAssoc(thisOp) && Prec(thisOp) < Prec(topOp)) ||
                            (!IsRightAssoc(thisOp) && Prec(thisOp) <= Prec(topOp)))
                        {
                            var opTok = ops.Pop();
                            BuildOperatorNode(opTok, output);
                            continue;
                        }
                        break;
                    }
                    ops.Push(t);
                    break;
                }
                case TokKind.LParen:
                    ops.Push(t);
                    break;
                case TokKind.RParen:
                    while (ops.Count > 0 && ops.Peek().Kind != TokKind.LParen)
                    {
                        var opTok = ops.Pop();
                        BuildOperatorNode(opTok, output);
                    }
                    if (ops.Count == 0) throw new FormatException("Unbalanced parentheses");
                    ops.Pop(); // discard '('
                    if (ops.Count > 0 && ops.Peek().Kind == TokKind.Func)
                    {
                        var fnTok = ops.Pop();
                        // consume arguments from output until marker (not implemented yet)
                        // for now, assume single argument or none
                        var args = new List<AstNode>();
                        if (output.Count > 0) args.Add(output.Pop());
                        output.Push(new FunctionNode(fnTok.Text, args));
                    }
                    break;
            }
            prev = t;
        }
        while (ops.Count > 0)
        {
            var opTok = ops.Pop();
            if (opTok.Kind == TokKind.LParen || opTok.Kind == TokKind.RParen)
                throw new FormatException("Unbalanced parentheses");
            BuildOperatorNode(opTok, output);
        }
        return output.Count > 0 ? output.Pop() : new NumberNode(0);
    }

    private static void BuildOperatorNode(Tok opTok, Stack<AstNode> output)
    {
        var op = ParseOp(opTok.Text, null);
        if (op == OpKind.UMinus)
        {
            var operand = output.Pop();
            output.Push(new UnaryOperatorNode(opTok.Text, operand));
        }
        else
        {
            var right = output.Pop();
            var left = output.Pop();
            output.Push(new OperatorNode(opTok.Text, left, right));
        }
    }

    private static void EmitFromAst(AstNode node, List<byte> outBytes)
    {
        switch (node)
        {
            case NumberNode n:
                EmitNumber(outBytes, n.Value);
                break;
            case StringNode s:
                EmitString(outBytes, s.Text);
                break;
            case BoolNode b:
                outBytes.Add(0x1D);
                outBytes.Add((byte)(b.Value ? 1 : 0));
                break;
            case RefNode r:
                // reuse helper by creating temporary Tok
                var t = new Tok(TokKind.Ref, "", 0, false, r.Row, r.Col, 0, 0, r.SheetIndex, r.HasSheet);
                EmitRef(outBytes, t, r.SheetIndex);
                break;
            case AreaNode a:
                var ta = new Tok(TokKind.Area, "", 0, false, a.Row1, a.Col1, a.Row2, a.Col2, a.SheetIndex, a.HasSheet);
                EmitArea(outBytes, ta, a.SheetIndex);
                break;
            case OperatorNode op:
                EmitFromAst(op.Left, outBytes);
                EmitFromAst(op.Right, outBytes);
                EmitOp(outBytes, ParseOp(op.Op, null));
                break;
            case UnaryOperatorNode u:
                EmitFromAst(u.Operand, outBytes);
                EmitOp(outBytes, ParseOp(u.Op, null));
                break;
            case FunctionNode f:
                foreach (var arg in f.Arguments)
                    EmitFromAst(arg, outBytes);
                EmitFuncVar(outBytes, f.Name, (byte)f.Arguments.Count);
                break;
        }
    }

    private static void MarkFuncHasArg(Stack<bool> funcHadArg)
    {
        if (funcHadArg.Count == 0) return;
        if (funcHadArg.Peek()) return;
        funcHadArg.Pop();
        funcHadArg.Push(true);
    }

    private static void EmitFuncVar(List<byte> outBytes, string name, byte cargs)
    {
        if (!FuncTable.TryGetValue(name, out var ftab))
            throw new NotSupportedException($"Unsupported function: {name}");
        outBytes.Add(0x22);
        outBytes.Add(cargs);
        outBytes.Add((byte)(ftab & 0xFF));
        outBytes.Add((byte)((ftab >> 8) & 0xFF));
    }

    private static void EmitOp(List<byte> outBytes, OpKind op)
    {
        outBytes.Add(op switch
        {
            OpKind.Add => (byte)0x03,
            OpKind.Sub => (byte)0x04,
            OpKind.Mul => (byte)0x05,
            OpKind.Div => (byte)0x06,
            OpKind.Pow => (byte)0x07,
            OpKind.Concat => (byte)0x08,
            OpKind.Lt => (byte)0x09,
            OpKind.Le => (byte)0x0A,
            OpKind.Eq => (byte)0x0B,
            OpKind.Ge => (byte)0x0C,
            OpKind.Gt => (byte)0x0D,
            OpKind.Ne => (byte)0x0E,
            OpKind.UMinus => (byte)0x13,
            _ => throw new NotSupportedException()
        });
    }

    private static int Prec(OpKind op) => op switch
    {
        OpKind.UMinus => 5,
        OpKind.Pow => 4,
        OpKind.Mul or OpKind.Div => 3,
        OpKind.Add or OpKind.Sub => 2,
        OpKind.Concat => 1,
        _ => 0
    };

    private static bool IsRightAssoc(OpKind op) => op is OpKind.Pow or OpKind.UMinus;

    private static OpKind ParseOp(string op, Tok? prev)
    {
        if (op == "-" && (prev == null || prev.Value.Kind is TokKind.Op or TokKind.LParen or TokKind.Comma))
            return OpKind.UMinus;
        return op switch
        {
            "+" => OpKind.Add,
            "-" => OpKind.Sub,
            "*" => OpKind.Mul,
            "/" => OpKind.Div,
            "^" => OpKind.Pow,
            "&" => OpKind.Concat,
            "=" => OpKind.Eq,
            "<>" => OpKind.Ne,
            "<" => OpKind.Lt,
            "<=" => OpKind.Le,
            ">" => OpKind.Gt,
            ">=" => OpKind.Ge,
            _ => throw new NotSupportedException($"Unsupported operator: {op}")
        };
    }

    private static void EmitNumber(List<byte> outBytes, double n)
    {
        if (n >= 0 && n <= 65535 && Math.Abs(n - Math.Round(n)) < 0.0000001)
        {
            outBytes.Add(0x1E);
            var v = (ushort)Math.Round(n);
            outBytes.Add((byte)(v & 0xFF));
            outBytes.Add((byte)((v >> 8) & 0xFF));
            return;
        }
        outBytes.Add(0x1F);
        Span<byte> tmp = stackalloc byte[8];
        BufferHelpers.WriteDoubleLittleEndian(tmp, n);
        for (var i = 0; i < 8; i++) outBytes.Add(tmp[i]);
    }

    private static void EmitString(List<byte> outBytes, string s)
    {
        if (s.Length > 255) s = s[..255];
        var need16 = false;
        foreach (var c in s) { if (c > 255) { need16 = true; break; } }
        outBytes.Add(0x17);
        outBytes.Add((byte)(need16 ? 1 : 0));
        outBytes.Add((byte)s.Length);
        if (need16)
        {
            foreach (var c in s)
            {
                outBytes.Add((byte)(c & 0xFF));
                outBytes.Add((byte)((c >> 8) & 0xFF));
            }
        }
        else
        {
            foreach (var c in s) outBytes.Add((byte)c);
        }
    }

    private static void EmitRef(List<byte> outBytes, Tok t, int currentSheetIndex)
    {
        if (t.HasSheet)
        {
            outBytes.Add(0x3A);
            outBytes.Add((byte)(t.SheetIndex & 0xFF));
            outBytes.Add((byte)((t.SheetIndex >> 8) & 0xFF));
        }
        else
        {
            outBytes.Add(0x24);
        }
        outBytes.Add((byte)(t.Row & 0xFF));
        outBytes.Add((byte)((t.Row >> 8) & 0xFF));
        outBytes.Add((byte)(t.Col & 0xFF));
        outBytes.Add((byte)((t.Col >> 8) & 0xFF));
    }

    private static void EmitArea(List<byte> outBytes, Tok t, int currentSheetIndex)
    {
        if (t.HasSheet)
        {
            outBytes.Add(0x3B);
            outBytes.Add((byte)(t.SheetIndex & 0xFF));
            outBytes.Add((byte)((t.SheetIndex >> 8) & 0xFF));
        }
        else
        {
            outBytes.Add(0x25);
        }
        outBytes.Add((byte)(t.Row & 0xFF));
        outBytes.Add((byte)((t.Row >> 8) & 0xFF));
        outBytes.Add((byte)(t.Row2 & 0xFF));
        outBytes.Add((byte)((t.Row2 >> 8) & 0xFF));
        outBytes.Add((byte)(t.Col & 0xFF));
        outBytes.Add((byte)((t.Col >> 8) & 0xFF));
        outBytes.Add((byte)(t.Col2 & 0xFF));
        outBytes.Add((byte)((t.Col2 >> 8) & 0xFF));
    }

    private static List<Tok> Tokenize(string formula, int currentSheetIndex, IReadOnlyDictionary<string, int> sheetIndexByName)
    {
        if (formula.StartsWith("=", StringComparison.Ordinal)) formula = formula[1..];
        var list = ListPool<Tok>.Rent(32);
        var i = 0;
        while (i < formula.Length)
        {
            var ch = formula[i];
            if (char.IsWhiteSpace(ch)) { i++; continue; }
            if (ch == '"')
            {
                i++;
                var start = i;
                while (i < formula.Length && formula[i] != '"') i++;
                var s = formula[start..Math.Min(i, formula.Length)];
                if (i < formula.Length && formula[i] == '"') i++;
                list.Add(new Tok(TokKind.String, s, 0, false, 0, 0, 0, 0, 0, false));
                continue;
            }
            if (ch == '(') { list.Add(new Tok(TokKind.LParen, "(", 0, false, 0, 0, 0, 0, 0, false)); i++; continue; }
            if (ch == ')') { list.Add(new Tok(TokKind.RParen, ")", 0, false, 0, 0, 0, 0, 0, false)); i++; continue; }
            if (ch == ',') { list.Add(new Tok(TokKind.Comma, ",", 0, false, 0, 0, 0, 0, 0, false)); i++; continue; }

            if (ch is '+' or '-' or '*' or '/' or '^' or '&')
            {
                list.Add(new Tok(TokKind.Op, ch.ToString(), 0, false, 0, 0, 0, 0, 0, false));
                i++;
                continue;
            }
            if (ch is '<' or '>' or '=')
            {
                var op = ch.ToString();
                if (i + 1 < formula.Length)
                {
                    var ch2 = formula[i + 1];
                    if ((ch == '<' && (ch2 == '=' || ch2 == '>')) || (ch == '>' && ch2 == '='))
                    {
                        op = new string(new[] { ch, ch2 });
                        i += 2;
                        list.Add(new Tok(TokKind.Op, op, 0, false, 0, 0, 0, 0, 0, false));
                        continue;
                    }
                }
                i++;
                list.Add(new Tok(TokKind.Op, op, 0, false, 0, 0, 0, 0, 0, false));
                continue;
            }

            if (char.IsDigit(ch) || ch == '.')
            {
                var start = i;
                i++;
                while (i < formula.Length && (char.IsDigit(formula[i]) || formula[i] == '.' || formula[i] == 'E' || formula[i] == 'e' || formula[i] == '+' || formula[i] == '-'))
                {
                    if ((formula[i] == '+' || formula[i] == '-') && !(formula[i - 1] is 'e' or 'E')) break;
                    i++;
                }
                var numText = formula[start..i];
                if (!double.TryParse(numText, NumberStyles.Any, CultureInfo.InvariantCulture, out var num))
                    throw new FormatException($"Invalid number: {numText}");
                list.Add(new Tok(TokKind.Number, numText, num, false, 0, 0, 0, 0, 0, false));
                continue;
            }

            if (ch == '\'' || char.IsLetter(ch) || ch == '$')
            {
                var start = i;
                var sheetName = (string?)null;
                var hasSheet = false;
                var sheetIdx = currentSheetIndex;

                var save = i;
                if (ch == '\'')
                {
                    i++;
                    var sb = new StringBuilder();
                    while (i < formula.Length)
                    {
                        if (formula[i] == '\'' && i + 1 < formula.Length && formula[i + 1] == '\'')
                        {
                            sb.Append('\'');
                            i += 2;
                            continue;
                        }
                        if (formula[i] == '\'') { i++; break; }
                        sb.Append(formula[i++]);
                    }
                    if (i < formula.Length && formula[i] == '!')
                    {
                        hasSheet = true;
                        sheetName = sb.ToString();
                        i++;
                    }
                    else
                    {
                        i = save;
                    }
                }
                else
                {
                    var j = i;
                    while (j < formula.Length && (char.IsLetterOrDigit(formula[j]) || formula[j] == '_' || formula[j] == '.')) j++;
                    if (j < formula.Length && formula[j] == '!')
                    {
                        hasSheet = true;
                        sheetName = formula[i..j];
                        i = j + 1;
                    }
                }

                if (hasSheet && sheetName != null && sheetIndexByName.TryGetValue(sheetName, out var idxSheet))
                    sheetIdx = idxSheet;

                // cell / area
                var refStart = i;
                while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] == '$' || formula[i] == ':')) i++;
                var refText = formula[refStart..i];
                if (TryParseA1OrArea(refText, out var r1, out var c1, out var r2, out var c2, out var isArea))
                {
                    if (isArea)
                        list.Add(new Tok(TokKind.Area, "", 0, false, r1, c1, r2, c2, sheetIdx, hasSheet));
                    else
                        list.Add(new Tok(TokKind.Ref, "", 0, false, r1, c1, 0, 0, sheetIdx, hasSheet));
                    continue;
                }

                // identifier (function / bool)
                i = start;
                var idStart = i;
                while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] == '_' || formula[i] == '.')) i++;
                var ident = formula[idStart..i];
                if (ident.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
                {
                    list.Add(new Tok(TokKind.Bool, ident, 0, true, 0, 0, 0, 0, 0, false));
                    continue;
                }
                if (ident.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
                {
                    list.Add(new Tok(TokKind.Bool, ident, 0, false, 0, 0, 0, 0, 0, false));
                    continue;
                }
                // function if followed by '('
                var k = i;
                while (k < formula.Length && char.IsWhiteSpace(formula[k])) k++;
                if (k < formula.Length && formula[k] == '(')
                {
                    list.Add(new Tok(TokKind.Func, ident, 0, false, 0, 0, 0, 0, 0, false));
                    continue;
                }
                throw new NotSupportedException($"Unsupported identifier: {ident}");
            }

            throw new NotSupportedException($"Unexpected char '{ch}' in formula");
        }
        return list;
    }

    private static bool TryParseA1OrArea(string s, out int r1, out int c1, out int r2, out int c2, out bool isArea)
    {
        r1 = c1 = r2 = c2 = 0;
        isArea = false;
        s = s.Replace("$", "", StringComparison.Ordinal);
        var colon = s.IndexOf(':');
        if (colon >= 0)
        {
            var a = s[..colon];
            var b = s[(colon + 1)..];
            if (TryParseCell(a, out r1, out c1) && TryParseCell(b, out r2, out c2))
            {
                isArea = true;
                return true;
            }
            if (int.TryParse(a, out var ra) && int.TryParse(b, out var rb))
            {
                r1 = Math.Max(0, ra - 1);
                r2 = Math.Max(0, rb - 1);
                c1 = 0;
                c2 = 255;
                isArea = true;
                return true;
            }
            if (TryParseColOnly(a, out var ca) && TryParseColOnly(b, out var cb))
            {
                c1 = ca;
                c2 = cb;
                r1 = 0;
                r2 = 65535;
                isArea = true;
                return true;
            }
            return false;
        }
        return TryParseCell(s, out r1, out c1);
    }

    private static bool TryParseCell(string s, out int row, out int col)
    {
        row = col = 0;
        if (string.IsNullOrEmpty(s)) return false;
        var i = 0;
        while (i < s.Length && char.IsLetter(s[i])) i++;
        if (i == 0 || i >= s.Length) return false;
        if (!int.TryParse(s.AsSpan(i), out var r)) return false;
        col = ParseCol(s.AsSpan(0, i));
        if (col < 0) return false;
        row = r - 1;
        return row >= 0 && row <= 65535 && col >= 0 && col <= 255;
    }

    private static bool TryParseColOnly(string s, out int col)
    {
        col = -1;
        if (string.IsNullOrEmpty(s)) return false;
        for (var i = 0; i < s.Length; i++) if (!char.IsLetter(s[i])) return false;
        col = ParseCol(s.AsSpan());
        return col >= 0 && col <= 255;
    }

    private static int ParseCol(ReadOnlySpan<char> s)
    {
        var col = 0;
        foreach (var c in s)
            col = col * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        return col - 1;
    }

    private static bool HasHighChar(ReadOnlySpan<char> s)
    {
        foreach (var c in s) if (c > 255) return true;
        return false;
    }
}

