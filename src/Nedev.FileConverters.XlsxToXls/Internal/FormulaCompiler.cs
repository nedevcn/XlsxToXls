using System.Buffers.Binary;
using System.Globalization;
using System.Text;

namespace Nedev.FileConverters.XlsxToXls.Internal;

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

        // Math
        ["ABS"] = 0x0018,
        ["INT"] = 0x0019,
        ["ROUND"] = 0x001B,

        // Logical
        ["AND"] = 0x0024,
        ["OR"] = 0x0025,
        ["NOT"] = 0x0026,

        // Lookup
            ["VLOOKUP"] = 0x0066,

            // Extended/Stage‑1 additions (codes chosen to avoid collision; may be refined later)
            ["SUMIF"] = 0x0100,
            ["COUNTIF"] = 0x0101,
            ["INDEX"] = 0x0102,
            ["MATCH"] = 0x0103,
            ["CONCAT"] = 0x0104,
            ["TEXT"] = 0x0105,
            ["LEN"] = 0x0106,
            ["TODAY"] = 0x0107,
            ["NOW"] = 0x0108,
            ["IFERROR"] = 0x0109
        };

    public static byte[] BuildExplicitListToken(string commaSeparated)
    {
        if (string.IsNullOrEmpty(commaSeparated)) return Array.Empty<byte>();
        var items = commaSeparated.Split(',');
        var sb = new StringBuilder();
        for (var i = 0; i < items.Length; i++)
        {
            if (i > 0) sb.Append('\0');
            sb.Append(items[i].Trim());
        }
        var s = sb.ToString();
        if (s.Length > 255) return Array.Empty<byte>();
        var need16 = HasHighChar(s.AsSpan());
        var list = new List<byte> { 0x17, (byte)(need16 ? 1 : 0), (byte)s.Length };
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

    public static byte[] Compile(string formula, int currentSheetIndex, IReadOnlyDictionary<string, int> sheetIndexByName)
    {
        var tokens = Tokenize(formula, currentSheetIndex, sheetIndexByName);
        var outBytes = new List<byte>(64);
        var ops = new Stack<(TokKind Kind, OpKind Op, string FuncName)>();
        var funcArgCounts = new Stack<int>();
        var funcHadArg = new Stack<bool>();

        Tok? prev = null;
        foreach (var t in tokens)
        {
            switch (t.Kind)
            {
                case TokKind.Number:
                    EmitNumber(outBytes, t.Number);
                    MarkFuncHasArg(funcHadArg);
                    break;
                case TokKind.String:
                    EmitString(outBytes, t.Text);
                    MarkFuncHasArg(funcHadArg);
                    break;
                case TokKind.Bool:
                    outBytes.Add(0x1D);
                    outBytes.Add((byte)(t.Bool ? 1 : 0));
                    MarkFuncHasArg(funcHadArg);
                    break;
                case TokKind.Ref:
                    EmitRef(outBytes, t, currentSheetIndex);
                    MarkFuncHasArg(funcHadArg);
                    break;
                case TokKind.Area:
                    EmitArea(outBytes, t, currentSheetIndex);
                    MarkFuncHasArg(funcHadArg);
                    break;
                case TokKind.Func:
                    ops.Push((TokKind.Func, default, t.Text));
                    break;
                case TokKind.Op:
                {
                    var op = ParseOp(t.Text, prev);
                    while (ops.Count > 0 && ops.Peek().Kind == TokKind.Op)
                    {
                        var top = ops.Peek().Op;
                        if ((IsRightAssoc(op) && Prec(op) < Prec(top)) || (!IsRightAssoc(op) && Prec(op) <= Prec(top)))
                        {
                            EmitOp(outBytes, ops.Pop().Op);
                            continue;
                        }
                        break;
                    }
                    ops.Push((TokKind.Op, op, ""));
                    break;
                }
                case TokKind.LParen:
                    ops.Push((TokKind.LParen, default, ""));
                    if (prev is { Kind: TokKind.Func })
                    {
                        funcArgCounts.Push(0);
                        funcHadArg.Push(false);
                    }
                    break;
                case TokKind.Comma:
                    while (ops.Count > 0 && ops.Peek().Kind != TokKind.LParen)
                    {
                        if (ops.Peek().Kind == TokKind.Op) EmitOp(outBytes, ops.Pop().Op);
                        else break;
                    }
                    if (funcArgCounts.Count > 0)
                    {
                        funcArgCounts.Push(funcArgCounts.Pop() + 1);
                        if (funcHadArg.Count > 0) { funcHadArg.Pop(); funcHadArg.Push(true); }
                    }
                    break;
                case TokKind.RParen:
                    while (ops.Count > 0 && ops.Peek().Kind != TokKind.LParen)
                    {
                        if (ops.Peek().Kind == TokKind.Op) EmitOp(outBytes, ops.Pop().Op);
                        else break;
                    }
                    if (ops.Count == 0) throw new FormatException("Unbalanced parentheses");
                    ops.Pop();
                    if (ops.Count > 0 && ops.Peek().Kind == TokKind.Func)
                    {
                        var fn = ops.Pop().FuncName;
                        var argCount = funcArgCounts.Count > 0 ? funcArgCounts.Pop() : 0;
                        var hadArg = funcHadArg.Count > 0 && funcHadArg.Pop();
                        var cargs = hadArg ? (byte)(argCount + 1) : (byte)0;
                        EmitFuncVar(outBytes, fn, cargs);
                    }
                    break;
            }
            prev = t;
        }

        while (ops.Count > 0)
        {
            var top = ops.Pop();
            if (top.Kind is TokKind.LParen or TokKind.RParen) throw new FormatException("Unbalanced parentheses");
            if (top.Kind == TokKind.Op) EmitOp(outBytes, top.Op);
        }

        return outBytes.ToArray();
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
        var list = new List<Tok>(32);
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

