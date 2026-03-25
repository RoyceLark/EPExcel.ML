using System.Text;
using System.Text.RegularExpressions;

namespace EPExcel.ML.Formulas;

public enum ShiftType { Row, Column }

/// <summary>
/// Recursive-descent formula evaluator — EPExcel 8 parity.
/// Handles arithmetic, comparisons, string concat, function calls,
/// cell/range references (A1, Sheet1!A1:B2), named ranges, array constants.
/// All 463 EPExcel functions via FunctionLibrary.
/// </summary>
public sealed partial class FormulaEngine
{
    private readonly ExcelWorkbook? _workbook;
    private readonly Dictionary<string, Func<object?[], ExcelWorksheet, object?>> _functions;

    public enum PrecisionStrategy { Excel, DotNet }
    public PrecisionStrategy RoundingStrategy { get; set; } = PrecisionStrategy.Excel;

    public FormulaEngine(ExcelWorkbook? workbook)
    {
        _workbook = workbook;
        _functions = FunctionLibrary.Build();
    }

    // ── Public API ────────────────────────────────────────────────────────────

    public object? Evaluate(string formula, ExcelWorksheet ws)
    {
        if (string.IsNullOrWhiteSpace(formula)) return null;
        try
        {
            var f = formula.TrimStart('=').Trim();
            var tokens = Tokenize(f);
            var (val, _) = ParseExpr(tokens, ws, 0);
            var result = val is RangeRef rr ? rr.Worksheet.GetCell(rr.FromRow, rr.FromCol)?.DisplayValue : val;
            if (RoundingStrategy == PrecisionStrategy.Excel && result is double d)
                return FunctionLibrary.Round15(d);
            return result;
        }
        catch (FormulaException ex) { return new CellError(ex.Code); }
        catch { return new CellError(ExcelErrorCode.Value); }
    }

    public void RegisterFunction(string name, Func<object?[], ExcelWorksheet, object?> fn)
        => _functions[name] = fn;

    // ── Live range reference ──────────────────────────────────────────────────

    public sealed class RangeRef
    {
        public ExcelWorksheet Worksheet { get; }
        public int FromRow { get; }
        public int FromCol { get; }
        public int ToRow { get; }
        public int ToCol { get; }
        public int RowCount => ToRow - FromRow + 1;
        public int ColCount => ToCol - FromCol + 1;

        public RangeRef(ExcelWorksheet ws, int fr, int fc, int tr, int tc)
        { Worksheet = ws; FromRow = fr; FromCol = fc; ToRow = tr; ToCol = tc; }

        public IEnumerable<object?> Values()
        {
            for (int r = FromRow; r <= ToRow; r++)
                for (int c = FromCol; c <= ToCol; c++)
                    yield return Worksheet.GetCell(r, c)?.DisplayValue;
        }
    }

    public static IEnumerable<object?> ResolveValues(object? arg) => arg switch
    {
        RangeRef rr => rr.Values(),
        object?[] arr => arr.SelectMany(ResolveValues),
        _ => new[] { arg }
    };

    public static object? Flatten(object? v) => v switch
    {
        RangeRef rr => rr.Worksheet.GetCell(rr.FromRow, rr.FromCol)?.DisplayValue,
        object?[] arr when arr.Length > 0 => arr[0],
        _ => v
    };

    // ── Formula address shifting ──────────────────────────────────────────────

    public static string ShiftFormula(string formula, ShiftType type, int position, int count)
    {
        return CellRefRegex().Replace(formula, m =>
        {
            bool ca = m.Groups["ca"].Value == "$";
            bool ra = m.Groups["ra"].Value == "$";
            int col = ExcelAddressParser.ColumnLetterToNumber(m.Groups["col"].Value);
            int row = int.Parse(m.Groups["row"].Value);

            if (type == ShiftType.Row && !ra && row >= position)
            {
                row += count;
                if (row < 1) return "#REF!";
            }
            else if (type == ShiftType.Column && !ca && col >= position)
            {
                col += count;
                if (col < 1) return "#REF!";
            }
            else return m.Value;

            return $"{(ca ? "$" : "")}{ExcelAddressParser.ColumnNumberToLetter(col)}{(ra ? "$" : "")}{row}";
        });
    }

    // ── Parser ────────────────────────────────────────────────────────────────

    private enum TT { Num, Str, Bool, Err, Name, Op, LParen, RParen, Comma, LBrace, RBrace }
    private sealed record Tok(TT Type, string Val);

    private (object? v, int p) ParseExpr(List<Tok> t, ExcelWorksheet ws, int p)
        => ParseConcat(t, ws, p);

    private (object? v, int p) ParseConcat(List<Tok> t, ExcelWorksheet ws, int p)
    {
        var (l, p2) = ParseCompare(t, ws, p);
        while (p2 < t.Count && t[p2].Val == "&")
        {
            var (r, p3) = ParseCompare(t, ws, p2 + 1);
            l = (Flatten(l)?.ToString() ?? "") + (Flatten(r)?.ToString() ?? "");
            p2 = p3;
        }
        return (l, p2);
    }

    private (object? v, int p) ParseCompare(List<Tok> t, ExcelWorksheet ws, int p)
    {
        var (l, p2) = ParseAddSub(t, ws, p);
        while (p2 < t.Count && t[p2].Type == TT.Op &&
               t[p2].Val is "=" or "<>" or "<" or ">" or "<=" or ">=")
        {
            var op = t[p2].Val;
            var (r, p3) = ParseAddSub(t, ws, p2 + 1);
            l = Compare(Flatten(l), Flatten(r), op);
            p2 = p3;
        }
        return (l, p2);
    }

    private (object? v, int p) ParseAddSub(List<Tok> t, ExcelWorksheet ws, int p)
    {
        var (l, p2) = ParseMulDiv(t, ws, p);
        while (p2 < t.Count && t[p2].Type == TT.Op && t[p2].Val is "+" or "-")
        {
            var op = t[p2].Val;
            var (r, p3) = ParseMulDiv(t, ws, p2 + 1);
            double lv = FunctionLibrary.Num(Flatten(l)), rv = FunctionLibrary.Num(Flatten(r));
            l = (object?)(op == "+" ? lv + rv : lv - rv);
            p2 = p3;
        }
        return (l, p2);
    }

    private (object? v, int p) ParseMulDiv(List<Tok> t, ExcelWorksheet ws, int p)
    {
        var (l, p2) = ParsePow(t, ws, p);
        while (p2 < t.Count && t[p2].Type == TT.Op && t[p2].Val is "*" or "/")
        {
            var op = t[p2].Val;
            var (r, p3) = ParsePow(t, ws, p2 + 1);
            double lv = FunctionLibrary.Num(Flatten(l)), rv = FunctionLibrary.Num(Flatten(r));
            l = op == "/" ? (rv == 0 ? (object?)new CellError(ExcelErrorCode.Div0) : lv / rv) : lv * rv;
            p2 = p3;
        }
        return (l, p2);
    }

    private (object? v, int p) ParsePow(List<Tok> t, ExcelWorksheet ws, int p)
    {
        var (l, p2) = ParseUnary(t, ws, p);
        if (p2 < t.Count && t[p2].Val == "^")
        {
            var (r, p3) = ParseUnary(t, ws, p2 + 1);
            l = Math.Pow(FunctionLibrary.Num(Flatten(l)), FunctionLibrary.Num(Flatten(r)));
            p2 = p3;
        }
        return (l, p2);
    }

    private (object? v, int p) ParseUnary(List<Tok> t, ExcelWorksheet ws, int p)
    {
        if (p < t.Count && t[p].Val == "-")
        {
            var (v, p2) = ParseUnary(t, ws, p + 1);
            return ((object?)(-FunctionLibrary.Num(Flatten(v))), p2);
        }
        if (p < t.Count && t[p].Val == "+") return ParseUnary(t, ws, p + 1);
        return ParsePercent(t, ws, p);
    }

    private (object? v, int p) ParsePercent(List<Tok> t, ExcelWorksheet ws, int p)
    {
        var (v, p2) = ParsePrimary(t, ws, p);
        if (p2 < t.Count && t[p2].Val == "%")
            return ((object?)(FunctionLibrary.Num(Flatten(v)) / 100.0), p2 + 1);
        return (v, p2);
    }

    private (object? v, int p) ParsePrimary(List<Tok> t, ExcelWorksheet ws, int p)
    {
        if (p >= t.Count) return (null, p);
        var tok = t[p];

        if (tok.Type == TT.LParen)
        {
            var (v, p2) = ParseExpr(t, ws, p + 1);
            return (v, p2 < t.Count && t[p2].Type == TT.RParen ? p2 + 1 : p2);
        }

        if (tok.Type == TT.LBrace)
        {
            var arr = new List<object?>();
            int p2 = p + 1;
            while (p2 < t.Count && t[p2].Type != TT.RBrace)
            {
                if (t[p2].Type is TT.Comma) { p2++; continue; }
                var (v, p3) = ParsePrimary(t, ws, p2);
                arr.Add(Flatten(v));
                p2 = p3;
            }
            return (arr.ToArray(), p2 < t.Count ? p2 + 1 : p2);
        }

        if (tok.Type == TT.Num)
            return ((object?)double.Parse(tok.Val, System.Globalization.CultureInfo.InvariantCulture), p + 1);
        if (tok.Type == TT.Str) return ((object?)tok.Val, p + 1);
        if (tok.Type == TT.Bool) return ((object?)(tok.Val == "TRUE"), p + 1);
        if (tok.Type == TT.Err) return ((object?)new CellError(ParseErrCode(tok.Val)), p + 1);

        if (tok.Type == TT.Name)
        {
            var name = tok.Val;
            if (p + 1 < t.Count && t[p + 1].Type == TT.LParen)
                return CallFunc(name, t, ws, p + 1);
            return (ResolveRef(name, ws), p + 1);
        }

        return (null, p + 1);
    }

    private (object? v, int p) CallFunc(string name, List<Tok> t, ExcelWorksheet ws, int parenPos)
    {
        var args = new List<object?>();
        int p = parenPos + 1;

        if (p < t.Count && t[p].Type == TT.RParen)
            return (Invoke(name, Array.Empty<object?>(), ws), p + 1);

        while (p < t.Count && t[p].Type != TT.RParen)
        {
            if (t[p].Type == TT.Comma) { args.Add(null); p++; continue; }
            var (arg, p2) = ParseExpr(t, ws, p);
            args.Add(arg);
            p = p2;
            if (p < t.Count && t[p].Type == TT.Comma) p++;
        }

        return (Invoke(name, args.ToArray(), ws), p < t.Count ? p + 1 : p);
    }

    private object? Invoke(string name, object?[] args, ExcelWorksheet ws)
    {
        if (_workbook?.Lambdas.TryGet(name) is { } lam) return lam(args, ws);
        if (_functions.TryGetValue(name, out var fn)) return fn(args, ws);
        return new CellError(ExcelErrorCode.Name);
    }

    private object? ResolveRef(string token, ExcelWorksheet ws)
    {
        if (token.Contains('!'))
        {
            int bang = token.LastIndexOf('!');
            var sheetName = token[..bang].Trim('\'', '"');
            var addr = token[(bang + 1)..];
            var targetWs = _workbook?.GetWorksheet(sheetName) ?? ws;
            return ResolveAddr(addr, targetWs);
        }
        if (_workbook?.NamedRanges.TryGetValue(token, out var nr) == true)
        {
            var rng = nr.Range;
            if (rng.IsSingleCell)
                return nr.Worksheet.GetCell(rng.FromRow, rng.FromCol)?.DisplayValue;
            return new RangeRef(nr.Worksheet, rng.FromRow, rng.FromCol, rng.ToRow, rng.ToCol);
        }
        return ResolveAddr(token, ws);
    }

    private object? ResolveAddr(string addr, ExcelWorksheet ws)
    {
        try
        {
            if (addr.Contains(':'))
            {
                var (fr, fc, tr, tc) = ExcelAddressParser.ParseRange(addr);
                if (fr == tr && fc == tc)
                    return ws.GetCell(fr, fc)?.DisplayValue;
                return new RangeRef(ws, fr, fc, tr, tc);
            }
            var (row, col) = ExcelAddressParser.ParseCell(addr);
            return ws.GetCell(row, col)?.DisplayValue;
        }
        catch { return new CellError(ExcelErrorCode.Ref); }
    }

    private static object? Compare(object? l, object? r, string op)
    {
        if (l is double ld && r is double rd)
            return op switch
            {
                "=" => ld == rd, "<>" => ld != rd, "<" => ld < rd,
                ">" => ld > rd, "<=" => ld <= rd, ">=" => ld >= rd, _ => (object?)false
            };
        int cmp = string.Compare(l?.ToString() ?? "", r?.ToString() ?? "", StringComparison.OrdinalIgnoreCase);
        return op switch
        {
            "=" => cmp == 0, "<>" => cmp != 0, "<" => cmp < 0,
            ">" => cmp > 0, "<=" => cmp <= 0, ">=" => cmp >= 0, _ => (object?)false
        };
    }

    private static ExcelErrorCode ParseErrCode(string s) => s.ToUpperInvariant() switch
    {
        "#DIV/0!" => ExcelErrorCode.Div0, "#VALUE!" => ExcelErrorCode.Value,
        "#REF!"   => ExcelErrorCode.Ref,  "#NAME?"  => ExcelErrorCode.Name,
        "#NUM!"   => ExcelErrorCode.Num,  "#N/A"    => ExcelErrorCode.NA,
        "#NULL!"  => ExcelErrorCode.Null, "#SPILL!" => ExcelErrorCode.Spill,
        _ => ExcelErrorCode.Value
    };

    // ── Tokenizer ─────────────────────────────────────────────────────────────

    private static List<Tok> Tokenize(string expr)
    {
        var tokens = new List<Tok>();
        int i = 0;
        while (i < expr.Length)
        {
            char c = expr[i];
            if (char.IsWhiteSpace(c)) { i++; continue; }

            if (c == '"')
            {
                var sb = new StringBuilder(); i++;
                while (i < expr.Length)
                {
                    if (expr[i] == '"')
                    {
                        i++;
                        if (i < expr.Length && expr[i] == '"') { sb.Append('"'); i++; }
                        else break;
                    }
                    else sb.Append(expr[i++]);
                }
                tokens.Add(new Tok(TT.Str, sb.ToString())); continue;
            }

            if (c == '#')
            {
                int j = i;
                while (j < expr.Length && !char.IsWhiteSpace(expr[j]) && expr[j] != ')' && expr[j] != ',') j++;
                tokens.Add(new Tok(TT.Err, expr[i..j])); i = j; continue;
            }

            if (char.IsDigit(c) || (c == '.' && i + 1 < expr.Length && char.IsDigit(expr[i + 1])))
            {
                int j = i;
                while (j < expr.Length && (char.IsDigit(expr[j]) || expr[j] == '.' ||
                    expr[j] is 'E' or 'e' ||
                    ((expr[j] is '+' or '-') && j > 0 && (expr[j - 1] is 'E' or 'e'))))
                    j++;
                tokens.Add(new Tok(TT.Num, expr[i..j])); i = j; continue;
            }

            if (char.IsLetter(c) || c is '_' or '\'')
            {
                int j = i;
                if (c == '\'') { j++; while (j < expr.Length && expr[j] != '\'') j++; if (j < expr.Length) j++; }
                while (j < expr.Length && (char.IsLetterOrDigit(expr[j]) || expr[j] is '_' or '.' or '$')) j++;
                if (j < expr.Length && expr[j] == '!')
                {
                    j++;
                    while (j < expr.Length && (char.IsLetterOrDigit(expr[j]) || expr[j] is '$' or ':')) j++;
                }
                else if (j < expr.Length && expr[j] == ':')
                {
                    j++;
                    while (j < expr.Length && (char.IsLetterOrDigit(expr[j]) || expr[j] == '$')) j++;
                }

                var s = expr[i..j].ToUpperInvariant();
                var tt = s is "TRUE" or "FALSE" ? TT.Bool : TT.Name;
                tokens.Add(new Tok(tt, s)); i = j; continue;
            }

            // Two-char operators
            if (i + 1 < expr.Length)
            {
                var two = expr[i..(i + 2)];
                if (two is "<=" or ">=" or "<>")
                {
                    tokens.Add(new Tok(TT.Op, two)); i += 2; continue;
                }
            }

            var tok = c switch
            {
                '(' => new Tok(TT.LParen, "("),
                ')' => new Tok(TT.RParen, ")"),
                '{' => new Tok(TT.LBrace, "{"),
                '}' => new Tok(TT.RBrace, "}"),
                ',' => new Tok(TT.Comma,  ","),
                ';' => new Tok(TT.Comma,  ","),
                '+' => new Tok(TT.Op, "+"),
                '-' => new Tok(TT.Op, "-"),
                '*' => new Tok(TT.Op, "*"),
                '/' => new Tok(TT.Op, "/"),
                '^' => new Tok(TT.Op, "^"),
                '&' => new Tok(TT.Op, "&"),
                '%' => new Tok(TT.Op, "%"),
                '=' => new Tok(TT.Op, "="),
                '<' => new Tok(TT.Op, "<"),
                '>' => new Tok(TT.Op, ">"),
                _ => new Tok(TT.Op, c.ToString())
            };
            tokens.Add(tok); i++;
        }
        return tokens;
    }

    [GeneratedRegex(@"(?<ca>\$?)(?<col>[A-Z]{1,3})(?<ra>\$?)(?<row>\d{1,7})", RegexOptions.Compiled)]
    private static partial Regex CellRefRegex();
}

public sealed class FormulaException(ExcelErrorCode code) : Exception
{
    public ExcelErrorCode Code { get; } = code;
}

// ── Dependency engine ─────────────────────────────────────────────────────────

public sealed partial class FormulaDependencyEngine
{
    private readonly ExcelWorkbook _workbook;
    private readonly Dictionary<CellAddress, HashSet<CellAddress>> _precs = new();
    private readonly Dictionary<CellAddress, HashSet<CellAddress>> _deps = new();
    private readonly HashSet<CellAddress> _dirty = new();

    public bool FollowDependencyChain { get; set; } = true;
    public bool AllowCircularReferences { get; set; }
    public int MaxIterations { get; set; } = 100;
    public double IterationDelta { get; set; } = 0.001;

    public FormulaDependencyEngine(ExcelWorkbook wb) => _workbook = wb;

    public void BuildGraph()
    {
        _precs.Clear(); _deps.Clear(); _dirty.Clear();
        foreach (var ws in _workbook.Worksheets)
            foreach (var cell in ws.AllCells())
            {
                if (cell.Formula == null) continue;
                var addr = new CellAddress(ws.Name, cell.Row, cell.Col);
                foreach (var prec in ExtractRefs(cell.Formula, ws))
                {
                    if (!_precs.TryGetValue(addr, out var ps)) { ps = new(); _precs[addr] = ps; }
                    ps.Add(prec);
                    if (!_deps.TryGetValue(prec, out var ds)) { ds = new(); _deps[prec] = ds; }
                    ds.Add(addr);
                }
                _dirty.Add(addr);
            }
    }

    public IEnumerable<CellAddress> GetCalculationOrder()
    {
        var result = new List<CellAddress>();
        var visited = new HashSet<CellAddress>();
        var inStack = new HashSet<CellAddress>();
        foreach (var ws in _workbook.Worksheets)
            foreach (var cell in ws.AllCells())
            {
                if (cell.Formula == null) continue;
                var addr = new CellAddress(ws.Name, cell.Row, cell.Col);
                if (!visited.Contains(addr))
                    Dfs(addr, visited, inStack, result);
            }
        return result;
    }

    private void Dfs(CellAddress addr, HashSet<CellAddress> visited,
        HashSet<CellAddress> inStack, List<CellAddress> result)
    {
        if (inStack.Contains(addr) && !AllowCircularReferences) return;
        if (visited.Contains(addr)) return;
        inStack.Add(addr);
        if (_precs.TryGetValue(addr, out var precs))
            foreach (var prec in precs)
                Dfs(prec, visited, inStack, result);
        inStack.Remove(addr);
        visited.Add(addr);
        result.Add(addr);
    }

    private IEnumerable<CellAddress> ExtractRefs(string formula, ExcelWorksheet ws)
    {
        if (formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
            yield break;

        foreach (System.Text.RegularExpressions.Match m in RefPattern().Matches(formula))
        {
            var sheet = m.Groups["sheet"].Value;
            var cellPart = m.Groups["cell"].Value;
            if (string.IsNullOrEmpty(cellPart)) continue;
            var targetWs = string.IsNullOrEmpty(sheet) ? ws : _workbook.GetWorksheet(sheet) ?? ws;
            if (cellPart.Contains(':'))
            {
                (int fr, int fc, int tr, int tc)? r = null;
                try { r = ExcelAddressParser.ParseRange(cellPart); }
                catch { }

                if (r != null)
                {
                    var (fr, fc, tr, tc) = r.Value;
                    for (int r2 = fr; r2 <= tr; r2++)
                        for (int c = fc; c <= tc; c++)
                            yield return new CellAddress(targetWs.Name, r2, c);
                }
            }
            else
            {
                (int r, int c)? cell = null;
                try { cell = ExcelAddressParser.ParseCell(cellPart); }
                catch { }

                if (cell != null)
                    yield return new CellAddress(targetWs.Name, cell.Value.r, cell.Value.c);
            }
        }
    }

    public void MarkDirty(CellAddress addr) => _dirty.Add(addr);
    public void ClearDirty(CellAddress addr) => _dirty.Remove(addr);
    public int DirtyCount => _dirty.Count;

    [GeneratedRegex(
        @"(?:'(?<sheet>[^']+)'!|(?<sheet>[A-Za-z0-9_]+)!)?(?<cell>\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?)",
        RegexOptions.Compiled)]
    private static partial System.Text.RegularExpressions.Regex RefPattern();
}

public readonly record struct CellAddress(string SheetName, int Row, int Col)
{
    public override string ToString() => $"'{SheetName}'!{ExcelAddressParser.ToAddress(Row, Col)}";
}

// ── Formula parser manager ────────────────────────────────────────────────────

public sealed class FormulaParserManager
{
    private TextWriter? _log;
    private readonly ExcelWorkbook _workbook;

    public FormulaParserManager(ExcelWorkbook wb) => _workbook = wb;

    public void AttachLogger(FileInfo f)
    {
        _log = new StreamWriter(f.FullName, false, System.Text.Encoding.UTF8) { AutoFlush = true };
        Log($"EPExcel.ML formula logger attached {DateTime.UtcNow:O}");
    }

    public void AttachLogger(TextWriter w)
    {
        _log = w;
        Log($"EPExcel.ML formula logger attached {DateTime.UtcNow:O}");
    }

    public void DetachLogger() { _log?.Flush(); _log = null; }

    internal void Log(string msg) =>
        _log?.WriteLine($"[{DateTime.UtcNow:HH:mm:ss.fff}] {msg}");

    internal void LogError(string formula, string sheet, int row, int col, Exception ex) =>
        Log($"ERROR '{sheet}'!{ExcelAddressParser.ToAddress(row, col)} ={formula}: {ex.Message}");
}

// ── Lambda registry ────────────────────────────────────────────────────────────

public sealed class LambdaRegistry
{
    private readonly Dictionary<string, Func<object?[], ExcelWorksheet, object?>> _lambdas =
        new(StringComparer.OrdinalIgnoreCase);

    public void Register(string name, Func<object?[], ExcelWorksheet, object?> fn) =>
        _lambdas[name] = fn;

    public Func<object?[], ExcelWorksheet, object?>? TryGet(string name) =>
        _lambdas.TryGetValue(name, out var fn) ? fn : null;
}
