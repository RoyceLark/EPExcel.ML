using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace EPExcel.ML;

// ── Error types ───────────────────────────────────────────────────────────────

public enum ExcelErrorCode { Div0, Value, Ref, Name, Num, NA, Null, Spill, Calc }

public sealed class CellError(ExcelErrorCode code)
{
    public ExcelErrorCode Code { get; } = code;
    public override string ToString() => Code switch
    {
        ExcelErrorCode.Div0  => "#DIV/0!",
        ExcelErrorCode.Value => "#VALUE!",
        ExcelErrorCode.Ref   => "#REF!",
        ExcelErrorCode.Name  => "#NAME?",
        ExcelErrorCode.Num   => "#NUM!",
        ExcelErrorCode.NA    => "#N/A",
        ExcelErrorCode.Null  => "#NULL!",
        ExcelErrorCode.Spill => "#SPILL!",
        _                    => "#CALC!"
    };
}

// ── Cell value types ──────────────────────────────────────────────────────────

public enum CellValueType { Empty, String, Number, Boolean, Error, Formula }

// ── ExcelCell ─────────────────────────────────────────────────────────────────

public sealed class ExcelCell
{
    private object? _value;

    public int Row { get; internal set; }
    public int Col { get; internal set; }
    public string? Formula { get; set; }
    public object? CalculatedValue { get; set; }
    public bool IsSpill { get; set; }
    public bool IsArrayFormula { get; set; }
    public string? NumberFormat { get; set; }
    public int StyleIndex { get; set; }
    public string? Comment { get; set; }
    public string? HyperlinkUrl { get; set; }

    public object? Value
    {
        get => Formula != null ? CalculatedValue ?? _value : _value;
        set { _value = value; CalculatedValue = null; }
    }

    public object? DisplayValue => CalculatedValue ?? _value;

    public CellValueType CellType => _value switch
    {
        null when Formula == null => CellValueType.Empty,
        string                   => CellValueType.String,
        double or int or long or float or decimal => CellValueType.Number,
        bool                     => CellValueType.Boolean,
        CellError                => CellValueType.Error,
        _                        => Formula != null ? CellValueType.Formula : CellValueType.Empty
    };

    public bool IsEmpty => _value == null && Formula == null && CalculatedValue == null;

    public string GetString() => _value?.ToString() ?? CalculatedValue?.ToString() ?? "";
    public double GetDouble() => Formulas.FunctionLibrary.Num(DisplayValue);
    public bool GetBool() => Formulas.FunctionLibrary.Bool(DisplayValue);

    public void SetValue(string v)  { _value = v; }
    public void SetValue(double v)  { _value = v; }
    public void SetValue(bool v)    { _value = v; }
    public void SetValue(DateTime v){ _value = v.ToOADate(); }
    public void SetValue(int v)     { _value = (double)v; }
    public void SetValue(long v)    { _value = (double)v; }
    public void SetValue(decimal v) { _value = (double)v; }
    public void Clear() { _value = null; Formula = null; CalculatedValue = null; }
}

// ── Address parser ────────────────────────────────────────────────────────────

public static partial class ExcelAddressParser
{
    public static (int Row, int Col) ParseCell(string address)
    {
        address = address.Replace("$", "").Trim().ToUpperInvariant();
        var m = CellRegex().Match(address);
        if (!m.Success) throw new ArgumentException($"Invalid cell address: {address}");
        return (int.Parse(m.Groups[2].Value), ColumnLetterToNumber(m.Groups[1].Value));
    }

    public static (int FromRow, int FromCol, int ToRow, int ToCol) ParseRange(string address)
    {
        address = address.Replace("$", "").Trim().ToUpperInvariant();
        var parts = address.Split(':');
        if (parts.Length != 2) throw new ArgumentException($"Invalid range: {address}");
        var (r1, c1) = ParseCell(parts[0]);
        var (r2, c2) = ParseCell(parts[1]);
        return (Math.Min(r1, r2), Math.Min(c1, c2), Math.Max(r1, r2), Math.Max(c1, c2));
    }

    public static int ColumnLetterToNumber(string col)
    {
        int result = 0;
        foreach (char c in col.ToUpperInvariant())
            result = result * 26 + (c - 'A' + 1);
        return result;
    }

    public static string ColumnNumberToLetter(int col)
    {
        string result = "";
        while (col > 0)
        {
            col--;
            result = (char)('A' + col % 26) + result;
            col /= 26;
        }
        return result;
    }

    public static string ToAddress(int row, int col) =>
        $"{ColumnNumberToLetter(col)}{row}";

    public static string ToRangeAddress(int fromRow, int fromCol, int toRow, int toCol) =>
        $"{ToAddress(fromRow, fromCol)}:{ToAddress(toRow, toCol)}";

    [GeneratedRegex(@"^([A-Z]{1,3})(\d{1,7})$")]
    private static partial Regex CellRegex();
}

// ── ExcelWorksheet ────────────────────────────────────────────────────────────

public sealed class ExcelWorksheet : IDisposable
{
    private readonly ConcurrentDictionary<(int, int), ExcelCell> _cells = new();
    private readonly Dictionary<int, double> _rowHeights = new();
    private readonly Dictionary<int, double> _colWidths = new();
    private readonly Dictionary<int, bool> _hiddenRows = new();
    private readonly Dictionary<int, bool> _hiddenCols = new();
    private ExcelWorkbook? _workbook;

    public string Name { get; internal set; }
    public int Index { get; internal set; }

    // Collections
    public List<ExcelTable> Tables { get; } = new();
    public List<ExcelPivotTable> PivotTables { get; } = new();
    public List<ExcelChart> Charts { get; } = new();
    public List<ExcelImage> Images { get; } = new();
    public List<ExcelShape> Shapes { get; } = new();
    public List<ExcelComment> Comments { get; } = new();
    public List<ExcelThreadedComment> ThreadedComments { get; } = new();
    public List<ConditionalFormattingRule> ConditionalFormattings { get; } = new();
    public List<ExcelDataValidation> DataValidations { get; } = new();
    public List<ExcelSparklineGroup> SparklineGroups { get; } = new();
    public List<ExcelSlicer> Slicers { get; } = new();
    public List<ExcelInCellCheckBox> InCellCheckBoxes { get; } = new();
    public List<ExcelQueryTable> QueryTables { get; } = new();
    public ExcelPageBreaks PageBreaks { get; } = new();
    public ExcelOutlineCollection Outline { get; } = new();

    // Sheet properties
    public ExcelPageSetup PageSetup { get; } = new();
    public bool ShowGridLines { get; set; } = true;
    public bool ShowRowColHeaders { get; set; } = true;
    public bool TabSelected { get; set; }
    public string? TabColor { get; set; }
    public bool Protected { get; set; }
    public string? PasswordHash { get; set; }
    public bool FreezeRows { get; set; }
    public bool FreezeCols { get; set; }
    public int FreezeRow { get; set; }
    public int FreezeCol { get; set; }
    public List<string> MergedCells { get; } = new();
    public ExcelSheetView View { get; } = new();
    public int DefaultRowHeight { get; set; } = 15;
    public int DefaultColWidth { get; set; } = 64;
    public string? CodeName { get; set; }
    public bool OutlineSymbolsBelow { get; set; } = true;
    public bool OutlineSymbolsRight { get; set; } = true;

    // Auto-filter
    public string? AutoFilterAddress { get; set; }

    internal ExcelWorksheet(string name, int index) { Name = name; Index = index; }

    internal void SetWorkbook(ExcelWorkbook wb) => _workbook = wb;
    public ExcelWorkbook? GetWorkbook() => _workbook;

    // ── Cell access ───────────────────────────────────────────────────────────

    public ExcelCell Cell(int row, int col)
    {
        if (row < 1 || col < 1) throw new ArgumentOutOfRangeException($"Row/Col must be >= 1. Got ({row},{col})");
        return _cells.GetOrAdd((row, col), k => new ExcelCell { Row = k.Item1, Col = k.Item2 });
    }

    public ExcelCell Cell(string address)
    {
        var (r, c) = ExcelAddressParser.ParseCell(address);
        return Cell(r, c);
    }

    public ExcelCell? GetCell(int row, int col) =>
        _cells.TryGetValue((row, col), out var cell) ? cell : null;

    public ExcelRange Cells(int fromRow, int fromCol, int toRow, int toCol) =>
        new(this, fromRow, fromCol, toRow, toCol);

    public ExcelRange Cells(string address)
    {
        if (address.Contains(':'))
        {
            var (fr, fc, tr, tc) = ExcelAddressParser.ParseRange(address);
            return new ExcelRange(this, fr, fc, tr, tc);
        }
        var (r, c) = ExcelAddressParser.ParseCell(address);
        return new ExcelRange(this, r, c, r, c);
    }

    //public ExcelRange Cells => new(this, 1, 1, MaxRow, MaxCol);

    public int MaxRow => _cells.Keys.Select(k => k.Item1).DefaultIfEmpty(0).Max();
    public int MaxCol => _cells.Keys.Select(k => k.Item2).DefaultIfEmpty(0).Max();
    public int Dimension => _cells.Count;

    public IEnumerable<ExcelCell> AllCells() => _cells.Values;

    // ── Row/Column operations ──────────────────────────────────────────────────

    public void SetRowHeight(int row, double height) => _rowHeights[row] = height;
    public void SetColumnWidth(int col, double width) => _colWidths[col] = width;
    public void SetColumnWidth(string col, double width) =>
        SetColumnWidth(ExcelAddressParser.ColumnLetterToNumber(col), width);

    public double GetRowHeight(int row) => _rowHeights.TryGetValue(row, out var h) ? h : DefaultRowHeight;
    public double GetColumnWidth(int col) => _colWidths.TryGetValue(col, out var w) ? w : DefaultColWidth;

    public void HideRow(int row) => _hiddenRows[row] = true;
    public void HideColumn(int col) => _hiddenCols[col] = true;
    public bool IsRowHidden(int row) => _hiddenRows.TryGetValue(row, out var h) && h;
    public bool IsColHidden(int col) => _hiddenCols.TryGetValue(col, out var h) && h;

    public IReadOnlyDictionary<int, double> RowHeights => _rowHeights;
    public IReadOnlyDictionary<int, double> ColWidths => _colWidths;

    public void InsertRow(int row, int count = 1)
    {
        var toMove = _cells.Keys.Where(k => k.Item1 >= row).OrderByDescending(k => k.Item1).ToList();
        foreach (var key in toMove)
        {
            if (_cells.TryRemove(key, out var cell))
            {
                cell.Row = key.Item1 + count;
                if (cell.Formula != null)
                    cell.Formula = Formulas.FormulaEngine.ShiftFormula(cell.Formula, ShiftType.Row, row, count);
                _cells[(cell.Row, cell.Col)] = cell;
            }
        }
        // Shift row heights
        var rhKeys = _rowHeights.Keys.Where(r => r >= row).OrderByDescending(r => r).ToList();
        foreach (var r in rhKeys)
        {
            _rowHeights[r + count] = _rowHeights[r];
            _rowHeights.Remove(r);
        }
    }

    public void DeleteRow(int row, int count = 1)
    {
        var toDelete = _cells.Keys.Where(k => k.Item1 >= row && k.Item1 < row + count).ToList();
        foreach (var key in toDelete) _cells.TryRemove(key, out _);
        var toMove = _cells.Keys.Where(k => k.Item1 > row + count - 1).OrderBy(k => k.Item1).ToList();
        foreach (var key in toMove)
        {
            if (_cells.TryRemove(key, out var cell))
            {
                cell.Row = key.Item1 - count;
                if (cell.Formula != null)
                    cell.Formula = Formulas.FormulaEngine.ShiftFormula(cell.Formula, ShiftType.Row, row, -count);
                _cells[(cell.Row, cell.Col)] = cell;
            }
        }
    }

    public void InsertColumn(int col, int count = 1)
    {
        var toMove = _cells.Keys.Where(k => k.Item2 >= col).OrderByDescending(k => k.Item2).ToList();
        foreach (var key in toMove)
        {
            if (_cells.TryRemove(key, out var cell))
            {
                cell.Col = key.Item2 + count;
                if (cell.Formula != null)
                    cell.Formula = Formulas.FormulaEngine.ShiftFormula(cell.Formula, ShiftType.Column, col, count);
                _cells[(cell.Row, cell.Col)] = cell;
            }
        }
    }

    public void DeleteColumn(int col, int count = 1)
    {
        var toDelete = _cells.Keys.Where(k => k.Item2 >= col && k.Item2 < col + count).ToList();
        foreach (var key in toDelete) _cells.TryRemove(key, out _);
        var toMove = _cells.Keys.Where(k => k.Item2 > col + count - 1).OrderBy(k => k.Item2).ToList();
        foreach (var key in toMove)
        {
            if (_cells.TryRemove(key, out var cell))
            {
                cell.Col = key.Item2 - count;
                if (cell.Formula != null)
                    cell.Formula = Formulas.FormulaEngine.ShiftFormula(cell.Formula, ShiftType.Column, col, -count);
                _cells[(cell.Row, cell.Col)] = cell;
            }
        }
    }

    // ── Formula calculation ────────────────────────────────────────────────────

    public void Calculate() => _workbook?.Calculate();

    internal void CalculateWithEngine(Formulas.FormulaEngine engine)
    {
        foreach (var cell in _cells.Values.Where(c => c.Formula != null))
        {
            try { cell.CalculatedValue = engine.Evaluate(cell.Formula!, this); }
            catch { cell.CalculatedValue = new CellError(ExcelErrorCode.Value); }
        }
    }

    // ── Pivot tables ───────────────────────────────────────────────────────────

    public ExcelPivotTable AddPivotTable(string name, ExcelRange dataRange, string outputCell)
    {
        var pt = new ExcelPivotTable
        {
            Name = name,
            DataRange = dataRange.Address,
            DataSheetName = dataRange.Worksheet.Name,
            OutputRange = outputCell,
        };
        PivotTables.Add(pt);
        return pt;
    }

    // ── Tables ─────────────────────────────────────────────────────────────────

    public ExcelTable AddTable(ExcelRange range, string name)
    {
        var t = new ExcelTable(name, range.Address) { ShowHeader = true };
        Tables.Add(t);
        return t;
    }

    // ── Charts ─────────────────────────────────────────────────────────────────

    public ExcelChart AddChart(ExcelChartType type, string name) =>
        AddChart(type, name, 2, 2, 18, 14);

    public ExcelChart AddChart(ExcelChartType type, string name,
        int fromRow, int fromCol, int toRow, int toCol)
    {
        var chart = new ExcelChart(name, type)
        {
            FromRow = fromRow, FromCol = fromCol,
            ToRow = toRow, ToCol = toCol,
        };
        Charts.Add(chart);
        return chart;
    }

    // ── Images ─────────────────────────────────────────────────────────────────

    public ExcelImage AddImage(byte[] imageData, string name,
        int fromRow, int fromCol, int toRow, int toCol)
    {
        var img = new ExcelImage(name, imageData)
        {
            FromRow = fromRow, FromCol = fromCol,
            ToRow = toRow, ToCol = toCol,
        };
        Images.Add(img);
        return img;
    }

    // ── Freeze panes ───────────────────────────────────────────────────────────

    public void FreezePanes(int row, int col)
    {
        FreezeRows = row > 0;
        FreezeCols = col > 0;
        FreezeRow = row;
        FreezeCol = col;
    }

    // ── Conditional formatting ─────────────────────────────────────────────────

    public ConditionalFormattingRule AddConditionalFormatting(string address,
        ConditionalFormattingType type)
    {
        var rule = new ConditionalFormattingRule
        {
            Address = address, Type = type,
            Priority = ConditionalFormattings.Count + 1
        };
        ConditionalFormattings.Add(rule);
        return rule;
    }

    // ── Data validation ────────────────────────────────────────────────────────

    public ExcelDataValidation AddDataValidation(string address,
        DataValidationType type = DataValidationType.List)
    {
        var dv = new ExcelDataValidation { Address = address, Type = type };
        DataValidations.Add(dv);
        return dv;
    }

    // ── In-cell checkboxes (EPExcel 7.5) ───────────────────────────────────────

    public ExcelInCellCheckBox AddInCellCheckBox(int row, int col, bool isChecked = false)
    {
        var cb = new ExcelInCellCheckBox(row, col, isChecked);
        InCellCheckBoxes.Add(cb);
        Cell(row, col).Value = isChecked;
        return cb;
    }

    // ── Auto filter ────────────────────────────────────────────────────────────

    public void SetAutoFilter(string rangeAddress) => AutoFilterAddress = rangeAddress;

    public void Dispose() { }
}

// ── ExcelWorkbook ─────────────────────────────────────────────────────────────

public sealed class ExcelWorkbook : IDisposable, IAsyncDisposable
{
    private readonly List<ExcelWorksheet> _sheets = new();
    private readonly Dictionary<string, ExcelNamedRange> _namedRanges =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly ILogger<ExcelWorkbook> _logger;

    public IReadOnlyList<ExcelWorksheet> Worksheets => _sheets.AsReadOnly();
    public IReadOnlyDictionary<string, ExcelNamedRange> NamedRanges => _namedRanges;
    public ExcelWorkbookProperties Properties { get; } = new();
    public ExcelWorkbookProtection Protection { get; } = new();
    public ExcelWorkbookView View { get; } = new();
    public ExcelCompatibilitySettings Compatibility { get; } = new();
    public ExcelTheme Theme { get; set; } = ExcelTheme.Office;
    public byte[]? VbaProjectBytes { get; set; }
    public bool IsMacroEnabled => VbaProjectBytes?.Length > 0;
    public ExcelSensitivityLabel? SensitivityLabel { get; set; }
    public List<ExcelExternalLink> ExternalLinks { get; } = new();
    public ExcelConnectionCollection Connections { get; } = new();
    public List<ExcelCustomTableStyle> CustomTableStyles { get; } = new();
    public Dictionary<string, byte[]> UnknownParts { get; } =
        new(StringComparer.OrdinalIgnoreCase);
    public LambdaRegistry Lambdas { get; } = new();
    public Formulas.FormulaParserManager FormulaParserManager { get; }
    public Formulas.FormulaDependencyEngine DependencyEngine { get; }
    public bool FullPrecision { get; set; } = true;
    public StyleRegistry Styles { get; }

    public ExcelWorkbook(ILogger<ExcelWorkbook>? logger = null)
    {
        _logger = logger ?? NullLogger<ExcelWorkbook>.Instance;
        FormulaParserManager = new Formulas.FormulaParserManager(this);
        DependencyEngine = new Formulas.FormulaDependencyEngine(this);
        Styles = new StyleRegistry();
    }

    // ── Worksheet management ──────────────────────────────────────────────────

    public ExcelWorksheet AddWorksheet(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Name required");
        if (name.Length > 31) throw new ArgumentException("Name must be ≤ 31 chars");
        if (_sheets.Any(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"Sheet '{name}' already exists");
        var ws = new ExcelWorksheet(name, _sheets.Count + 1);
        ws.SetWorkbook(this);
        _sheets.Add(ws);
        return ws;
    }

    public ExcelWorksheet? GetWorksheet(string name) =>
        _sheets.FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

    public ExcelWorksheet GetWorksheet(int index)
    {
        if (index < 0 || index >= _sheets.Count)
            throw new ArgumentOutOfRangeException(nameof(index));
        return _sheets[index];
    }

    public bool RemoveWorksheet(string name)
    {
        var ws = GetWorksheet(name);
        return ws != null && _sheets.Remove(ws);
    }

    public void MoveWorksheet(string name, int toIndex)
    {
        var ws = GetWorksheet(name) ?? throw new ArgumentException($"Sheet '{name}' not found");
        _sheets.Remove(ws);
        _sheets.Insert(Math.Clamp(toIndex, 0, _sheets.Count), ws);
    }

    public ExcelWorksheet CopyWorksheet(string source, string dest)
    {
        var src = GetWorksheet(source) ?? throw new ArgumentException($"Sheet '{source}' not found");
        var dst = AddWorksheet(dest);
        foreach (var cell in src.AllCells())
        {
            var dc = dst.Cell(cell.Row, cell.Col);
            dc.Value = cell.Value;
            dc.Formula = cell.Formula;
            dc.NumberFormat = cell.NumberFormat;
            dc.StyleIndex = cell.StyleIndex;
        }
        return dst;
    }

    // ── Named ranges ─────────────────────────────────────────────────────────

    public ExcelNamedRange AddNamedRange(string name, ExcelWorksheet ws, ExcelRange range)
    {
        var nr = new ExcelNamedRange(name, ws, range);
        _namedRanges[name] = nr;
        return nr;
    }

    // ── Calculation ───────────────────────────────────────────────────────────

    public void Calculate(bool precisionAsDisplayed = false)
    {
        FormulaParserManager.Log("Calculate() started");

        // Step 1: Pivot tables first (GETPIVOTDATA needs them)
        var pivotEngine = new IO.PivotCalculationEngine(this);
        pivotEngine.CalculateAll(refreshCache: true);

        // Step 2: Build dependency graph for correct ordering
        DependencyEngine.BuildGraph();

        // Step 3: Create formula engine
        var engine = new Formulas.FormulaEngine(this)
        {
            RoundingStrategy = Formulas.FormulaEngine.PrecisionStrategy.Excel
        };

        // Step 4: Calculate in topological order
        int n = 0;
        foreach (var addr in DependencyEngine.GetCalculationOrder())
        {
            var ws = GetWorksheet(addr.SheetName);
            if (ws == null) continue;
            var cell = ws.GetCell(addr.Row, addr.Col);
            if (cell?.Formula == null) continue;
            try
            {
                cell.CalculatedValue = engine.Evaluate(cell.Formula, ws);
                DependencyEngine.ClearDirty(addr);
                n++;
            }
            catch (Exception ex)
            {
                FormulaParserManager.LogError(cell.Formula, addr.SheetName, addr.Row, addr.Col, ex);
                cell.CalculatedValue = new CellError(ExcelErrorCode.Value);
            }
        }

        FormulaParserManager.Log($"Calculate() done: {n} cells");
        _logger.LogInformation("Calculated {N} cells", n);
    }

    public void Calculate(Action<ExcelCalculationOptions> configure)
    {
        var opts = new ExcelCalculationOptions();
        configure(opts);
        DependencyEngine.FollowDependencyChain = opts.FollowDependencyChain;
        DependencyEngine.AllowCircularReferences = opts.AllowCircularReferences;
        DependencyEngine.MaxIterations = opts.MaxIterations;
        Calculate(opts.PrecisionAndRoundingStrategy == ExcelPrecisionStrategy.Excel);
    }

    // ── Table styles ──────────────────────────────────────────────────────────

    public ExcelCustomTableStyle CreateTableStyle(string name)
    {
        var s = new ExcelCustomTableStyle(name);
        CustomTableStyles.Add(s);
        return s;
    }

    public void Dispose() { foreach (var ws in _sheets) ws.Dispose(); }

    public async ValueTask DisposeAsync()
    {
        foreach (var ws in _sheets) ws.Dispose();
        await ValueTask.CompletedTask;
    }
}

// ── Supporting model types ────────────────────────────────────────────────────

public sealed class ExcelNamedRange(string name, ExcelWorksheet ws, ExcelRange range)
{
    public string Name { get; } = name;
    public ExcelWorksheet Worksheet { get; } = ws;
    public ExcelRange Range { get; } = range;
    public bool IsHidden { get; set; }
    public string? Comment { get; set; }
    public override string ToString() => $"'{Worksheet.Name}'!{Range.Address}";
}

public sealed class ExcelWorkbookProperties
{
    public string? Title { get; set; }
    public string? Subject { get; set; }
    public string? Author { get; set; }
    public string? Company { get; set; }
    public string? Category { get; set; }
    public string? Keywords { get; set; }
    public string? Comments { get; set; }
    public DateTime Created { get; set; } = DateTime.UtcNow;
    public DateTime Modified { get; set; } = DateTime.UtcNow;
    public string Application { get; set; } = "EPExcel.ML 1.0";
    public string AppVersion { get; set; } = "1.0.0";
}

public sealed class ExcelWorkbookProtection
{
    public bool LockStructure { get; set; }
    public bool LockWindows { get; set; }
    public string? PasswordHash { get; set; }
    public void SetPassword(string pwd) =>
        PasswordHash = Convert.ToBase64String(
            System.Security.Cryptography.SHA256.HashData(
                System.Text.Encoding.UTF8.GetBytes(pwd)));
}

public sealed class ExcelWorkbookView
{
    public int ActiveTab { get; set; }
    public int FirstSheet { get; set; }
    public int TabRatio { get; set; } = 600;
    public ExcelWorkbookVisibility Visibility { get; set; } = ExcelWorkbookVisibility.Visible;
    public bool ShowHorizontalScroll { get; set; } = true;
    public bool ShowVerticalScroll { get; set; } = true;
    public bool ShowSheetTabs { get; set; } = true;
    public int XWindow { get; set; }
    public int YWindow { get; set; }
    public int WindowWidth { get; set; } = 14400;
    public int WindowHeight { get; set; } = 8640;

    internal string ToOoxml() =>
        $"""<workbookView xWindow="{XWindow}" yWindow="{YWindow}" windowWidth="{WindowWidth}" windowHeight="{WindowHeight}" tabRatio="{TabRatio}" firstSheet="{FirstSheet}" activeTab="{ActiveTab}" showHorizontalScroll="{(ShowHorizontalScroll ? 1 : 0)}" showVerticalScroll="{(ShowVerticalScroll ? 1 : 0)}" showSheetTabs="{(ShowSheetTabs ? 1 : 0)}" visibility="{Visibility switch { ExcelWorkbookVisibility.Hidden => "hidden", ExcelWorkbookVisibility.VeryHidden => "veryHidden", _ => "visible" }}"/>""";
}

public enum ExcelWorkbookVisibility { Visible, Hidden, VeryHidden }

public sealed class ExcelCompatibilitySettings
{
    public bool IsWorksheet1Based { get; set; } = true;
    public bool ForceFullCalcOnLoad { get; set; }
    public bool Use1904DateSystem { get; set; }
    public ExcelCalcMode CalcMode { get; set; } = ExcelCalcMode.Automatic;
    public bool Iterate { get; set; }
    public int IterateCount { get; set; } = 100;
    public double IterateDelta { get; set; } = 0.001;
}

public enum ExcelCalcMode { Automatic, AutomaticExceptTables, Manual }

public sealed class ExcelSheetView
{
    public ExcelViewType View { get; set; } = ExcelViewType.Normal;
    public double Zoom { get; set; } = 100;
    public bool ShowFormulas { get; set; }
    public bool ShowZeros { get; set; } = true;
    public bool RightToLeft { get; set; }
    public string? TopLeftCell { get; set; }
}

public enum ExcelViewType { Normal, PageBreakPreview, PageLayout }

public sealed class ExcelPageSetup
{
    public ExcelPaperSize PaperSize { get; set; } = ExcelPaperSize.A4;
    public ExcelOrientation Orientation { get; set; } = ExcelOrientation.Default;
    public double LeftMargin { get; set; } = 0.7;
    public double RightMargin { get; set; } = 0.7;
    public double TopMargin { get; set; } = 0.75;
    public double BottomMargin { get; set; } = 0.75;
    public double HeaderMargin { get; set; } = 0.3;
    public double FooterMargin { get; set; } = 0.3;
    public bool FitToPage { get; set; }
    public int FitToWidth { get; set; } = 1;
    public int FitToHeight { get; set; }
    public int Scale { get; set; } = 100;
    public bool CenterHorizontally { get; set; }
    public bool CenterVertically { get; set; }
    public bool BlackAndWhite { get; set; }
    public bool Draft { get; set; }
    public string? OddHeader { get; set; }
    public string? OddFooter { get; set; }
    public string? EvenHeader { get; set; }
    public string? EvenFooter { get; set; }
    public string? FirstHeader { get; set; }
    public string? FirstFooter { get; set; }
    public bool DifferentOddEven { get; set; }
    public bool DifferentFirst { get; set; }
    public int Copies { get; set; } = 1;
}

public enum ExcelPaperSize { A4 = 9, Letter = 1, Legal = 5, A3 = 8, A5 = 11, B4 = 12, B5 = 13, Tabloid = 3 }
public enum ExcelOrientation { Default, Portrait, Landscape }

// ── In-cell checkbox ───────────────────────────────────────────────────────────

public sealed class ExcelInCellCheckBox(int row, int col, bool isChecked = false)
{
    public int Row { get; } = row;
    public int Col { get; } = col;
    public bool IsChecked { get; set; } = isChecked;
    public string? LinkedCell { get; set; }
}

// ── Sensitivity label ─────────────────────────────────────────────────────────

public sealed class ExcelSensitivityLabel
{
    public string? Id { get; set; }
    public string? Name { get; set; }
    public string? Description { get; set; }
    public string? Color { get; set; }
    public bool EncryptionEnabled { get; set; }
    public string? SiteId { get; set; }
    public string? SetDate { get; set; }
    public string? SetBy { get; set; }
}

// ── Calculation options ────────────────────────────────────────────────────────

public sealed class ExcelCalculationOptions
{
    public bool FollowDependencyChain { get; set; } = true;
    public bool AllowCircularReferences { get; set; }
    public int MaxIterations { get; set; } = 100;
    public double IterationDelta { get; set; } = 0.001;
    public ExcelPrecisionStrategy PrecisionAndRoundingStrategy { get; set; } = ExcelPrecisionStrategy.Excel;
}

public enum ExcelPrecisionStrategy { Excel, DotNet }

// ── Style registry ────────────────────────────────────────────────────────────

public sealed class StyleRegistry
{
    private readonly List<CellStyleDef> _styles = [new CellStyleDef()]; // default style at index 0
    private readonly Dictionary<string, int> _numberFormats = new();
    private int _nextNumFmtId = 164;

    public int RegisterStyle(CellStyleDef style)
    {
        // Check for existing matching style
        for (int i = 0; i < _styles.Count; i++)
            if (_styles[i].Matches(style)) return i;
        _styles.Add(style);
        return _styles.Count - 1;
    }

    public CellStyleDef GetStyle(int index) =>
        index >= 0 && index < _styles.Count ? _styles[index] : _styles[0];

    public IReadOnlyList<CellStyleDef> AllStyles => _styles.AsReadOnly();

    public int RegisterNumberFormat(string format)
    {
        if (_numberFormats.TryGetValue(format, out var id)) return id;
        id = _nextNumFmtId++;
        _numberFormats[format] = id;
        return id;
    }

    public IReadOnlyDictionary<string, int> CustomNumberFormats => _numberFormats;
}
