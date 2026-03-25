namespace EPExcel.ML;

/// <summary>
/// EPExcel-compatible fluent style API.
/// Usage: ws.Cells("A1:C3").Style.Font.Bold = true;
///        ws.Cells("A1").Style.Fill.SetBackground("FF0000");
/// Also supports EPExcel-style: ws.Cell(1,1).StyleIndex = wb.Styles.RegisterStyle(myStyle);
/// </summary>
public sealed class ExcelRangeStyle
{
    private readonly ExcelRange _range;
    private readonly ExcelWorkbook? _wb;

    public ExcelRangeStyle(ExcelRange range)
    {
        _range = range;
        _wb = range.Worksheet.GetWorkbook();
    }

    public ExcelRangeFontStyle Font => new(_range, _wb);
    public ExcelRangeFillStyle Fill => new(_range, _wb);
    public ExcelRangeBorderStyle Border => new(_range, _wb);
    public ExcelRangeAlignmentStyle Alignment => new(_range, _wb);

    public string? NumberFormat
    {
        get
        {
            var cell = _range.Worksheet.GetCell(_range.FromRow, _range.FromCol);
            return cell?.NumberFormat;
        }
        set
        {
            for (int r = _range.FromRow; r <= _range.ToRow; r++)
                for (int c = _range.FromCol; c <= _range.ToCol; c++)
                {
                    var cell = _range.Worksheet.Cell(r, c);
                    cell.NumberFormat = value;
                    if (_wb != null)
                    {
                        var st = _wb.Styles.GetStyle(cell.StyleIndex).Clone();
                        st.NumberFormat = value;
                        if (value != null) _wb.Styles.RegisterNumberFormat(value);
                        cell.StyleIndex = _wb.Styles.RegisterStyle(st);
                    }
                }
        }
    }

    public bool WrapText
    {
        get => _wb?.Styles.GetStyle(_range.Worksheet.GetCell(_range.FromRow, _range.FromCol)?.StyleIndex ?? 0).WrapText ?? false;
        set => ApplyToAll(st => st.WrapText = value);
    }

    public bool Locked
    {
        get => _wb?.Styles.GetStyle(_range.Worksheet.GetCell(_range.FromRow, _range.FromCol)?.StyleIndex ?? 0).Locked ?? true;
        set => ApplyToAll(st => st.Locked = value);
    }

    internal void ApplyToAll(Action<CellStyleDef> modify)
    {
        if (_wb == null) return;
        for (int r = _range.FromRow; r <= _range.ToRow; r++)
            for (int c = _range.FromCol; c <= _range.ToCol; c++)
            {
                var cell = _range.Worksheet.Cell(r, c);
                var st = _wb.Styles.GetStyle(cell.StyleIndex).Clone();
                modify(st);
                cell.StyleIndex = _wb.Styles.RegisterStyle(st);
            }
    }
}

// ── Font fluent ───────────────────────────────────────────────────────────────

public sealed class ExcelRangeFontStyle
{
    private readonly ExcelRange _range;
    private readonly ExcelWorkbook? _wb;
    public ExcelRangeFontStyle(ExcelRange r, ExcelWorkbook? wb) { _range = r; _wb = wb; }

    private CellStyleDef Current => _wb?.Styles.GetStyle(
        _range.Worksheet.GetCell(_range.FromRow, _range.FromCol)?.StyleIndex ?? 0)
        ?? new CellStyleDef();

    private void Set(Action<CellStyleDef> m) => new ExcelRangeStyle(_range).ApplyToAll(m);

    public bool Bold { get => Current.Font.Bold; set => Set(s => s.Font.Bold = value); }
    public bool Italic { get => Current.Font.Italic; set => Set(s => s.Font.Italic = value); }
    public bool Underline { get => Current.Font.Underline; set => Set(s => s.Font.Underline = value); }
    public bool Strikethrough { get => Current.Font.Strikethrough; set => Set(s => s.Font.Strikethrough = value); }
    public double Size { get => Current.Font.Size; set => Set(s => s.Font.Size = value); }
    public string Name { get => Current.Font.Name; set => Set(s => s.Font.Name = value); }
    public string? Color { get => Current.Font.Color; set => Set(s => s.Font.Color = value); }
    public int ThemeColor { get => Current.Font.ThemeColor; set => Set(s => s.Font.ThemeColor = value); }
    public double Tint { get => Current.Font.Tint; set => Set(s => s.Font.Tint = value); }
    public ExcelVerticalAlignment VerticalAlign { get => Current.Font.VerticalAlign; set => Set(s => s.Font.VerticalAlign = value); }

    /// <summary>Set color from ExcelColor instance.</summary>
    public void SetColor(ExcelColor color) => Color = color.ToHex();
}

// ── Fill fluent ───────────────────────────────────────────────────────────────

public sealed class ExcelRangeFillStyle
{
    private readonly ExcelRange _range;
    private readonly ExcelWorkbook? _wb;
    public ExcelRangeFillStyle(ExcelRange r, ExcelWorkbook? wb) { _range = r; _wb = wb; }

    private CellStyleDef Current => _wb?.Styles.GetStyle(
        _range.Worksheet.GetCell(_range.FromRow, _range.FromCol)?.StyleIndex ?? 0)
        ?? new CellStyleDef();

    private void Set(Action<CellStyleDef> m) => new ExcelRangeStyle(_range).ApplyToAll(m);

    public ExcelFillPattern PatternType { get => Current.Fill.PatternType; set => Set(s => s.Fill.PatternType = value); }
    public ExcelColorStyle BackgroundColor => new(_range, _wb, true);
    public ExcelColorStyle ForegroundColor => new(_range, _wb, false);

    /// <summary>Set solid background color from hex string. EPExcel: Fill.BackgroundColor.SetColor()</summary>
    public void SetBackground(string hexColor) => BackgroundColor.SetColor(hexColor);

    /// <summary>Set solid background color from ExcelColor.</summary>
    public void SetBackground(ExcelColor color) => BackgroundColor.SetColor(color.ToHex());

    /// <summary>Set gradient fill.</summary>
    public void SetGradient(double degree, params (double position, string color)[] stops)
    {
        var grad = new ExcelGradientFill { Degree = degree };
        foreach (var (pos, col) in stops) grad.Stops.Add((pos, col));
        Set(s => { s.Fill.Gradient = grad; s.Fill.PatternType = ExcelFillPattern.None; });
    }

    /// <summary>Clear fill to none.</summary>
    public void Clear() => Set(s => { s.Fill.PatternType = ExcelFillPattern.None; });
}

public sealed class ExcelColorStyle(ExcelRange r, ExcelWorkbook? wb, bool isBg)
{
    public void SetColor(string hex)
    {
        new ExcelRangeStyle(r).ApplyToAll(s =>
        {
            if (isBg) { s.Fill.BackgroundColor = hex; s.Fill.PatternType = ExcelFillPattern.Solid; }
            else s.Fill.ForegroundColor = hex;
        });
    }

    public void SetColor(ExcelColor color) => SetColor(color.ToHex());
}

// ── Border fluent ─────────────────────────────────────────────────────────────

public sealed class ExcelRangeBorderStyle
{
    private readonly ExcelRange _range;
    private readonly ExcelWorkbook? _wb;
    public ExcelRangeBorderStyle(ExcelRange r, ExcelWorkbook? wb) { _range = r; _wb = wb; }

    private void Set(Action<CellStyleDef> m) => new ExcelRangeStyle(_range).ApplyToAll(m);

    public ExcelRangeBorderSide Top => new(_range, _wb, "Top");
    public ExcelRangeBorderSide Bottom => new(_range, _wb, "Bottom");
    public ExcelRangeBorderSide Left => new(_range, _wb, "Left");
    public ExcelRangeBorderSide Right => new(_range, _wb, "Right");
    public ExcelRangeBorderSide Diagonal => new(_range, _wb, "Diagonal");

    public bool DiagonalUp { set => Set(s => s.Border.DiagonalUp = value); }
    public bool DiagonalDown { set => Set(s => s.Border.DiagonalDown = value); }

    /// <summary>EPExcel parity: Border.BorderAround(ExcelBorderStyle.Thin)</summary>
    public void BorderAround(BorderLineStyle style, string? color = null) =>
        _range.BorderAround(style, color);

    /// <summary>Set all four sides at once.</summary>
    public void SetAll(BorderLineStyle style, string? color = null) =>
        Set(s => s.Border.SetAll(style, color));
}

public sealed class ExcelRangeBorderSide
{
    private readonly ExcelRange _range;
    private readonly ExcelWorkbook? _wb;
    private readonly string _side;

    public ExcelRangeBorderSide(ExcelRange r, ExcelWorkbook? wb, string side) { _range = r; _wb = wb; _side = side; }

    private void Set(Action<CellStyleDef> m) => new ExcelRangeStyle(_range).ApplyToAll(m);

    public BorderLineStyle Style
    {
        set => Set(s =>
        {
            var b = _side switch { "Top" => s.Border.Top, "Bottom" => s.Border.Bottom,
                "Left" => s.Border.Left, "Right" => s.Border.Right, _ => s.Border.Diagonal };
            b.Style = value;
        });
    }

    public string? Color
    {
        set => Set(s =>
        {
            var b = _side switch { "Top" => s.Border.Top, "Bottom" => s.Border.Bottom,
                "Left" => s.Border.Left, "Right" => s.Border.Right, _ => s.Border.Diagonal };
            b.Color = value;
        });
    }
}

// ── Alignment fluent ──────────────────────────────────────────────────────────

public sealed class ExcelRangeAlignmentStyle
{
    private readonly ExcelRange _range;
    private readonly ExcelWorkbook? _wb;
    public ExcelRangeAlignmentStyle(ExcelRange r, ExcelWorkbook? wb) { _range = r; _wb = wb; }

    private CellStyleDef Current => _wb?.Styles.GetStyle(
        _range.Worksheet.GetCell(_range.FromRow, _range.FromCol)?.StyleIndex ?? 0)
        ?? new CellStyleDef();

    private void Set(Action<CellStyleDef> m) => new ExcelRangeStyle(_range).ApplyToAll(m);

    public ExcelHorizontalAlignment Horizontal { get => Current.Alignment.Horizontal; set => Set(s => s.Alignment.Horizontal = value); }
    public ExcelVerticalCellAlignment Vertical { get => Current.Alignment.Vertical; set => Set(s => s.Alignment.Vertical = value); }
    public bool WrapText { get => Current.Alignment.WrapText; set => Set(s => s.Alignment.WrapText = value); }
    public int Indent { get => Current.Alignment.Indent; set => Set(s => s.Alignment.Indent = value); }
    public int TextRotation { get => Current.Alignment.TextRotation; set => Set(s => s.Alignment.TextRotation = value); }
    public bool ShrinkToFit { get => Current.Alignment.ShrinkToFit; set => Set(s => s.Alignment.ShrinkToFit = value); }
}

// ── ExcelWorksheet style convenience extensions ───────────────────────────────

public static class WorksheetStyleExtensions
{
    /// <summary>Get style for a single cell — EPExcel parity: ws.Cells["A1"].Style.Font.Bold = true</summary>
    public static ExcelRangeStyle StyleOf(this ExcelWorksheet ws, int row, int col) =>
        ws.Cells(row, col, row, col).Style;

    /// <summary>Get style for a range — EPExcel parity: ws.Cells["A1:C3"].Style.Fill.BackgroundColor.SetColor(...)</summary>
    public static ExcelRangeStyle StyleOf(this ExcelWorksheet ws, string address) =>
        ws.Cells(address).Style;

    /// <summary>Apply a predefined named style to cells.</summary>
    public static void ApplyNamedStyle(this ExcelWorksheet ws, ExcelRange range, string styleName)
    {
        var wb = ws.GetWorkbook();
        if (wb?.NamedStyles()[styleName] is { } ns)
        {
            var style = ns.Style;
            var idx = wb.Styles.RegisterStyle(style);
            for (int r = range.FromRow; r <= range.ToRow; r++)
                for (int c = range.FromCol; c <= range.ToCol; c++)
                    ws.Cell(r, c).StyleIndex = idx;
        }
    }
}

// ── NamedStyles on ExcelWorkbook ─────────────────────────────────────────────
// Extension to wire NamedStyles into ExcelWorkbook

public static class WorkbookNamedStyleExtensions
{
    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<ExcelWorkbook, ExcelNamedStyleCollection>
        _collections = new();

    public static ExcelNamedStyleCollection NamedStyles(this ExcelWorkbook wb) =>
        _collections.GetOrCreateValue(wb);

    /// <summary>Create a named style. EPExcel parity: wb.Styles.CreateNamedStyle("Heading1")</summary>
    public static ExcelNamedStyle CreateNamedStyle(this ExcelWorkbook wb, string name)
    {
        var ns = wb.NamedStyles().Add(name);
        return ns;
    }
}
