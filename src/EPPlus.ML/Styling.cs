namespace EPExcel.ML;

// ── Cell styling ──────────────────────────────────────────────────────────────

public sealed class CellStyleDef
{
    public string? NumberFormat { get; set; }
    public FontDef Font { get; set; } = new();
    public FillDef Fill { get; set; } = new();
    public BorderDef Border { get; set; } = new();
    public AlignmentDef Alignment { get; set; } = new();
    public bool WrapText { get; set; }
    public bool Locked { get; set; } = true;
    public bool Hidden { get; set; }

    public bool Matches(CellStyleDef other) =>
        NumberFormat == other.NumberFormat &&
        Font.Bold == other.Font.Bold &&
        Font.Italic == other.Font.Italic &&
        Font.Underline == other.Font.Underline &&
        Font.Size == other.Font.Size &&
        Font.Color == other.Font.Color &&
        Font.Name == other.Font.Name &&
        Fill.PatternType == other.Fill.PatternType &&
        Fill.BackgroundColor == other.Fill.BackgroundColor &&
        Fill.ForegroundColor == other.Fill.ForegroundColor &&
        Border.Top.Style == other.Border.Top.Style &&
        Border.Bottom.Style == other.Border.Bottom.Style &&
        Border.Left.Style == other.Border.Left.Style &&
        Border.Right.Style == other.Border.Right.Style &&
        Alignment.Horizontal == other.Alignment.Horizontal &&
        Alignment.Vertical == other.Alignment.Vertical &&
        WrapText == other.WrapText;
}

public sealed class FontDef
{
    public string Name { get; set; } = "Calibri";
    public double Size { get; set; } = 11;
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public bool Strikethrough { get; set; }
    public string? Color { get; set; }
    public int ThemeColor { get; set; } = -1;
    public double Tint { get; set; }
    public ExcelVerticalAlignment VerticalAlign { get; set; }
}

public enum ExcelVerticalAlignment { None, Baseline, Subscript, Superscript }

public sealed class FillDef
{
    public ExcelFillPattern PatternType { get; set; } = ExcelFillPattern.None;
    public string? BackgroundColor { get; set; }
    public string? ForegroundColor { get; set; }
    public ExcelGradientFill? Gradient { get; set; }
}

public enum ExcelFillPattern
{
    None, Solid, DarkGray, MediumGray, LightGray, Gray125, Gray0625,
    DarkHorizontal, DarkVertical, DarkDown, DarkUp, DarkGrid, DarkTrellis,
    LightHorizontal, LightVertical, LightDown, LightUp, LightGrid, LightTrellis
}

public sealed class ExcelGradientFill
{
    public double Degree { get; set; }
    public List<(double Position, string Color)> Stops { get; } = new();
}

public sealed class BorderDef
{
    public BorderSideDef Top { get; } = new();
    public BorderSideDef Bottom { get; } = new();
    public BorderSideDef Left { get; } = new();
    public BorderSideDef Right { get; } = new();
    public BorderSideDef Diagonal { get; } = new();
    public bool DiagonalUp { get; set; }
    public bool DiagonalDown { get; set; }

    public void SetAll(BorderLineStyle style, string? color = null)
    {
        Top.Style = Bottom.Style = Left.Style = Right.Style = style;
        if (color != null)
            Top.Color = Bottom.Color = Left.Color = Right.Color = color;
    }
}

public sealed class BorderSideDef
{
    public BorderLineStyle Style { get; set; } = BorderLineStyle.None;
    public string? Color { get; set; }
}

public enum BorderLineStyle
{
    None, Hair, Dotted, DashDot, Thin, SlantDashDot, MediumDashDot,
    DashDotDot, MediumDashDotDot, MediumDotted, Medium, Double, Thick
}

public sealed class AlignmentDef
{
    public ExcelHorizontalAlignment Horizontal { get; set; } = ExcelHorizontalAlignment.General;
    public ExcelVerticalCellAlignment Vertical { get; set; } = ExcelVerticalCellAlignment.Bottom;
    public bool WrapText { get; set; }
    public int Indent { get; set; }
    public int TextRotation { get; set; }
    public bool ShrinkToFit { get; set; }
    public bool JustifyLastLine { get; set; }
}

public enum ExcelHorizontalAlignment { General, Left, Center, Right, Fill, Justify, CenterContinuous, Distributed }
public enum ExcelVerticalCellAlignment { Top, Center, Bottom, Justify, Distributed }

// ── Conditional formatting ─────────────────────────────────────────────────────

public sealed class ConditionalFormattingRule
{
    public string Address { get; set; } = "";
    public ConditionalFormattingType Type { get; set; }
    public int Priority { get; set; } = 1;
    public bool StopIfTrue { get; set; }
    public string? Operator { get; set; }
    public object? Value1 { get; set; }
    public object? Value2 { get; set; }
    public string? Formula { get; set; }
    public CellStyleDef Style { get; set; } = new();
    public ColorScaleDef? ColorScale { get; set; }
    public DataBarDef? DataBar { get; set; }
    public IconSetDef? IconSet { get; set; }
    public bool AboveAverage { get; set; } = true;
    public bool EqualAverage { get; set; }
    public int StdDev { get; set; }
    public int Rank { get; set; } = 10;
    public bool Bottom { get; set; }
    public bool Percent { get; set; }
    public string? Text { get; set; }
    public int TimePeriod { get; set; }
}

public enum ConditionalFormattingType
{
    CellValue, Expression, ColorScale, DataBar, IconSet,
    Top10, AboveAverage, DuplicateValues, UniqueValues,
    ContainsText, NotContainsText, BeginsWith, EndsWith,
    ContainsBlanks, NotContainsBlanks, ContainsErrors,
    TimePeriod, BelowAverage
}

public sealed class ColorScaleDef
{
    public string MinColor { get; set; } = "FF63BE7B";
    public string MidColor { get; set; } = "FFFFEB84";
    public string MaxColor { get; set; } = "FFF8696B";
    public string MinType { get; set; } = "min";
    public string MidType { get; set; } = "percentile";
    public string MaxType { get; set; } = "max";
    public double MidValue { get; set; } = 50;
}

public sealed class DataBarDef
{
    public string Color { get; set; } = "FF638EC6";
    public bool ShowValue { get; set; } = true;
    public string MinType { get; set; } = "min";
    public string MaxType { get; set; } = "max";
}

public sealed class IconSetDef
{
    public string SetName { get; set; } = "3Arrows";
    public bool ShowValue { get; set; } = true;
    public bool Reverse { get; set; }
}

// ── Data validation ────────────────────────────────────────────────────────────

public sealed class ExcelDataValidation
{
    public string Address { get; set; } = "";
    public DataValidationType Type { get; set; }
    public DataValidationOperator Operator { get; set; } = DataValidationOperator.Between;
    public string? Formula1 { get; set; }
    public string? Formula2 { get; set; }
    public bool AllowBlank { get; set; } = true;
    public bool ShowInputMessage { get; set; }
    public bool ShowErrorAlert { get; set; }
    public string? PromptTitle { get; set; }
    public string? Prompt { get; set; }
    public string? ErrorTitle { get; set; }
    public string? Error { get; set; }
    public DataValidationAlertStyle ErrorStyle { get; set; } = DataValidationAlertStyle.Stop;
    public bool InCellDropdown { get; set; } = true;
}

public enum DataValidationType { None, Whole, Decimal, List, Date, Time, TextLength, Custom }
public enum DataValidationOperator { Between, NotBetween, Equal, NotEqual, LessThan, LessThanOrEqual, GreaterThan, GreaterThanOrEqual }
public enum DataValidationAlertStyle { Stop, Warning, Information }

// ── Custom table styles ────────────────────────────────────────────────────────

public sealed class ExcelCustomTableStyle(string name)
{
    public string Name { get; } = name;
    public bool IsTableStyle { get; set; } = true;
    public bool IsPivotStyle { get; set; }
    public List<ExcelTableStyleElement> Elements { get; } = new();

    public ExcelTableStyleElement AddElement(ExcelTableStyleElementType type)
    {
        var el = new ExcelTableStyleElement(type);
        Elements.Add(el);
        return el;
    }
}

public sealed class ExcelTableStyleElement(ExcelTableStyleElementType type)
{
    public ExcelTableStyleElementType Type { get; } = type;
    public CellStyleDef Style { get; set; } = new();
    public int Size { get; set; } = 1;
}

public enum ExcelTableStyleElementType
{
    WholeTable, HeaderRow, TotalRow, FirstColumn, LastColumn,
    FirstRowStripe, SecondRowStripe, FirstColumnStripe, SecondColumnStripe,
    HeaderRowFirstColumn, HeaderRowLastColumn, TotalRowFirstColumn,
    TotalRowLastColumn, FirstHeaderCell, LastHeaderCell,
    FirstTotalCell, LastTotalCell
}
