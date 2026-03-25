using SkiaSharp;

namespace EPExcel.ML;

// ── Chart ─────────────────────────────────────────────────────────────────────

public sealed class ExcelChart(string name, ExcelChartType type)
{
    public string Name { get; } = name;
    public ExcelChartType ChartType { get; } = type;
    public string? Title { get; set; }
    public bool ShowTitle { get; set; } = true;
    public bool ShowLegend { get; set; } = true;
    public ExcelLegendPosition LegendPosition { get; set; } = ExcelLegendPosition.Right;
    public int FromRow { get; set; } = 2;
    public int FromCol { get; set; } = 2;
    public int ToRow { get; set; } = 18;
    public int ToCol { get; set; } = 14;
    public int Style { get; set; } = 2;
    public bool RoundedCorners { get; set; }
    public List<ExcelChartSeries> Series { get; } = new();
    public ExcelChartAxis PrimaryXAxis { get; } = new();
    public ExcelChartAxis PrimaryYAxis { get; } = new();
    public ExcelChartAxis? SecondaryYAxis { get; set; }
    public bool UseSecondaryAxis { get; set; }
    public bool VaryColors { get; set; }
    public ExcelBarDirection BarDirection { get; set; } = ExcelBarDirection.Bar;
    public ExcelBarGrouping BarGrouping { get; set; } = ExcelBarGrouping.Clustered;
    public bool Smooth { get; set; }
    public bool ShowDataLabels { get; set; }
    public ExcelPieChart? PieOptions { get; set; }
    public string? BackgroundColor { get; set; }
    public int StyleId { get; set; } = 2;

    public ExcelChartSeries AddSeries(ExcelRange values, ExcelRange? categories = null)
    {
        var s = new ExcelChartSeries
        {
            ValuesAddress = values.FullAddress,
            ValuesSheet = values.Worksheet.Name,
            CategoriesAddress = categories?.FullAddress,
            CategoriesSheet = categories?.Worksheet.Name,
        };
        Series.Add(s);
        return s;
    }

    public ExcelChartSeries AddSeries(string valuesAddress, string? catAddress = null)
    {
        var s = new ExcelChartSeries { ValuesAddress = valuesAddress, CategoriesAddress = catAddress };
        Series.Add(s);
        return s;
    }

    public byte[] Render(int width = 600, int height = 400) =>
        ChartRenderer.Render(this, width, height);
}

public enum ExcelChartType
{
    ColumnClustered, ColumnStacked, ColumnStacked100,
    BarClustered, BarStacked, BarStacked100,
    LineSeries, LineMarkers, LineStackedSeries, LineMarkersStacked,
    Pie, PieExploded, Doughnut,
    Area, AreaStacked, AreaStacked100,
    XYScatter, XYScatterLines, XYScatterSmooth,
    Bubble,
    Radar, RadarFilled,
    Surface, Surface3D,
}

public enum ExcelLegendPosition { Top, Bottom, Left, Right, TopRight }
public enum ExcelBarDirection { Bar, Column }
public enum ExcelBarGrouping { Clustered, Stacked, PercentStacked, Standard }

public sealed class ExcelChartSeries
{
    public string? Name { get; set; }
    public string? ValuesAddress { get; set; }
    public string? ValuesSheet { get; set; }
    public string? CategoriesAddress { get; set; }
    public string? CategoriesSheet { get; set; }
    public string? Color { get; set; }
    public string? MarkerColor { get; set; }
    public int MarkerSize { get; set; } = 5;
    public ExcelMarkerStyle Marker { get; set; } = ExcelMarkerStyle.None;
    public bool UseSecondaryAxis { get; set; }
    public int Order { get; set; }
    public string? HeaderAddress { get; set; }
    public bool Smooth { get; set; }
    public int LineWidth { get; set; } = 2;
    public List<ExcelDataLabel> DataLabels { get; } = new();
}

public enum ExcelMarkerStyle { None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture }

public sealed class ExcelChartAxis
{
    public string? Title { get; set; }
    public bool ShowMajorGridLines { get; set; } = true;
    public bool ShowMinorGridLines { get; set; }
    public double? MinValue { get; set; }
    public double? MaxValue { get; set; }
    public double? MajorUnit { get; set; }
    public double? MinorUnit { get; set; }
    public bool LogScale { get; set; }
    public bool Crosses { get; set; }
    public ExcelAxisPosition CrossesAt { get; set; } = ExcelAxisPosition.AutoZero;
    public int TextRotation { get; set; }
    public bool Reverse { get; set; }
    public bool Hidden { get; set; }
    public string? NumberFormat { get; set; }
}

public enum ExcelAxisPosition { AutoZero, Min, Max }

public sealed class ExcelDataLabel
{
    public bool ShowValue { get; set; } = true;
    public bool ShowCategoryName { get; set; }
    public bool ShowSeriesName { get; set; }
    public bool ShowLegendKey { get; set; }
    public bool ShowPercent { get; set; }
    public string? Separator { get; set; }
}

public sealed class ExcelPieChart
{
    public int FirstSliceAngle { get; set; }
    public int Explosion { get; set; }
    public int HoleSize { get; set; } = 50;
}

// ── Images ────────────────────────────────────────────────────────────────────

public sealed class ExcelImage(string name, byte[] data)
{
    public string Name { get; } = name;
    public byte[] ImageData { get; } = data;
    public string ContentType { get; set; } = "image/png";
    public int FromRow { get; set; } = 1;
    public int FromCol { get; set; } = 1;
    public int ToRow { get; set; } = 10;
    public int ToCol { get; set; } = 6;
    public string? Description { get; set; }
    public bool IsDecorative { get; set; }
    public int Rotation { get; set; }
    public bool LockAspectRatio { get; set; } = true;
    public int OffsetX { get; set; }
    public int OffsetY { get; set; }
}

// ── Shapes ────────────────────────────────────────────────────────────────────

public sealed class ExcelShape(string name, ExcelShapeType type)
{
    public string Name { get; } = name;
    public ExcelShapeType ShapeType { get; } = type;
    public int FromRow { get; set; } = 1;
    public int FromCol { get; set; } = 1;
    public int ToRow { get; set; } = 5;
    public int ToCol { get; set; } = 4;
    public string? Text { get; set; }
    public string? FillColor { get; set; }
    public string? BorderColor { get; set; }
    public double BorderWidth { get; set; } = 1;
    public ExcelShapeAdjustmentValues AdjustmentValues { get; } = new();
    public string? Font { get; set; } = "Calibri";
    public int FontSize { get; set; } = 11;
    public bool Bold { get; set; }
    public string? TextColor { get; set; }
    public ExcelHorizontalAlignment TextAlign { get; set; } = ExcelHorizontalAlignment.Center;
    public int Rotation { get; set; }
    public bool FlipH { get; set; }
    public bool FlipV { get; set; }
    public string? HyperlinkUrl { get; set; }
}

public enum ExcelShapeType
{
    Rectangle = 0, RoundRect, Ellipse, Triangle, RtTriangle, Parallelogram,
    Trapezoid, Diamond, Pentagon, Hexagon, Heptagon, Octagon, Star4, Star5,
    Star6, Star8, Star16, Star24, Star32, SmileyFace, Donut, Heart,
    LightningBolt, Sun, Moon, Cloud, Arc, Plus,
    RightArrow, LeftArrow, UpArrow, DownArrow, LeftRightArrow, UpDownArrow,
    QuadArrow, BentArrow, CircularArrow, Chevron, NotchedRightArrow,
    FlowChartProcess, FlowChartDecision, FlowChartInputOutput,
    FlowChartTerminator, FlowChartDocument, FlowChartConnector,
    WedgeRectCallout, WedgeRoundRectCallout, WedgeEllipseCallout, CloudCallout,
    IrregularSeal1, IrregularSeal2, Wave, DoubleWave,
    Ribbon, Ribbon2, HorizontalScroll, VerticalScroll,
    TextBox = 100
}

public static class ShapeGeometryMap
{
    private static readonly Dictionary<ExcelShapeType, string> _map = new()
    {
        [ExcelShapeType.Rectangle] = "rect",
        [ExcelShapeType.RoundRect] = "roundRect",
        [ExcelShapeType.Ellipse] = "ellipse",
        [ExcelShapeType.Triangle] = "triangle",
        [ExcelShapeType.RtTriangle] = "rtTriangle",
        [ExcelShapeType.Diamond] = "diamond",
        [ExcelShapeType.Pentagon] = "pentagon",
        [ExcelShapeType.Hexagon] = "hexagon",
        [ExcelShapeType.Star4] = "star4",
        [ExcelShapeType.Star5] = "star5",
        [ExcelShapeType.Star8] = "star8",
        [ExcelShapeType.Heart] = "heart",
        [ExcelShapeType.LightningBolt] = "lightningBolt",
        [ExcelShapeType.Sun] = "sun",
        [ExcelShapeType.Moon] = "moon",
        [ExcelShapeType.Cloud] = "cloud",
        [ExcelShapeType.RightArrow] = "rightArrow",
        [ExcelShapeType.LeftArrow] = "leftArrow",
        [ExcelShapeType.UpArrow] = "upArrow",
        [ExcelShapeType.DownArrow] = "downArrow",
        [ExcelShapeType.Chevron] = "chevron",
        [ExcelShapeType.FlowChartProcess] = "flowChartProcess",
        [ExcelShapeType.FlowChartDecision] = "flowChartDecision",
        [ExcelShapeType.FlowChartTerminator] = "flowChartTerminator",
        [ExcelShapeType.WedgeRectCallout] = "wedgeRectCallout",
        [ExcelShapeType.WedgeEllipseCallout] = "wedgeEllipseCallout",
        [ExcelShapeType.TextBox] = "rect",
    };

    public static string GetName(ExcelShapeType t) =>
        _map.TryGetValue(t, out var n) ? n : "rect";
}

public sealed class ExcelShapeAdjustmentValues
{
    private readonly Dictionary<int, int> _vals = new();
    public void Set(int index, int value) => _vals[index] = value;
    public int Get(int index) => _vals.TryGetValue(index, out var v) ? v : 0;

    public string ToOoxml()
    {
        if (!_vals.Any()) return "";
        var entries = _vals.OrderBy(kv => kv.Key)
            .Select(kv => $"""<a:gd name="adj{kv.Key + 1}" fmla="val {kv.Value}"/>""");
        return $"<a:avLst>{string.Join("", entries)}</a:avLst>";
    }
}

// ── Comments ──────────────────────────────────────────────────────────────────

public sealed class ExcelComment
{
    public int Row { get; set; }
    public int Col { get; set; }
    public string Text { get; set; } = "";
    public string? Author { get; set; }
    public bool Visible { get; set; }
    public int Width { get; set; } = 144;
    public int Height { get; set; } = 79;
}

public sealed class ExcelThreadedComment
{
    public int Row { get; set; }
    public int Col { get; set; }
    public List<ExcelThreadedCommentThread> Threads { get; } = new();
}

public sealed class ExcelThreadedCommentThread
{
    public string? AuthorId { get; set; }
    public string Text { get; set; } = "";
    public DateTime Created { get; set; } = DateTime.UtcNow;
    public bool Resolved { get; set; }
    public List<ExcelCommentMention> Mentions { get; } = new();
}

public sealed class ExcelCommentMention
{
    public string? UserId { get; set; }
    public string? DisplayName { get; set; }
    public int StartIndex { get; set; }
    public int Length { get; set; }
}

// ── Sparklines ────────────────────────────────────────────────────────────────

public sealed class ExcelSparklineGroup
{
    public string DataRange { get; set; } = "";
    public string LocationRange { get; set; } = "";
    public ExcelSparklineType Type { get; set; }
    public string? HighColor { get; set; }
    public string? LowColor { get; set; }
    public string? FirstColor { get; set; }
    public string? LastColor { get; set; }
    public string? NegativeColor { get; set; }
    public string? MarkersColor { get; set; }
    public string? SeriesColor { get; set; }
    public bool ShowMarkers { get; set; }
    public bool ShowHigh { get; set; }
    public bool ShowLow { get; set; }
    public bool ShowFirst { get; set; }
    public bool ShowLast { get; set; }
    public bool ShowNegative { get; set; }
}

public enum ExcelSparklineType { Line, Column, Stacked }

// ── Slicers ───────────────────────────────────────────────────────────────────

public sealed class ExcelSlicer
{
    public string Name { get; set; } = "";
    public string? Caption { get; set; }
    public string FieldName { get; set; } = "";
    public bool IsTableSlicer { get; set; }
    public string? TableName { get; set; }
    public string? PivotTableName { get; set; }
    public int Row { get; set; }
    public int Col { get; set; }
    public int Width { get; set; } = 144;
    public int Height { get; set; } = 200;
    public int Columns { get; set; } = 1;
    public string StyleName { get; set; } = "SlicerStyleLight1";
    public List<string> SelectedValues { get; } = new();
}

// ── Tables ────────────────────────────────────────────────────────────────────

public sealed class ExcelTable(string name, string address)
{
    public string Name { get; } = name;
    public string Address { get; set; } = address;
    public bool ShowHeader { get; set; } = true;
    public bool ShowTotals { get; set; }
    public bool ShowFirstColumn { get; set; }
    public bool ShowLastColumn { get; set; }
    public bool ShowRowStripes { get; set; } = true;
    public bool ShowColumnStripes { get; set; }
    public bool ShowFilter { get; set; } = true;
    public string? StyleName { get; set; } = "TableStyleMedium9";
    public List<ExcelTableColumn> Columns { get; } = new();
    public int Id { get; set; }
    public string? DisplayName { get; set; }
}

public sealed class ExcelTableColumn
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public string? TotalsRowLabel { get; set; }
    public string? TotalsRowFunction { get; set; }
    public string? Formula { get; set; }
    public string? DataDxfId { get; set; }
    public string? HeaderDxfId { get; set; }
    public ExcelTableColumnDataType DataType { get; set; }
}

public enum ExcelTableColumnDataType { None, String, Number, Boolean, DateTime }

// ── Pivot tables ──────────────────────────────────────────────────────────────

public sealed class ExcelPivotTable
{
    public string Name { get; set; } = "";
    public string DataRange { get; set; } = "";
    public string DataSheetName { get; set; } = "";
    public string OutputRange { get; set; } = "";
    public string StyleName { get; set; } = "PivotStyleMedium9";
    public bool ShowRowHeaders { get; set; } = true;
    public bool ShowColumnHeaders { get; set; } = true;
    public bool ShowRowStripes { get; set; } = true;
    public bool ShowColumnStripes { get; set; }
    public bool GrandTotalRow { get; set; } = true;
    public bool GrandTotalCol { get; set; } = true;
    public bool Compact { get; set; } = true;
    public bool Outline { get; set; }
    public bool SubtotalTop { get; set; } = true;
    public bool RepeatItemLabels { get; set; }
    public bool InsertBlankRow { get; set; }

    public List<ExcelPivotField> Fields { get; } = new();
    public List<ExcelPivotRowField> RowFields { get; } = new();
    public List<ExcelPivotColField> ColFields { get; } = new();
    public List<ExcelPivotPageField> PageFields { get; } = new();
    public List<ExcelPivotDataField> DataFields { get; } = new();

    // Calculated data
    public bool IsCalculated { get; internal set; }
    public Dictionary<PivotDataKey, PivotAccumulator> CalculatedData { get; } = new();

    public void Calculate(bool refreshCache = true, ExcelWorkbook? workbook = null)
    {
        if (workbook == null) return;
        var engine = new IO.PivotCalculationEngine(workbook);
        engine.Calculate(this, workbook);
    }

    public double GetPivotData(string dataField,
        IEnumerable<PivotDataCriteria>? criteria = null,
        ExcelWorkbook? workbook = null)
    {
        if (!IsCalculated && workbook != null) Calculate(true, workbook);
        if (!CalculatedData.Any()) return 0;
        var crit = criteria?.ToList() ?? [];
        return IO.PivotCalculationEngine.QueryPivot(
            this, dataField,
            crit.Select(c => c.FieldName).ToArray(),
            crit.Select(c => c.Value).ToArray());
    }
}

public sealed class ExcelPivotField
{
    public string Name { get; set; } = "";
    public List<object?> Values { get; } = new();
    public List<string> UniqueValues { get; } = new();
    public bool IsNumeric { get; set; }
    public bool Compact { get; set; } = true;
    public bool Outline { get; set; }
    public bool ShowBlankItems { get; set; } = true;
    public bool InsertBlankRow { get; set; }
    public bool SubtotalTop { get; set; } = true;
}

public sealed class ExcelPivotRowField { public string FieldName { get; set; } = ""; public int DefaultSubtotal { get; set; } = 1; }
public sealed class ExcelPivotColField { public string FieldName { get; set; } = ""; }
public sealed class ExcelPivotPageField { public string FieldName { get; set; } = ""; public string? SelectedItem { get; set; } }
public sealed class ExcelPivotDataField
{
    public string FieldName { get; set; } = "";
    public string? Caption { get; set; }
    public PivotDataFunction Function { get; set; } = PivotDataFunction.Sum;
    public string? NumberFormat { get; set; }
}

public enum PivotDataFunction { Sum, Count, Average, Max, Min, Product, StdDev, StdDevP, Var, VarP, CountNums }

public sealed class PivotDataKey : IEquatable<PivotDataKey>
{
    public string DataField { get; }
    public List<string> RowValues { get; }
    public List<string> ColValues { get; }

    public PivotDataKey(string dataField, List<string> rowValues, List<string> colValues)
    {
        DataField = dataField;
        RowValues = rowValues;
        ColValues = colValues;
    }

    public bool Equals(PivotDataKey? other)
    {
        if (other == null) return false;
        return DataField.Equals(other.DataField, StringComparison.OrdinalIgnoreCase) &&
               RowValues.SequenceEqual(other.RowValues, StringComparer.OrdinalIgnoreCase) &&
               ColValues.SequenceEqual(other.ColValues, StringComparer.OrdinalIgnoreCase);
    }

    public override bool Equals(object? obj) => Equals(obj as PivotDataKey);

    public override int GetHashCode()
    {
        var hc = new HashCode();
        hc.Add(DataField, StringComparer.OrdinalIgnoreCase);
        foreach (var v in RowValues) hc.Add(v, StringComparer.OrdinalIgnoreCase);
        foreach (var v in ColValues) hc.Add(v, StringComparer.OrdinalIgnoreCase);
        return hc.ToHashCode();
    }
}

public sealed class PivotAccumulator
{
    private readonly List<double> _values = new();
    public double Result { get; private set; }
    public int Count => _values.Count;

    public void Accumulate(double value, PivotDataFunction fn)
    {
        _values.Add(value);
        Result = fn switch
        {
            PivotDataFunction.Count     => _values.Count,
            PivotDataFunction.Average   => _values.Average(),
            PivotDataFunction.Max       => _values.Max(),
            PivotDataFunction.Min       => _values.Min(),
            PivotDataFunction.Product   => _values.Aggregate(1.0, (a, v) => a * v),
            PivotDataFunction.StdDev    => _values.Count < 2 ? 0 : Math.Sqrt(_values.Sum(x => Math.Pow(x - _values.Average(), 2)) / (_values.Count - 1)),
            PivotDataFunction.StdDevP   => _values.Count < 1 ? 0 : Math.Sqrt(_values.Sum(x => Math.Pow(x - _values.Average(), 2)) / _values.Count),
            PivotDataFunction.Var       => _values.Count < 2 ? 0 : _values.Sum(x => Math.Pow(x - _values.Average(), 2)) / (_values.Count - 1),
            PivotDataFunction.VarP      => _values.Count < 1 ? 0 : _values.Sum(x => Math.Pow(x - _values.Average(), 2)) / _values.Count,
            PivotDataFunction.CountNums => _values.Count(v => !double.IsNaN(v)),
            _                           => _values.Sum(),
        };
    }
}

public sealed class PivotDataCriteria
{
    public string FieldName { get; }
    public string Value { get; }
    public PivotDataCriteria(ExcelPivotRowField f, string v) { FieldName = f.FieldName; Value = v; }
    public PivotDataCriteria(ExcelPivotColField f, string v) { FieldName = f.FieldName; Value = v; }
    public PivotDataCriteria(string f, string v) { FieldName = f; Value = v; }
}

// ── Query table ───────────────────────────────────────────────────────────────

public sealed class ExcelQueryTable(string name, int connId, string range)
{
    public string Name { get; } = name;
    public int ConnectionId { get; } = connId;
    public string Range { get; } = range;
    public bool AutoFit { get; set; } = true;
    public bool RowNumbers { get; set; }
    public bool RefreshOnLoad { get; set; } = true;
    public string? StyleName { get; set; } = "TableStyleMedium9";
}
