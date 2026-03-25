using SkiaSharp;

namespace EPExcel.ML;

public static class ChartRenderer
{
    private static readonly string[] DefaultColors = [
        "#4472C4","#ED7D31","#A9D18E","#FFC000","#5B9BD5",
        "#70AD47","#255E91","#9E480E","#636363","#997300"
    ];

    public static byte[] Render(ExcelChart chart, int width = 600, int height = 400)
    {
        var info = new SKImageInfo(width, height, SKColorType.Rgba8888, SKAlphaType.Premul);
        using var surface = SKSurface.Create(info);
        var canvas = surface.Canvas;
        canvas.Clear(SKColors.White);
        DrawChart(canvas, chart, new SKRect(10, 10, width - 10, height - 10));
        using var image = surface.Snapshot();
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        return data.ToArray();
    }

    private static void DrawChart(SKCanvas canvas, ExcelChart chart, SKRect rect)
    {
        float titleH = 0;
        if (!string.IsNullOrEmpty(chart.Title) && chart.ShowTitle)
        {
            using var tp = new SKPaint { Color = SKColors.Black, TextSize = 16, IsAntialias = true, FakeBoldText = true, TextAlign = SKTextAlign.Center };
            canvas.DrawText(chart.Title, rect.MidX, rect.Top + 20, tp);
            titleH = 30;
        }

        float legendW = chart.ShowLegend && chart.Series.Any() ? 100 : 0;
        var plot = new SKRect(rect.Left + 50, rect.Top + titleH + 10, rect.Right - legendW - 10, rect.Bottom - 40);

        using var axPaint = new SKPaint { Color = SKColors.Gray, StrokeWidth = 1, IsAntialias = true };
        canvas.DrawLine(plot.Left, plot.Top, plot.Left, plot.Bottom, axPaint);
        canvas.DrawLine(plot.Left, plot.Bottom, plot.Right, plot.Bottom, axPaint);

        if (!chart.Series.Any())
        {
            using var np = new SKPaint { Color = SKColors.Gray, TextSize = 14, IsAntialias = true, TextAlign = SKTextAlign.Center };
            canvas.DrawText("No Data", plot.MidX, plot.MidY, np);
            return;
        }

        switch (chart.ChartType)
        {
            case ExcelChartType.Pie or ExcelChartType.PieExploded or ExcelChartType.Doughnut:
                DrawPie(canvas, chart, plot); break;
            case ExcelChartType.LineSeries or ExcelChartType.LineMarkers:
                DrawLine(canvas, chart, plot); break;
            default:
                DrawColumn(canvas, chart, plot); break;
        }

        if (chart.ShowLegend && legendW > 0)
            DrawLegend(canvas, chart, new SKRect(plot.Right + 5, plot.Top, rect.Right - 5, plot.Bottom));
    }

    private static void DrawColumn(SKCanvas canvas, ExcelChart chart, SKRect rect)
    {
        int sc = chart.Series.Count, maxPts = chart.Series.Max(s => s.DataLabels.Count > 0 ? s.DataLabels.Count : 5);
        if (maxPts == 0) maxPts = 5;
        float bgw = rect.Width / maxPts, bw = bgw / (sc + 1) * 0.8f;
        for (int si = 0; si < sc; si++)
        {
            var vals = Enumerable.Range(1, maxPts).Select(i => (double)i).ToList();
            double mx = vals.Max(); if (mx == 0) mx = 1;
            using var p = new SKPaint { Color = ParseColor(chart.Series[si].Color ?? DefaultColors[si % DefaultColors.Length]), IsAntialias = true };
            for (int i = 0; i < vals.Count; i++)
            {
                float x = rect.Left + i * bgw + si * bw + bw * 0.1f;
                float bh = (float)(vals[i] / mx * rect.Height);
                canvas.DrawRect(x, rect.Bottom - bh, bw, bh, p);
            }
        }
    }

    private static void DrawLine(SKCanvas canvas, ExcelChart chart, SKRect rect)
    {
        int maxPts = 5;
        for (int si = 0; si < chart.Series.Count; si++)
        {
            var vals = Enumerable.Range(1, maxPts).Select(i => (double)i).ToList();
            double mx = vals.Max(); if (mx == 0) mx = 1;
            using var lp = new SKPaint { Color = ParseColor(chart.Series[si].Color ?? DefaultColors[si % DefaultColors.Length]), StrokeWidth = 2, IsAntialias = true, IsStroke = true };
            var pts = vals.Select((v, i) => new SKPoint(rect.Left + i * rect.Width / (maxPts - 1), rect.Bottom - (float)(v / mx * rect.Height))).ToArray();
            for (int i = 0; i < pts.Length - 1; i++) canvas.DrawLine(pts[i], pts[i + 1], lp);
        }
    }

    private static void DrawPie(SKCanvas canvas, ExcelChart chart, SKRect rect)
    {
        var vals = Enumerable.Range(1, 4).Select(i => (double)i).ToList();
        double total = vals.Sum(); float cx = rect.MidX, cy = rect.MidY, r = Math.Min(rect.Width, rect.Height) / 2 - 10;
        float startAngle = -90;
        for (int i = 0; i < vals.Count; i++)
        {
            float sweep = (float)(vals[i] / total * 360);
            using var sp = new SKPaint { Color = ParseColor(DefaultColors[i % DefaultColors.Length]), IsAntialias = true };
            var bounds = new SKRect(cx - r, cy - r, cx + r, cy + r);
            using var path = new SKPath();
            path.MoveTo(cx, cy); path.ArcTo(bounds, startAngle, sweep, false); path.Close();
            canvas.DrawPath(path, sp);
            startAngle += sweep;
        }
        if (chart.ChartType == ExcelChartType.Doughnut)
        {
            using var hp = new SKPaint { Color = SKColors.White, IsAntialias = true };
            canvas.DrawCircle(cx, cy, r * 0.5f, hp);
        }
    }

    private static void DrawLegend(SKCanvas canvas, ExcelChart chart, SKRect rect)
    {
        float y = rect.Top + 10;
        using var tp = new SKPaint { Color = SKColors.Black, TextSize = 11, IsAntialias = true };
        for (int i = 0; i < chart.Series.Count && y < rect.Bottom; i++)
        {
            using var sp = new SKPaint { Color = ParseColor(chart.Series[i].Color ?? DefaultColors[i % DefaultColors.Length]), IsAntialias = true };
            canvas.DrawRect(rect.Left, y, 12, 12, sp);
            canvas.DrawText(chart.Series[i].Name ?? $"Series {i + 1}", rect.Left + 16, y + 10, tp);
            y += 18;
        }
    }

    private static SKColor ParseColor(string hex)
    {
        try
        {
            hex = hex.TrimStart('#');
            if (hex.Length == 6) return new SKColor(Convert.ToByte(hex[0..2], 16), Convert.ToByte(hex[2..4], 16), Convert.ToByte(hex[4..6], 16));
            if (hex.Length == 8) return new SKColor(Convert.ToByte(hex[2..4], 16), Convert.ToByte(hex[4..6], 16), Convert.ToByte(hex[6..8], 16), Convert.ToByte(hex[0..2], 16));
        }
        catch { }
        return SKColors.Blue;
    }
}
