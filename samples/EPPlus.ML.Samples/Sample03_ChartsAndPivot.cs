using EPExcel.ML;
using EPExcel.ML.IO;

namespace EPExcel.ML.Samples;

/// <summary>
/// Sample 03 — Charts, pivot tables, sparklines.
/// </summary>
public static class Sample03_ChartsAndPivot
{
    public static async Task RunAsync()
    {
        Console.WriteLine("\n=== Sample 03: Charts & Pivot Tables ===");

        var wb = new ExcelWorkbook();

        // ── Source data ───────────────────────────────────────────────────────
        var data = wb.AddWorksheet("Sales Data");

        string[] headers   = ["Region", "Product", "Quarter", "Revenue", "Units"];
        string[] regions   = ["North", "South", "East", "West"];
        string[] products  = ["Widget A", "Widget B", "Widget C"];
        string[] quarters  = ["Q1", "Q2", "Q3", "Q4"];

        for (int c = 0; c < headers.Length; c++)
        {
            data.Cell(1, c + 1).Value = headers[c];
            data.Cells(1, c + 1, 1, c + 1).Style.Font.Bold = true;
            data.Cells(1, c + 1, 1, c + 1).Style.Fill.SetBackground("4472C4");
            data.Cells(1, c + 1, 1, c + 1).Style.Font.Color = "FFFFFF";
        }

        var rng  = new Random(99);
        int row  = 2;
        foreach (var region in regions)
        foreach (var product in products)
        foreach (var quarter in quarters)
        {
            data.Cell(row, 1).Value = region;
            data.Cell(row, 2).Value = product;
            data.Cell(row, 3).Value = quarter;
            data.Cell(row, 4).Value = Math.Round(rng.NextDouble() * 40000 + 10000, 0);
            data.Cell(row, 5).Value = (int)(rng.NextDouble() * 400 + 50);
            row++;
        }

        // ── Column chart sheet ────────────────────────────────────────────────
        var chartWs = wb.AddWorksheet("Revenue Chart");

        // Summary data for chart
        chartWs.Cell(1, 1).Value = "Quarter";
        chartWs.Cell(1, 2).Value = "North";
        chartWs.Cell(1, 3).Value = "South";
        chartWs.Cell(1, 4).Value = "East";
        chartWs.Cell(1, 5).Value = "West";

        for (int q = 0; q < 4; q++)
        {
            chartWs.Cell(q + 2, 1).Value = quarters[q];
            foreach (var (r, ri) in regions.Select((r, i) => (r, i)))
            {
                double total = Enumerable.Range(0, 3)
                    .Sum(_ => Math.Round(rng.NextDouble() * 40000 + 10000, 0));
                chartWs.Cell(q + 2, ri + 2).Value = total;
                chartWs.Cell(q + 2, ri + 2).NumberFormat = "#,##0";
            }
        }

        // Create column chart (EPExcel parity: ws.Drawings.AddChart(...))
        var chart = chartWs.AddChart(ExcelChartType.ColumnClustered, "Revenue by Region");
        chart.Title = "Quarterly Revenue by Region";
        chart.ShowTitle = true;
        chart.ShowLegend = true;
        chart.LegendPosition = ExcelLegendPosition.Bottom;
        chart.FromRow = 8; chart.FromCol = 1;
        chart.ToRow   = 28; chart.ToCol  = 9;

        foreach (var (r, ri) in regions.Select((r, i) => (r, i)))
        {
            var series = chart.AddSeries(
                $"'Revenue Chart'!$B$2:$E$5",
                $"'Revenue Chart'!$A$2:$A$5");
            series.Name  = r;
            series.Color = new[] { "4472C4", "ED7D31", "A9D18E", "FFC000" }[ri];
        }

        ChartStyleManager.ApplyStyle(chart, 2);

        // ── Line chart ────────────────────────────────────────────────────────
        var lineChart = chartWs.AddChart(ExcelChartType.LineMarkers, "Trend Chart");
        lineChart.Title   = "Revenue Trend";
        lineChart.FromRow = 8; lineChart.FromCol = 11;
        lineChart.ToRow   = 28; lineChart.ToCol  = 19;
        lineChart.Smooth  = true;

        var lineSeries = lineChart.AddSeries("'Revenue Chart'!$B$2:$B$5",
                                              "'Revenue Chart'!$A$2:$A$5");
        lineSeries.Name       = "North Trend";
        lineSeries.Color      = "4472C4";
        lineSeries.Marker     = ExcelMarkerStyle.Circle;
        lineSeries.LineWidth  = 3;

        // ── Pivot table ───────────────────────────────────────────────────────
        var pivotWs = wb.AddWorksheet("Pivot Summary");
        int lastDataRow = row - 1;

        var pt = pivotWs.AddPivotTable("RevenuePivot",
            data.Cells(1, 1, lastDataRow, 5), "A1");
        pt.Name             = "Revenue by Region & Product";
        pt.RowFields.Add(new ExcelPivotRowField { FieldName = "Region" });
        pt.ColFields.Add(new ExcelPivotColField { FieldName = "Product" });
        pt.DataFields.Add(new ExcelPivotDataField
        {
            FieldName    = "Revenue",
            Caption      = "Total Revenue",
            Function     = PivotDataFunction.Sum,
            NumberFormat = "#,##0"
        });
        pt.DataFields.Add(new ExcelPivotDataField
        {
            FieldName    = "Units",
            Caption      = "Total Units",
            Function     = PivotDataFunction.Sum,
            NumberFormat = "#,##0"
        });
        pt.StyleName    = "PivotStyleMedium9";
        pt.GrandTotalRow = true;
        pt.GrandTotalCol = true;

        // Calculate pivot
        var pivotEngine = new PivotCalculationEngine(wb);
        pivotEngine.CalculateAll();

        // ── Sparklines ────────────────────────────────────────────────────────
        var sparkWs = wb.AddWorksheet("Sparklines");

        sparkWs.Cell(1, 1).Value = "Product";
        sparkWs.Cell(1, 2).Value = "Jan";
        sparkWs.Cell(1, 3).Value = "Feb";
        sparkWs.Cell(1, 4).Value = "Mar";
        sparkWs.Cell(1, 5).Value = "Apr";
        sparkWs.Cell(1, 6).Value = "May";
        sparkWs.Cell(1, 7).Value = "Jun";
        sparkWs.Cell(1, 8).Value = "Trend";

        string[] sparkProducts = ["Widget A", "Widget B", "Widget C"];
        for (int p = 0; p < 3; p++)
        {
            sparkWs.Cell(p + 2, 1).Value = sparkProducts[p];
            for (int m = 0; m < 6; m++)
                sparkWs.Cell(p + 2, m + 2).Value = Math.Round(rng.NextDouble() * 1000 + 500, 0);
        }

        sparkWs.SparklineGroups.Add(new ExcelSparklineGroup
        {
            Type          = ExcelSparklineType.Line,
            DataRange     = "B2:G4",
            LocationRange = "H2:H4",
            SeriesColor   = "4472C4",
            HighColor     = "00B050",
            LowColor      = "FF0000",
            ShowHigh      = true,
            ShowLow       = true,
            ShowMarkers   = true,
        });

        // ── Save ──────────────────────────────────────────────────────────────
        await using var ms = new MemoryStream();
        await new XlsxWriter(wb).WriteAsync(ms);
        await File.WriteAllBytesAsync("Sample03_ChartsAndPivot.xlsx", ms.ToArray());
        Console.WriteLine("  Created: Sample03_ChartsAndPivot.xlsx");
        Console.WriteLine($"  Sheets: {wb.Worksheets.Count}");
        Console.WriteLine($"  Data rows: {lastDataRow - 1}");
    }
}
