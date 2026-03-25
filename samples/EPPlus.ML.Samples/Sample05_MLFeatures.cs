using EPExcel.ML;
using EPExcel.ML.IO;

namespace EPExcel.ML.Samples;

/// <summary>
/// Sample 05 — Microsoft.ML integration: forecasting, anomaly detection, clustering.
/// This is EXCLUSIVE to EPExcel.ML — EPExcel has no ML capabilities.
/// </summary>
public static class Sample05_MLFeatures
{
    public static async Task RunAsync()
    {
        Console.WriteLine("\n=== Sample 05: ML Features (EPExcel.ML Exclusive) ===");

        var wb = new ExcelWorkbook();
        var ws = wb.AddWorksheet("ML Analysis");

        // ── Time-series data ──────────────────────────────────────────────────
        ws.Cell(1, 1).Value = "Month";
        ws.Cell(1, 2).Value = "Actual Sales";
        ws.Cell(1, 3).Value = "Forecast";
        ws.Cell(1, 4).Value = "Is Anomaly";

        ws.Cells(1, 1, 1, 4).Style.Font.Bold = true;
        ws.Cells(1, 1, 1, 4).Style.Fill.SetBackground("1F4E79");
        ws.Cells(1, 1, 1, 4).Style.Font.Color = "FFFFFF";

        // Simulated 24-month sales data with seasonality + trend + anomaly
        double[] salesData = [
            12000, 11500, 13200, 14100, 15800, 18200,   // Year 1 H1
            16500, 15200, 14800, 13900, 12100, 14500,   // Year 1 H2
            13200, 12800, 14500, 15600, 17200, 19800,   // Year 2 H1
            85000, // Anomaly — spike
            16900, 15800, 14200, 13500, 15800            // Year 2 H2 (partial)
        ];

        string[] monthLabels = [
            "Jan-23","Feb-23","Mar-23","Apr-23","May-23","Jun-23",
            "Jul-23","Aug-23","Sep-23","Oct-23","Nov-23","Dec-23",
            "Jan-24","Feb-24","Mar-24","Apr-24","May-24","Jun-24",
            "Jul-24","Aug-24","Sep-24","Oct-24","Nov-24","Dec-24"
        ];

        for (int i = 0; i < salesData.Length; i++)
        {
            ws.Cell(i + 2, 1).Value = monthLabels[i];
            ws.Cell(i + 2, 2).Value = salesData[i];
            ws.Cell(i + 2, 2).NumberFormat = "#,##0";
        }

        // ── Forecasting via ML ─────────────────────────────────────────────────
        Console.WriteLine("  Running ML time-series forecast...");
        var mlEngine = wb.ML();
        var dataRange = ws.Cells(2, 2, salesData.Length + 1, 2);

        try
        {
            // Forecast next 6 months
            var forecasts = mlEngine.Forecast(dataRange, horizon: 6, windowSize: 6);

            for (int i = 0; i < forecasts.Length; i++)
            {
                int row = salesData.Length + 2 + i;
                ws.Cell(row, 1).Value = $"Forecast M{i + 1}";
                ws.Cell(row, 3).Value = Math.Round(forecasts[i], 0);
                ws.Cell(row, 3).NumberFormat = "#,##0";
                ws.Cells(row, 1, row, 4).Style.Fill.SetBackground("E2EFDA");
            }
            Console.WriteLine($"  Forecast: {forecasts.Length} months predicted");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Forecast note: {ex.Message}");
        }

        // ── Anomaly detection ─────────────────────────────────────────────────
        Console.WriteLine("  Running anomaly detection...");
        try
        {
            var anomalies = mlEngine.DetectAnomalies(dataRange, confidence: 0.90);

            foreach (int anomalyRow in anomalies)
            {
                int excelRow = anomalyRow + 1; // 1-based
                if (excelRow >= 2 && excelRow <= salesData.Length + 1)
                {
                    ws.Cell(excelRow, 4).Value = "⚠ ANOMALY";
                    ws.Cells(excelRow, 1, excelRow, 4).Style.Fill.SetBackground("FFC7CE");
                    ws.Cells(excelRow, 4, excelRow, 4).Style.Font.Color = "9C0006";
                    ws.Cells(excelRow, 4, excelRow, 4).Style.Font.Bold = true;
                    Console.WriteLine($"  Anomaly detected at row {excelRow}: {salesData[anomalyRow - 1]:N0}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Anomaly note: {ex.Message}");
        }

        // ── Clustering ────────────────────────────────────────────────────────
        var clusterWs = wb.AddWorksheet("Customer Clusters");

        clusterWs.Cell(1, 1).Value = "Customer";
        clusterWs.Cell(1, 2).Value = "Revenue";
        clusterWs.Cell(1, 3).Value = "Orders";
        clusterWs.Cell(1, 4).Value = "Avg Order";
        clusterWs.Cell(1, 5).Value = "Cluster";
        clusterWs.Cells(1, 1, 1, 5).Style.Font.Bold = true;
        clusterWs.Cells(1, 1, 1, 5).Style.Fill.SetBackground("4472C4");
        clusterWs.Cells(1, 1, 1, 5).Style.Font.Color = "FFFFFF";

        // Customer data
        var customers = new[]
        {
            ("ACME Corp",      150000.0, 45.0, 3333.0),
            ("BetaCo",          12000.0,  8.0, 1500.0),
            ("GammaTech",      280000.0, 92.0, 3043.0),
            ("Delta LLC",        5000.0,  3.0, 1667.0),
            ("Epsilon Inc",    320000.0,110.0, 2909.0),
            ("Zeta Corp",        8500.0,  6.0, 1417.0),
            ("Eta Systems",    175000.0, 58.0, 3017.0),
            ("Theta Ltd",        3200.0,  2.0, 1600.0),
            ("Iota Partners",  220000.0, 75.0, 2933.0),
            ("Kappa Group",      6800.0,  4.0, 1700.0),
        };

        for (int i = 0; i < customers.Length; i++)
        {
            var (name, rev, orders, avg) = customers[i];
            clusterWs.Cell(i + 2, 1).Value = name;
            clusterWs.Cell(i + 2, 2).Value = rev;
            clusterWs.Cell(i + 2, 3).Value = orders;
            clusterWs.Cell(i + 2, 4).Value = avg;
            clusterWs.Cell(i + 2, 2).NumberFormat = "#,##0";
            clusterWs.Cell(i + 2, 4).NumberFormat = "#,##0";
        }

        Console.WriteLine("  Running K-Means clustering (k=3)...");
        try
        {
            var featureRange = clusterWs.Cells(2, 2, customers.Length + 1, 4);
            var labels = mlEngine.Cluster(featureRange, k: 3);

            string[] clusterColors = ["4472C4", "ED7D31", "A9D18E"];
            string[] clusterNames  = ["High Value", "Medium Value", "Low Value"];

            for (int i = 0; i < labels.Length && i < customers.Length; i++)
            {
                int clusterIdx = Math.Clamp(labels[i] - 1, 0, 2);
                clusterWs.Cell(i + 2, 5).Value = clusterNames[clusterIdx];
                clusterWs.Cells(i + 2, 5, i + 2, 5).Style.Fill
                    .SetBackground(clusterColors[clusterIdx]);
                clusterWs.Cells(i + 2, 5, i + 2, 5).Style.Font.Color = "FFFFFF";
            }
            Console.WriteLine($"  Clustered {labels.Length} customers into 3 segments");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Cluster note: {ex.Message}");
        }

        // ── ML summary sheet ──────────────────────────────────────────────────
        var summaryWs = wb.AddWorksheet("ML Summary");
        summaryWs.Cell(1, 1).Value = "EPExcel.ML — Machine Learning Capabilities";
        summaryWs.Cells(1, 1, 1, 3).Style.Font.Bold = true;
        summaryWs.Cells(1, 1, 1, 3).Style.Font.Size = 14;
        summaryWs.Cells(1, 1, 1, 3).Style.Font.Color = "1F4E79";

        var features = new[]
        {
            ("🔮 Time-Series Forecasting", "SSA algorithm via Microsoft.ML.TimeSeries", "wb.ML().Forecast(range, horizon: 6)"),
            ("⚠️ Anomaly Detection",       "IID Spike Detection — finds outliers",       "wb.ML().DetectAnomalies(range, 0.95)"),
            ("📊 K-Means Clustering",       "Segment customers/data into N groups",       "wb.ML().Cluster(range, k: 3)"),
            ("📈 Linear Regression",        "Train on worksheet data, predict values",    "wb.ML().TrainLinearRegression(X, Y)"),
            ("🤖 AutoML Regression",        "Auto-selects best model in N seconds",       "await wb.ML().AutoMLRegressionAsync(X, Y)"),
            ("🔗 Range Extensions",         "Direct ML on any range object",              "range.Forecast(6) / range.Cluster(3)"),
        };

        summaryWs.Cell(3, 1).Value = "Feature";
        summaryWs.Cell(3, 2).Value = "Description";
        summaryWs.Cell(3, 3).Value = "API Example";
        summaryWs.Cells(3, 1, 3, 3).Style.Font.Bold = true;
        summaryWs.Cells(3, 1, 3, 3).Style.Fill.SetBackground("4472C4");
        summaryWs.Cells(3, 1, 3, 3).Style.Font.Color = "FFFFFF";

        for (int i = 0; i < features.Length; i++)
        {
            var (feat, desc, api) = features[i];
            summaryWs.Cell(i + 4, 1).Value = feat;
            summaryWs.Cell(i + 4, 2).Value = desc;
            summaryWs.Cell(i + 4, 3).Value = api;
            summaryWs.Cells(i + 4, 3, i + 4, 3).Style.Font.Name = "Consolas";
            if (i % 2 == 0)
                summaryWs.Cells(i + 4, 1, i + 4, 3).Style.Fill.SetBackground("EEF2FF");
        }

        summaryWs.SetColumnWidth(1, 28);
        summaryWs.SetColumnWidth(2, 38);
        summaryWs.SetColumnWidth(3, 44);

        await using var ms = new MemoryStream();
        await new XlsxWriter(wb).WriteAsync(ms);
        await File.WriteAllBytesAsync("Sample05_MLFeatures.xlsx", ms.ToArray());
        Console.WriteLine("  Created: Sample05_MLFeatures.xlsx");
    }
}
