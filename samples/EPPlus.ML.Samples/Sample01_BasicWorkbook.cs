using EPExcel.ML;
using EPExcel.ML.IO;

namespace EPExcel.ML.Samples;

/// <summary>
/// Sample 01 — Basic workbook creation.
/// EPExcel migration: replace new ExcelPackage() with new ExcelWorkbook()
/// </summary>
public static class Sample01_BasicWorkbook
{
    public static async Task RunAsync()
    {
        Console.WriteLine("\n=== Sample 01: Basic Workbook ===");

        var wb = new ExcelWorkbook();
        wb.Properties.Title   = "EPExcel.ML Demo";
        wb.Properties.Author  = "EPExcel.ML";
        wb.Properties.Company = "My Company";

        var ws = wb.AddWorksheet("Sales Report");

        // ── Headers ───────────────────────────────────────────────────────────
        string[] headers = ["Month", "Revenue", "Expenses", "Profit", "Margin %"];
        for (int c = 0; c < headers.Length; c++)
        {
            var cell = ws.Cell(1, c + 1);
            cell.Value = headers[c];
            // EPExcel: ws.Cells[1,c+1].Style.Font.Bold = true
            ws.Cells(1, c + 1, 1, c + 1).Style.Font.Bold = true;
            ws.Cells(1, c + 1, 1, c + 1).Style.Fill.SetBackground("4472C4");
            ws.Cells(1, c + 1, 1, c + 1).Style.Font.Color = "FFFFFF";
        }

        // ── Data rows ─────────────────────────────────────────────────────────
        string[] months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        var rng = new Random(42);
        for (int i = 0; i < 12; i++)
        {
            int row = i + 2;
            double rev  = Math.Round(rng.NextDouble() * 50000 + 30000, 2);
            double exp  = Math.Round(rng.NextDouble() * 30000 + 15000, 2);

            ws.Cell(row, 1).Value = months[i];
            ws.Cell(row, 2).Value = rev;
            ws.Cell(row, 3).Value = exp;
            ws.Cell(row, 4).Formula = $"=B{row}-C{row}";         // Profit
            ws.Cell(row, 5).Formula = $"=IF(B{row}>0,D{row}/B{row},0)"; // Margin

            // Format currency
            ws.Cell(row, 2).NumberFormat = "#,##0.00";
            ws.Cell(row, 3).NumberFormat = "#,##0.00";
            ws.Cell(row, 4).NumberFormat = "#,##0.00";
            ws.Cell(row, 5).NumberFormat = "0.00%";

            // Alternate row shading
            if (i % 2 == 0)
                ws.Cells(row, 1, row, 5).Style.Fill.SetBackground("EEF2FF");
        }

        // ── Totals row ────────────────────────────────────────────────────────
        int totalRow = 14;
        ws.Cell(totalRow, 1).Value = "TOTAL";
        ws.Cell(totalRow, 2).Formula = "=SUM(B2:B13)";
        ws.Cell(totalRow, 3).Formula = "=SUM(C2:C13)";
        ws.Cell(totalRow, 4).Formula = "=SUM(D2:D13)";
        ws.Cell(totalRow, 5).Formula = "=IF(B14>0,D14/B14,0)";

        ws.Cells(totalRow, 1, totalRow, 5).Style.Font.Bold = true;
        ws.Cells(totalRow, 1, totalRow, 5).Style.Fill.SetBackground("1F4E79");
        ws.Cells(totalRow, 1, totalRow, 5).Style.Font.Color = "FFFFFF";
        for (int c = 2; c <= 5; c++)
            ws.Cell(totalRow, c).NumberFormat = c == 5 ? "0.00%" : "#,##0.00";

        // ── Auto-size columns ─────────────────────────────────────────────────
        ws.SetColumnWidth(1, 12);
        ws.SetColumnWidth(2, 14);
        ws.SetColumnWidth(3, 14);
        ws.SetColumnWidth(4, 14);
        ws.SetColumnWidth(5, 12);

        // ── Calculate formulas ────────────────────────────────────────────────
        wb.Calculate();

        // ── Save ──────────────────────────────────────────────────────────────
        await using var ms = new MemoryStream();
        await new XlsxWriter(wb).WriteAsync(ms);
        await File.WriteAllBytesAsync("Sample01_BasicWorkbook.xlsx", ms.ToArray());

        Console.WriteLine($"  Created: Sample01_BasicWorkbook.xlsx");
        Console.WriteLine($"  Sheets: {wb.Worksheets.Count}");
        Console.WriteLine($"  Rows: {ws.MaxRow}");
    }
}
