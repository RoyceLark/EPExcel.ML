using EPExcel.ML;
using EPExcel.ML.IO;
using EPExcel.ML.Formulas;

namespace EPExcel.ML.Samples;

/// <summary>
/// Sample 04 — Formula engine: SUMIFS, XLOOKUP, dynamic arrays, custom LAMBDA.
/// </summary>
public static class Sample04_Formulas
{
    public static async Task RunAsync()
    {
        Console.WriteLine("\n=== Sample 04: Formulas ===");

        var wb = new ExcelWorkbook();
        var ws = wb.AddWorksheet("Formulas");

        // ── Source data ───────────────────────────────────────────────────────
        ws.Cell(1, 1).Value = "Region";
        ws.Cell(1, 2).Value = "Category";
        ws.Cell(1, 3).Value = "Month";
        ws.Cell(1, 4).Value = "Amount";

        var data = new[]
        {
            ("North", "Hardware", "Jan", 15000),
            ("North", "Software", "Jan", 8000),
            ("South", "Hardware", "Jan", 12000),
            ("South", "Software", "Feb", 9500),
            ("East",  "Hardware", "Feb", 18000),
            ("East",  "Software", "Feb", 7000),
            ("North", "Hardware", "Mar", 22000),
            ("West",  "Software", "Mar", 6500),
        };

        for (int i = 0; i < data.Length; i++)
        {
            var (region, cat, month, amount) = data[i];
            ws.Cell(i + 2, 1).Value = region;
            ws.Cell(i + 2, 2).Value = cat;
            ws.Cell(i + 2, 3).Value = month;
            ws.Cell(i + 2, 4).Value = (double)amount;
        }

        // ── SUMIFS ────────────────────────────────────────────────────────────
        ws.Cell(12, 1).Value = "SUMIFS examples:";
        ws.Cells(12, 1, 12, 1).Style.Font.Bold = true;

        ws.Cell(13, 1).Value = "North Hardware total:";
        ws.Cell(13, 2).Formula = "=SUMIFS(D2:D9,A2:A9,\"North\",B2:B9,\"Hardware\")";
        ws.Cell(13, 2).NumberFormat = "#,##0";

        ws.Cell(14, 1).Value = "All Software (any region):";
        ws.Cell(14, 2).Formula = "=SUMIFS(D2:D9,B2:B9,\"Software\")";
        ws.Cell(14, 2).NumberFormat = "#,##0";

        ws.Cell(15, 1).Value = "Feb Hardware:";
        ws.Cell(15, 2).Formula = "=SUMIFS(D2:D9,C2:C9,\"Feb\",B2:B9,\"Hardware\")";
        ws.Cell(15, 2).NumberFormat = "#,##0";

        // ── COUNTIFS ──────────────────────────────────────────────────────────
        ws.Cell(17, 1).Value = "COUNTIFS examples:";
        ws.Cells(17, 1, 17, 1).Style.Font.Bold = true;

        ws.Cell(18, 1).Value = "Count North rows:";
        ws.Cell(18, 2).Formula = "=COUNTIFS(A2:A9,\"North\")";

        ws.Cell(19, 1).Value = "Count Hardware > 15000:";
        ws.Cell(19, 2).Formula = "=COUNTIFS(B2:B9,\"Hardware\",D2:D9,\">15000\")";

        // ── XLOOKUP ───────────────────────────────────────────────────────────
        ws.Cell(21, 1).Value = "XLOOKUP:";
        ws.Cells(21, 1, 21, 1).Style.Font.Bold = true;

        // Lookup table
        ws.Cell(22, 4).Value = "Code";
        ws.Cell(22, 5).Value = "Description";
        ws.Cell(23, 4).Value = "HW";
        ws.Cell(23, 5).Value = "Hardware Products";
        ws.Cell(24, 4).Value = "SW";
        ws.Cell(24, 5).Value = "Software Products";
        ws.Cell(25, 4).Value = "SVC";
        ws.Cell(25, 5).Value = "Services";

        ws.Cell(22, 1).Value = "Lookup 'SW':";
        ws.Cell(22, 2).Formula = "=XLOOKUP(\"SW\",D23:D25,E23:E25,\"Not Found\")";

        ws.Cell(23, 1).Value = "Lookup 'HW':";
        ws.Cell(23, 2).Formula = "=XLOOKUP(\"HW\",D23:D25,E23:E25,\"Not Found\")";

        // ── Statistical functions ─────────────────────────────────────────────
        ws.Cell(25, 1).Value = "Statistics:";
        ws.Cells(25, 1, 25, 1).Style.Font.Bold = true;

        ws.Cell(26, 1).Value = "Average amount:";
        ws.Cell(26, 2).Formula = "=AVERAGE(D2:D9)";
        ws.Cell(26, 2).NumberFormat = "#,##0.00";

        ws.Cell(27, 1).Value = "Std Dev:";
        ws.Cell(27, 2).Formula = "=STDEV(D2:D9)";
        ws.Cell(27, 2).NumberFormat = "#,##0.00";

        ws.Cell(28, 1).Value = "NORM.INV(0.95):";
        ws.Cell(28, 2).Formula = "=NORM.INV(0.95,AVERAGE(D2:D9),STDEV(D2:D9))";
        ws.Cell(28, 2).NumberFormat = "#,##0.00";

        ws.Cell(29, 1).Value = "LINEST slope:";
        ws.Cell(29, 2).Formula = "=INDEX(LINEST(D2:D9,ROW(D2:D9)-1,TRUE,TRUE),1,1)";
        ws.Cell(29, 2).NumberFormat = "0.00";

        // ── DATE functions ────────────────────────────────────────────────────
        ws.Cell(31, 1).Value = "Date functions:";
        ws.Cells(31, 1, 31, 1).Style.Font.Bold = true;

        ws.Cell(32, 1).Value = "Today:";
        ws.Cell(32, 2).Formula = "=TODAY()";
        ws.Cell(32, 2).NumberFormat = "dd-mmm-yyyy";

        ws.Cell(33, 1).Value = "Days until year end:";
        ws.Cell(33, 2).Formula = "=DATE(YEAR(TODAY()),12,31)-TODAY()";

        ws.Cell(34, 1).Value = "WORKDAY +30:";
        ws.Cell(34, 2).Formula = "=WORKDAY(TODAY(),30)";
        ws.Cell(34, 2).NumberFormat = "dd-mmm-yyyy";

        // ── Custom LAMBDA function ────────────────────────────────────────────
        // EPExcel doesn't support custom LAMBDA registration — EPExcel.ML exclusive!
        wb.Lambdas.Register("TAX", (args, sheet) =>
        {
            double amount = FunctionLibrary.Num(FormulaEngine.Flatten(args.Length > 0 ? args[0] : null));
            double rate   = args.Length > 1 ? FunctionLibrary.Num(FormulaEngine.Flatten(args[1])) : 0.18;
            return amount * rate;
        });

        ws.Cell(36, 1).Value = "Custom LAMBDA (TAX function):";
        ws.Cells(36, 1, 36, 1).Style.Font.Bold = true;

        ws.Cell(37, 1).Value = "Tax on 50000 @ 18%:";
        ws.Cell(37, 2).Formula = "=TAX(50000, 0.18)";
        ws.Cell(37, 2).NumberFormat = "#,##0.00";

        ws.Cell(38, 1).Value = "Tax on North Hardware:";
        ws.Cell(38, 2).Formula = "=TAX(B13, 0.18)";
        ws.Cell(38, 2).NumberFormat = "#,##0.00";

        // ── Calculate all formulas ─────────────────────────────────────────────
        wb.Calculate();

        ws.SetColumnWidth(1, 26);
        ws.SetColumnWidth(2, 20);

        await using var ms = new MemoryStream();
        await new XlsxWriter(wb).WriteAsync(ms);
        await File.WriteAllBytesAsync("Sample04_Formulas.xlsx", ms.ToArray());
        Console.WriteLine("  Created: Sample04_Formulas.xlsx");

        // Print a few calculated values
        Console.WriteLine($"  North Hardware SUMIFS: {ws.Cell(13, 2).CalculatedValue}");
        Console.WriteLine($"  XLOOKUP 'SW': {ws.Cell(22, 2).CalculatedValue}");
        Console.WriteLine($"  Custom TAX(50000): {ws.Cell(37, 2).CalculatedValue}");
    }
}

// Note: use ws.Cells(row,col,row,col).Style for single-cell styling
// The Style() shortcut is available via ExcelRange, not ExcelCell directly.
