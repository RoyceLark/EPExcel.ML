using EPExcel.ML;
using EPExcel.ML.IO;

namespace EPExcel.ML.Samples;

/// <summary>
/// Sample 06 — EPExcel migration guide + encryption + export.
/// Shows direct EPExcel → EPExcel.ML API mapping.
/// </summary>
public static class Sample06_MigrationAndEncryption
{
    public static async Task RunAsync()
    {
        Console.WriteLine("\n=== Sample 06: Migration & Encryption ===");

        // ── EPExcel → EPExcel.ML migration ──────────────────────────────────────
        Console.WriteLine("  EPExcel.ML migration patterns:");

        // EPExcel: ExcelPackage.License.SetNonCommercialPersonal("Name");
        // EPExcel.ML: No license needed — MIT, free for all use.

        // EPExcel: using var pkg = new ExcelPackage(stream);
        // EPExcel.ML: var wb = await new XlsxReader().ReadAsync(stream);

        // EPExcel: var ws = pkg.Workbook.Worksheets.Add("Sheet1");
        // EPExcel.ML:
        var wb = new ExcelWorkbook();
        var ws = wb.AddWorksheet("Migration Demo");

        // EPExcel: ws.Cells[1,1].Value = "Hello";
        // EPExcel.ML:
        ws.Cell(1, 1).Value = "EPExcel → EPExcel.ML Migration Demo";

        // EPExcel: ws.Cells[1,1].Style.Font.Bold = true;
        // EPExcel.ML:
        ws.Cells(1, 1, 1, 4).Style.Font.Bold = true;
        ws.Cells(1, 1, 1, 4).Style.Font.Size = 14;

        // EPExcel: ws.Cells[1,1].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //         ws.Cells[1,1].Style.Fill.BackgroundColor.SetColor(Color.SteelBlue);
        // EPExcel.ML:
        ws.Cells(1, 1, 1, 4).Style.Fill.SetBackground("4472C4");
        ws.Cells(1, 1, 1, 4).Style.Font.Color = "FFFFFF";

        // ── Migration comparison table ─────────────────────────────────────────
        ws.Cell(3, 1).Value = "EPExcel API";
        ws.Cell(3, 2).Value = "EPExcel.ML Equivalent";
        ws.Cell(3, 3).Value = "Notes";
        ws.Cells(3, 1, 3, 3).Style.Font.Bold = true;
        ws.Cells(3, 1, 3, 3).Style.Fill.SetBackground("1F4E79");
        ws.Cells(3, 1, 3, 3).Style.Font.Color = "FFFFFF";

        var migrations = new[]
        {
            ("new ExcelPackage(stream)",
             "await new XlsxReader().ReadAsync(stream)",
             "Async by default"),
            ("new ExcelPackage(file, \"pwd\")",
             "await EncryptedXlsxReader.ReadAsync(path, \"pwd\")",
             "AES-256 ECMA-376"),
            ("pkg.SaveAs(stream)",
             "await new XlsxWriter(wb).WriteAsync(stream)",
             "Async by default"),
            ("pkg.SaveAs(file, \"pwd\")",
             "await EncryptedXlsxWriter.WriteAsync(wb, path, \"pwd\")",
             "AES-256 ECMA-376"),
            ("pkg.Workbook.Worksheets.Add(\"S\")",
             "wb.AddWorksheet(\"S\")",
             "Same API"),
            ("ws.Cells[r,c].Value = v",
             "ws.Cell(r,c).Value = v",
             "Cell() not Cells[]"),
            ("ws.Cells[r,c].Style.Font.Bold",
             "ws.Cells(r,c,r,c).Style.Font.Bold",
             "Style instead of Style"),
            ("ws.Cells[addr].Style.Fill.BackgroundColor.SetColor()",
             "ws.Cells(addr).Style.Fill.SetBackground(\"hex\")",
             "Hex string instead of Color"),
            ("ws.Cells[r1,c1,r2,c2].LoadFromCollection(list)",
             "ws.Cells(r1,c1,r2,c2).LoadFromCollection(list)",
             "Same API"),
            ("ws.Cells[addr].Formula = \"=SUM(A1:A10)\"",
             "ws.Cell(r,c).Formula = \"=SUM(A1:A10)\"",
             "Same formula syntax"),
            ("pkg.Workbook.Calculate()",
             "wb.Calculate()",
             "With dependency ordering"),
            ("ExcelPackage.License.SetNonCommercialPersonal()",
             "// No license needed — MIT",
             "Free for commercial use"),
        };

        for (int i = 0; i < migrations.Length; i++)
        {
            var (old, @new, note) = migrations[i];
            int row = i + 4;
            ws.Cell(row, 1).Value = old;
            ws.Cell(row, 2).Value = @new;
            ws.Cell(row, 3).Value = note;
            if (i % 2 == 0)
                ws.Cells(row, 1, row, 3).Style.Fill.SetBackground("F0F4FF");
        }

        ws.SetColumnWidth(1, 40);
        ws.SetColumnWidth(2, 45);
        ws.SetColumnWidth(3, 30);

        // ── Read existing XLSX ─────────────────────────────────────────────────
        Console.WriteLine("  Testing read/write round-trip...");

        // Write
        byte[] xlsxBytes;
        using (var ms = new MemoryStream())
        {
            await new XlsxWriter(wb).WriteAsync(ms);
            xlsxBytes = ms.ToArray();
        }

        // Read back
        ExcelWorkbook wb2;
        using (var ms2 = new MemoryStream(xlsxBytes))
            wb2 = await new XlsxReader().ReadAsync(ms2);

        Console.WriteLine($"  Round-trip: {wb2.Worksheets.Count} sheets, " +
                          $"cell(1,1)='{wb2.GetWorksheet("Migration Demo")?.GetCell(1,1)?.GetString()}'");

        // ── Encrypted write ────────────────────────────────────────────────────
        Console.WriteLine("  Writing encrypted workbook...");
        var wbSecret = new ExcelWorkbook();
        var wsSecret = wbSecret.AddWorksheet("Confidential");
        wsSecret.Cell(1, 1).Value = "This file is encrypted with AES-256";
        wsSecret.Cell(2, 1).Value = "Password: demo123";
        wsSecret.Cell(3, 1).Value = "Encrypted at: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm");

        await EncryptedXlsxWriter.WriteAsync(wbSecret, "Sample06_Encrypted.xlsx", "demo123");
        Console.WriteLine("  Created: Sample06_Encrypted.xlsx (password: demo123)");

        // ── Export formats ────────────────────────────────────────────────────
        Console.WriteLine("  Testing export formats...");

        var exportWb = new ExcelWorkbook();
        var exportWs = exportWb.AddWorksheet("Data");
        exportWs.Cell(1, 1).Value = "Name";
        exportWs.Cell(1, 2).Value = "Dept";
        exportWs.Cell(1, 3).Value = "Salary";
        var people = new[] { ("Alice", "Eng", 95000), ("Bob", "Sales", 72000), ("Carol", "HR", 68000) };
        for (int i = 0; i < people.Length; i++)
        {
            exportWs.Cell(i + 2, 1).Value = people[i].Item1;
            exportWs.Cell(i + 2, 2).Value = people[i].Item2;
            exportWs.Cell(i + 2, 3).Value = (double)people[i].Item3;
        }

        string csv  = Exporter.ToCsv(exportWs);
        string html = Exporter.ToHtml(exportWs);
        string json = Exporter.ToJson(exportWs);
        string md   = Exporter.ToMarkdown(exportWs);

        await File.WriteAllTextAsync("Sample06_Export.csv",  csv);
        await File.WriteAllTextAsync("Sample06_Export.html", html);
        await File.WriteAllTextAsync("Sample06_Export.json", json);
        await File.WriteAllTextAsync("Sample06_Export.md",   md);

        Console.WriteLine($"  CSV:  {csv.Split('\n').Length - 1} rows");
        Console.WriteLine($"  HTML: {html.Length} chars");
        Console.WriteLine($"  JSON: {json.Length} chars");
        Console.WriteLine($"  MD:   {md.Split('\n').Length} lines");

        // Main migration file
        await using var finalMs = new MemoryStream();
        await new XlsxWriter(wb).WriteAsync(finalMs);
        await File.WriteAllBytesAsync("Sample06_Migration.xlsx", finalMs.ToArray());
        Console.WriteLine("  Created: Sample06_Migration.xlsx");
    }
}
