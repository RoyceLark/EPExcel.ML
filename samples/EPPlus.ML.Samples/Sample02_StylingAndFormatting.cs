using EPExcel.ML;
using EPExcel.ML.IO;

namespace EPExcel.ML.Samples;

/// <summary>
/// Sample 02 — Styling, conditional formatting, data validation.
/// </summary>
public static class Sample02_StylingAndFormatting
{
    public static async Task RunAsync()
    {
        Console.WriteLine("\n=== Sample 02: Styling & Formatting ===");

        var wb = new ExcelWorkbook();
        var ws = wb.AddWorksheet("Styled Report");

        // ── Font styling ──────────────────────────────────────────────────────
        ws.Cell(1, 1).Value = "EPExcel.ML Style Showcase";
        ws.Cells(1, 1, 1, 6).Style.Font.Bold  = true;
        ws.Cells(1, 1, 1, 6).Style.Font.Size  = 16;
        ws.Cells(1, 1, 1, 6).Style.Font.Color = "1F4E79";
        ws.SetRowHeight(1, 28);

        // ── Color fills ───────────────────────────────────────────────────────
        string[] colors  = ["4472C4", "ED7D31", "A9D18E", "FFC000", "5B9BD5", "70AD47"];
        string[] labels  = ["Blue", "Orange", "Green", "Gold", "SkyBlue", "Lime"];
        for (int i = 0; i < 6; i++)
        {
            ws.Cell(3, i + 1).Value = labels[i];
            ws.Cells(3, i + 1, 3, i + 1).Style.Fill.SetBackground(colors[i]);
            ws.Cells(3, i + 1, 3, i + 1).Style.Font.Color = "FFFFFF";
            ws.Cells(3, i + 1, 3, i + 1).Style.Font.Bold  = true;
            ws.Cells(3, i + 1, 3, i + 1).Style.Alignment.Horizontal =
                ExcelHorizontalAlignment.Center;
        }

        // ── HSL color via ExcelColor ──────────────────────────────────────────
        ws.Cell(4, 1).Value = "HSL Color";
        var hslColor = ExcelColor.FromHsl(120, 0.6, 0.4);
        ws.Cells(4, 1, 4, 1).Style.Fill.SetBackground(hslColor);
        ws.Cells(4, 1, 4, 1).Style.Font.Color = "FFFFFF";

        ws.Cell(4, 2).Value = "Preset: Coral";
        ws.Cells(4, 2, 4, 2).Style.Fill.SetBackground(ExcelColor.FromPreset("Coral"));

        ws.Cell(4, 3).Value = "Theme Color";
        ws.Cells(4, 3, 4, 3).Style.Fill.SetBackground(ExcelColor.FromTheme(4, 0.4));

        // ── Borders ───────────────────────────────────────────────────────────
        ws.Cell(6, 1).Value = "Thick border";
        ws.Cells(6, 1, 8, 3).Style.Border.BorderAround(BorderLineStyle.Thick, "1F4E79");
        ws.Cells(6, 1, 8, 3).Style.Border.SetAll(BorderLineStyle.Thin, "AAAAAA");

        ws.Cell(7, 2).Value = "Inner cells";
        ws.Cell(8, 1).Value = "Double bottom";
        ws.Cells(8, 1, 8, 3).Style.Border.Bottom.Style = BorderLineStyle.Double;

        // ── Gradient fill ─────────────────────────────────────────────────────
        ws.Cell(10, 1).Value = "Gradient Fill";
        ws.Cells(10, 1, 10, 3).Style.Fill.SetGradient(90,
            (0.0, "4472C4"), (1.0, "FFFFFF"));

        // ── Alignment ────────────────────────────────────────────────────────
        ws.Cell(12, 1).Value = "Left Align";
        ws.Cells(12, 1, 12, 1).Style.Alignment.Horizontal = ExcelHorizontalAlignment.Left;

        ws.Cell(12, 2).Value = "Centered";
        ws.Cells(12, 2, 12, 2).Style.Alignment.Horizontal = ExcelHorizontalAlignment.Center;

        ws.Cell(12, 3).Value = "Right Align";
        ws.Cells(12, 3, 12, 3).Style.Alignment.Horizontal = ExcelHorizontalAlignment.Right;

        ws.Cell(12, 4).Value = "Rotated 45°";
        ws.Cells(12, 4, 12, 4).Style.Alignment.TextRotation = 45;
        ws.SetRowHeight(12, 40);

        ws.Cell(12, 5).Value = "Wrapped long text in a cell";
        ws.Cells(12, 5, 12, 5).Style.Alignment.WrapText = true;
        ws.SetColumnWidth(5, 14);

        // ── Conditional formatting ────────────────────────────────────────────
        // Score data
        ws.Cell(14, 1).Value = "Student";
        ws.Cell(14, 2).Value = "Score";
        ws.Cells(14, 1, 14, 2).Style.Font.Bold = true;

        string[] students = ["Alice", "Bob", "Charlie", "Diana", "Eve", "Frank"];
        double[] scores   = [92, 45, 78, 61, 88, 34];
        for (int i = 0; i < students.Length; i++)
        {
            ws.Cell(15 + i, 1).Value = students[i];
            ws.Cell(15 + i, 2).Value = scores[i];
        }

        // Red for <50, green for >=80 (EPExcel parity)
        var cfRed = ws.AddConditionalFormatting("B15:B20", ConditionalFormattingType.CellValue);
        cfRed.Operator = "lessThan";
        cfRed.Value1   = 50;
        cfRed.Style.Fill.BackgroundColor = "FFC7CE";
        cfRed.Style.Font.Color = "9C0006";

        var cfGreen = ws.AddConditionalFormatting("B15:B20", ConditionalFormattingType.CellValue);
        cfGreen.Operator = "greaterThanOrEqual";
        cfGreen.Value1   = 80;
        cfGreen.Style.Fill.BackgroundColor = "C6EFCE";
        cfGreen.Style.Font.Color = "276221";

        // Data bar on scores
        var cfBar = ws.AddConditionalFormatting("B15:B20", ConditionalFormattingType.DataBar);
        cfBar.DataBar = new DataBarDef { Color = "638EC6", ShowValue = true };
        cfBar.Priority = 3;

        // ── Data validation ───────────────────────────────────────────────────
        var dvGrade = ws.AddDataValidation("C15:C20", DataValidationType.List);
        dvGrade.Formula1       = "\"A,B,C,D,F\"";
        dvGrade.ShowInputMessage = true;
        dvGrade.PromptTitle    = "Enter Grade";
        dvGrade.Prompt         = "Select a grade from the dropdown";
        dvGrade.ShowErrorAlert = true;
        dvGrade.ErrorTitle     = "Invalid Grade";
        dvGrade.Error          = "Please select A, B, C, D, or F";

        ws.Cell(14, 3).Value = "Grade";
        ws.Cells(14, 3, 14, 3).Style.Font.Bold = true;

        var dvScore = ws.AddDataValidation("B15:B20", DataValidationType.Whole);
        dvScore.Operator  = DataValidationOperator.Between;
        dvScore.Formula1  = "0";
        dvScore.Formula2  = "100";
        dvScore.AllowBlank = false;

        // ── Page setup ────────────────────────────────────────────────────────
        ws.PageSetup.Orientation    = ExcelOrientation.Landscape;
        ws.PageSetup.PaperSize      = ExcelPaperSize.A4;
        ws.PageSetup.FitToPage      = true;
        ws.PageSetup.FitToWidth     = 1;
        ws.PageSetup.OddHeader      = "&C&\"Calibri,Bold\"&14EPExcel.ML Style Report";
        ws.PageSetup.OddFooter      = "&LPage &P of &N&R&D";
        ws.PageSetup.TopMargin      = 1.0;
        ws.PageSetup.BottomMargin   = 1.0;

        // Save
        await using var ms = new MemoryStream();
        await new XlsxWriter(wb).WriteAsync(ms);
        await File.WriteAllBytesAsync("Sample02_Styling.xlsx", ms.ToArray());
        Console.WriteLine("  Created: Sample02_Styling.xlsx");
    }
}


