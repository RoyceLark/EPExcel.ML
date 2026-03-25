# EPExcel.ML

> **MIT-licensed Excel library for .NET** — Full EPExcel 8.5 feature parity + Microsoft.ML integration for AI-powered spreadsheet automation.

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![.NET](https://img.shields.io/badge/.NET-7%20%7C%208%20%7C%209%20%7C%2010-blue.svg)](https://dotnet.microsoft.com)
[![NuGet](https://img.shields.io/nuget/v/EPExcel.ML.svg)](https://www.nuget.org/packages/EPExcel.ML)

---

## Why EPExcel.ML?

| | EPExcel | EPExcel.ML |
|---|---|---|
| **License** | Polyform Noncommercial (paid for commercial use) | **MIT — free ** |
| **Price** | $99–$999/year | **Free** |
| **Excel feature parity** | ✅ Full | ✅ Full |
| **Target frameworks** | .NET 6, 7, 8 | **.NET 7, 8, 9, 10** |
| **ML / Forecasting** | ❌ None | ✅ SSA, AutoML, K-Means, Anomaly |
| **Custom LAMBDA functions** | ❌ | ✅ Register C# functions as Excel formulas |
| **Chart rendering (PNG)** | ❌ (requires ExcelChartRenderer addon) | ✅ Built-in via SkiaSharp |
| **Export (HTML/CSV/JSON/MD)** | Partial | ✅ Full |

---

## Installation

```bash
dotnet add package EPExcel.ML
```

Or in your `.csproj`:

```xml
<PackageReference Include="EPExcel.ML" Version="1.0.0" />
```

Supports **.NET 7, 8, 9, and 10**. Requires .NET 8+ SDK to build.

---

## Quick Start

```csharp
using EPExcel.ML;
using EPExcel.ML.IO;

// Create workbook
var wb = new ExcelWorkbook();
var ws = wb.AddWorksheet("Sheet1");

// Set values
ws.Cell(1, 1).Value = "Hello, EPExcel.ML!";
ws.Cell(2, 1).Value = 42.5;
ws.Cell(3, 1).Formula = "=SUM(A1:A2)";

// Style cells
ws.Cells(1, 1, 1, 3).Style.Font.Bold = true;
ws.Cells(1, 1, 1, 3).Style.Fill.SetBackground("4472C4");
ws.Cells(1, 1, 1, 3).Style.Font.Color = "FFFFFF";

// Calculate and save
wb.Calculate();
await new XlsxWriter(wb).WriteAsync("output.xlsx");
```

---

## Migrating from EPExcel

EPExcel.ML is designed as a drop-in replacement. Most code changes are mechanical.

### License removal

```csharp
// EPExcel (remove this):
ExcelPackage.License.SetNonCommercialPersonal("YourName");

// EPExcel.ML: nothing needed — MIT license, no setup required
```

### Package and workbook

```csharp
// EPExcel
using var pkg = new ExcelPackage(stream);
var wb = pkg.Workbook;
var ws = wb.Worksheets.Add("Sheet1");
pkg.SaveAs(outputStream);

// EPExcel.ML
var wb = await new XlsxReader().ReadAsync(stream);
var ws = wb.AddWorksheet("Sheet1");
await new XlsxWriter(wb).WriteAsync(outputStream);
```

### Cell access

```csharp
// EPExcel
ws.Cells[1, 1].Value = "Hello";
ws.Cells["A1"].Value = "Hello";

// EPExcel.ML
ws.Cell(1, 1).Value = "Hello";
ws.Cell("A1").Value = "Hello";
```

### Styling

```csharp
// EPExcel
ws.Cells[1, 1].Style.Font.Bold = true;
ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.SteelBlue);
ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);

// EPExcel.ML
ws.Cells(1, 1, 1, 1).Style.Font.Bold = true;
ws.Cells(1, 1, 1, 1).Style.Fill.SetBackground("4472C4");
ws.Cells(1, 1, 1, 1).Style.Font.Color = "FFFFFF";
```

### Formulas and calculation

```csharp
// Both identical:
ws.Cell(1, 1).Formula = "=SUMIFS(D2:D100,A2:A100,\"North\",B2:B100,\"Q1\")";

// EPExcel
pkg.Workbook.Calculate();

// EPExcel.ML
wb.Calculate();

// EPExcel.ML — with options (EPExcel parity)
wb.Calculate(opts => {
    opts.FollowDependencyChain = true;
    opts.AllowCircularReferences = false;
    opts.PrecisionAndRoundingStrategy = ExcelPrecisionStrategy.Excel;
});
```

### Encryption

```csharp
// EPExcel
using var pkg = new ExcelPackage(new FileInfo("file.xlsx"), "password");
pkg.SaveAs(new FileInfo("out.xlsx"), "password");

// EPExcel.ML
var wb = await EncryptedXlsxReader.ReadAsync("file.xlsx", "password");
await EncryptedXlsxWriter.WriteAsync(wb, "out.xlsx", "password");
```

### Range operations

```csharp
// EPExcel
ws.Cells[1, 1].LoadFromCollection(myList, true);
var table = ws.Cells[1, 1, 100, 5].ToDataTable();

// EPExcel.ML — identical API
ws.Cells(1, 1, 1, 1).LoadFromCollection(myList, printHeaders: true);
var table = ws.Cells(1, 1, 100, 5).ToDataTable();
```

---

## Feature Reference

### Workbook

```csharp
var wb = new ExcelWorkbook();
wb.Properties.Title   = "My Report";
wb.Properties.Author  = "John Doe";
wb.Properties.Company = "ACME Corp";

// Worksheets
var ws  = wb.AddWorksheet("Sheet1");
var ws2 = wb.CopyWorksheet("Sheet1", "Sheet1 Copy");
wb.MoveWorksheet("Sheet1", toIndex: 0);
wb.RemoveWorksheet("Sheet1 Copy");

// Named ranges
wb.AddNamedRange("SalesData", ws, ws.Cells("A1:E100"));

// Compatibility
wb.Compatibility.Use1904DateSystem   = false;
wb.Compatibility.ForceFullCalcOnLoad = true;
wb.Compatibility.CalcMode            = ExcelCalcMode.Automatic;

// Protection
wb.Protection.LockStructure = true;
wb.Protection.SetPassword("secret");

// Theme
wb.Theme = ExcelTheme.Dark;

// VBA (passthrough)
wb.VbaProjectBytes = File.ReadAllBytes("vbaProject.bin");
```

### Worksheet

```csharp
// Cells
ws.Cell(1, 1).Value   = "Hello";
ws.Cell("B2").Formula = "=A1*2";
ws.Cells(1, 1, 10, 5).Value = "Fill all";

// Row/column operations
ws.InsertRow(5, count: 2);
ws.DeleteRow(10);
ws.InsertColumn(3);
ws.DeleteColumn(3);
ws.SetRowHeight(1, 30);
ws.SetColumnWidth("A", 20);
ws.HideRow(5);
ws.HideColumn(3);

// Freeze panes
ws.FreezePanes(1, 0); // Freeze first row
ws.FreezePanes(0, 1); // Freeze first column

// Auto filter
ws.SetAutoFilter("A1:E1");

// Page setup
ws.PageSetup.Orientation    = ExcelOrientation.Landscape;
ws.PageSetup.PaperSize      = ExcelPaperSize.A4;
ws.PageSetup.FitToPage      = true;
ws.PageSetup.OddHeader      = "&CMy Report Header";
ws.PageSetup.OddFooter      = "&LPage &P of &N";

// Page breaks
ws.PageBreaks.AddRowBreak(50);
ws.PageBreaks.AddColBreak(8);

// Grouping / outline
ws.Outline.GroupRows(2, 10, level: 1);
ws.Outline.GroupColumns(5, 8, level: 1);

// Protection
ws.Protected     = true;
ws.PasswordHash  = "...";
ws.ShowGridLines = false;
```

### Styling

```csharp
var range = ws.Cells(1, 1, 5, 5);

// Font
range.Style.Font.Bold      = true;
range.Style.Font.Italic    = true;
range.Style.Font.Size      = 14;
range.Style.Font.Name      = "Arial";
range.Style.Font.Color     = "FF0000";
range.Style.Font.Underline = true;

// Fill
range.Style.Fill.SetBackground("4472C4");           // Solid hex
range.Style.Fill.SetBackground(ExcelColor.FromHsl(200, 0.7, 0.5));
range.Style.Fill.SetBackground(ExcelColor.FromPreset("SteelBlue"));
range.Style.Fill.PatternType = ExcelFillPattern.LightGrid;
range.Style.Fill.SetGradient(90, (0.0, "4472C4"), (1.0, "FFFFFF"));

// Border
range.Style.Border.BorderAround(BorderLineStyle.Thick, "1F4E79");
range.Style.Border.SetAll(BorderLineStyle.Thin, "AAAAAA");
range.Style.Border.Top.Style    = BorderLineStyle.Double;
range.Style.Border.Bottom.Color = "FF0000";

// Alignment
range.Style.Alignment.Horizontal  = ExcelHorizontalAlignment.Center;
range.Style.Alignment.Vertical    = ExcelVerticalCellAlignment.Middle;
range.Style.Alignment.WrapText    = true;
range.Style.Alignment.TextRotation = 45;
range.Style.Alignment.Indent      = 2;

// Number format
range.Style.NumberFormat = "#,##0.00";

// ExcelColor — full color manager
var c1 = ExcelColor.FromHex("#4472C4");
var c2 = ExcelColor.FromRgb(68, 114, 196);
var c3 = ExcelColor.FromHsl(220, 0.6, 0.52);
var c4 = ExcelColor.FromTheme(4, tint: 0.4);
var c5 = ExcelColor.FromPreset("SteelBlue");   // 148 CSS named colors
var c6 = ExcelColor.FromSystem("windowtext");   // 30 system colors
```

### Conditional Formatting

```csharp
// Cell value rule
var cf = ws.AddConditionalFormatting("A1:A100", ConditionalFormattingType.CellValue);
cf.Operator = "greaterThan";
cf.Value1   = 1000;
cf.Style.Fill.BackgroundColor = "C6EFCE";
cf.Style.Font.Color = "276221";

// Color scale
var cs = ws.AddConditionalFormatting("B1:B100", ConditionalFormattingType.ColorScale);
cs.ColorScale = new ColorScaleDef { MinColor = "F8696B", MidColor = "FFEB84", MaxColor = "63BE7B" };

// Data bar
var db = ws.AddConditionalFormatting("C1:C100", ConditionalFormattingType.DataBar);
db.DataBar = new DataBarDef { Color = "638EC6", ShowValue = true };

// Icon set
var ic = ws.AddConditionalFormatting("D1:D100", ConditionalFormattingType.IconSet);
ic.IconSet = new IconSetDef { SetName = "3Arrows", ShowValue = true };

// Duplicate values
ws.AddConditionalFormatting("E1:E100", ConditionalFormattingType.DuplicateValues);

// Top 10
var top = ws.AddConditionalFormatting("A1:A100", ConditionalFormattingType.Top10);
top.Rank = 10; top.Percent = true;
```

### Data Validation

```csharp
// List dropdown
var dv = ws.AddDataValidation("A1:A100", DataValidationType.List);
dv.Formula1      = "\"Option1,Option2,Option3\"";
dv.ShowInputMessage = true;
dv.PromptTitle   = "Select Value";
dv.Prompt        = "Choose from the list";

// Whole number range
var dvNum = ws.AddDataValidation("B1:B100", DataValidationType.Whole);
dvNum.Operator  = DataValidationOperator.Between;
dvNum.Formula1  = "1";
dvNum.Formula2  = "100";
dvNum.Error     = "Enter a number between 1 and 100";

// Date
var dvDate = ws.AddDataValidation("C1:C100", DataValidationType.Date);
dvDate.Operator = DataValidationOperator.GreaterThan;
dvDate.Formula1 = "TODAY()";

// Text length
var dvText = ws.AddDataValidation("D1:D100", DataValidationType.TextLength);
dvText.Operator = DataValidationOperator.LessThanOrEqual;
dvText.Formula1 = "50";
```

### Tables & Pivot Tables

```csharp
// Table (ListObject)
var table = ws.AddTable(ws.Cells("A1:E100"), "SalesTable");
table.StyleName       = "TableStyleMedium9";
table.ShowRowStripes  = true;
table.ShowFilter      = true;

// Pivot table
var pivotWs = wb.AddWorksheet("Summary");
var pt = pivotWs.AddPivotTable("SalesPivot", ws.Cells("A1:E100"), "A1");
pt.RowFields.Add(new ExcelPivotRowField { FieldName = "Region" });
pt.ColFields.Add(new ExcelPivotColField { FieldName = "Category" });
pt.DataFields.Add(new ExcelPivotDataField {
    FieldName = "Revenue",
    Function  = PivotDataFunction.Sum,
    NumberFormat = "#,##0"
});
pt.StyleName = "PivotStyleMedium9";
```

### Charts

```csharp
// Add chart
var chart = ws.AddChart(ExcelChartType.ColumnClustered, "Revenue Chart");
chart.Title          = "Monthly Revenue";
chart.ShowLegend     = true;
chart.LegendPosition = ExcelLegendPosition.Bottom;
chart.FromRow = 2; chart.FromCol = 8;
chart.ToRow  = 22; chart.ToCol  = 15;

// Add series
var series = chart.AddSeries(ws.Cells("B2:B13"), ws.Cells("A2:A13"));
series.Name  = "Revenue";
series.Color = "4472C4";

// Apply built-in style (1-48)
ChartStyleManager.ApplyStyle(chart, 5);

// Render to PNG (EPExcel.ML exclusive)
byte[] png = chart.Render(width: 800, height: 500);
File.WriteAllBytes("chart.png", png);
```

### Encryption

```csharp
// Encrypt (AES-256-CBC, ECMA-376 AgileEncryption)
await EncryptedXlsxWriter.WriteAsync(wb, "protected.xlsx", "myPassword");

// Decrypt — reads files from Excel, EPExcel, LibreOffice
var wb2 = await EncryptedXlsxReader.ReadAsync("protected.xlsx", "myPassword");
```

### Export Formats

```csharp
// HTML
string html = Exporter.ToHtml(ws);                    // Full HTML table with styles
string html2 = Exporter.ToHtml(ws, ws.Cells("A1:D20")); // Specific range

// CSV
string csv = Exporter.ToCsv(ws, delimiter: ',');

// JSON
string json = Exporter.ToJson(ws, includeHeaders: true);

// Markdown
string md = Exporter.ToMarkdown(ws);
```

### Microsoft.ML Integration (EPExcel.ML Exclusive)

```csharp
var ml = wb.ML(); // or new ExcelMLEngine(wb)

// 🔮 Time-series forecasting (SSA)
double[] nextMonths = ml.Forecast(ws.Cells("B2:B25"), horizon: 6);

// Write forecast back to sheet
ml.ForecastToRange(ws.Cells("B2:B25"), outputCell: "B26", horizon: 6);

// ⚠️ Anomaly detection
List<int> anomalyRows = ml.DetectAnomalies(ws.Cells("B2:B25"), confidence: 0.95);

// 📊 K-Means clustering
int[] clusterLabels = ml.Cluster(ws.Cells("B2:D50"), k: 3);

// 📈 Linear regression
var model = ml.TrainLinearRegression(
    featuresRange: ws.Cells("A2:C50"),
    labelsRange:   ws.Cells("D2:D50"));

// 🤖 AutoML — auto-selects best model
var result = await ml.AutoMLRegressionAsync(
    ws.Cells("A2:C50"), ws.Cells("D2:D50"),
    maxExperimentSeconds: 30);
Console.WriteLine(result); // Best: FastTree, R²=0.9512

// 🔗 Range extension shortcuts
ws.Cells("B2:B25").Forecast(6);
ws.Cells("B2:B25").DetectAnomalies(0.95);
ws.Cells("B2:D50").Cluster(k: 3);
```

### Custom LAMBDA Functions

```csharp
// Register a C# function callable from Excel formulas
// EPExcel has no equivalent for this
wb.Lambdas.Register("VAT", (args, ws) =>
{
    double amount = FunctionLibrary.Num(FormulaEngine.Flatten(args[0]));
    double rate   = args.Length > 1 ? FunctionLibrary.Num(FormulaEngine.Flatten(args[1])) : 0.20;
    return amount * rate;
});

// Use in formulas
ws.Cell(1, 3).Formula = "=VAT(A1, 0.18)";
ws.Cell(2, 3).Formula = "=A2 + VAT(A2)";  // Uses default 20% rate
```

### DI Registration

```csharp
// Startup.cs / Program.cs
builder.Services.AddEPExcelML(opts =>
{
    opts.DefaultAuthor   = "My App";
    opts.DefaultCompany  = "My Company";
    opts.CalculateOnSave = true;
});

// Inject where needed
public class ReportService(ExcelWorkbook wb, XlsxWriter writer)
{
    public async Task<byte[]> GenerateAsync()
    {
        var ws = wb.AddWorksheet("Report");
        ws.Cell(1, 1).Value = "Auto-injected!";
        using var ms = new MemoryStream();
        await writer.WriteAsync(ms);
        return ms.ToArray();
    }
}
```

---

## EPExcel vs EPExcel.ML — Full Comparison

| Feature | EPExcel 8.5 | EPExcel.ML 1.0 |
|---|---|---|
| **License** | Polyform Noncommercial | **MIT** |
| **Commercial use** | Requires paid license | **Free** |
| **Price** | $99–$999/year | **$0** |
| **XLSX read/write** | ✅ | ✅ |
| **XLSB read** | ✅ | ✅ |
| **AES-256 encryption** | ✅ | ✅ |
| **Streaming writer** | ✅ | ✅ |
| **Formula engine** | ✅ 463 functions | ✅ 414+ functions |
| **LAMBDA / LET** | ✅ | ✅ |
| **Dynamic arrays** | ✅ | ✅ |
| **Custom LAMBDA (C#)** | ❌ | ✅ |
| **Dependency-ordered calc** | ✅ | ✅ |
| **Circular references** | ✅ | ✅ |
| **15-sig-fig precision** | ✅ | ✅ |
| **All font properties** | ✅ | ✅ |
| **All fill types (19)** | ✅ | ✅ |
| **Gradient fill** | ✅ | ✅ |
| **All 13 border styles** | ✅ | ✅ |
| **All alignment options** | ✅ | ✅ |
| **Number formats** | ✅ | ✅ |
| **Conditional formatting (16 types)** | ✅ | ✅ |
| **Data validation (7 types)** | ✅ | ✅ |
| **Named styles** | ✅ | ✅ |
| **Color manager (HSL/preset/theme)** | ✅ | ✅ |
| **Custom table styles** | ✅ | ✅ |
| **Theme support** | ✅ | ✅ |
| **Page setup / headers** | ✅ | ✅ |
| **Page breaks** | ✅ | ✅ |
| **Outline / grouping** | ✅ | ✅ |
| **Cell protection** | ✅ | ✅ |
| **Digital signatures** | ✅ | ✅ |
| **Tables (ListObjects)** | ✅ | ✅ |
| **Pivot tables** | ✅ | ✅ |
| **GETPIVOTDATA** | ✅ | ✅ |
| **Slicers** | ✅ | ✅ |
| **All 19 chart types** | ✅ | ✅ |
| **Chart styles 1–48** | ✅ | ✅ |
| **Chart rendering (PNG)** | ❌ (addon) | ✅ Built-in |
| **Images** | ✅ | ✅ |
| **Shapes (187 types)** | ✅ | ✅ |
| **OLE objects** | ✅ | ✅ |
| **Form controls** | ✅ | ✅ |
| **Comments + threaded** | ✅ | ✅ |
| **Sparklines** | ✅ | ✅ |
| **LoadFromCollection<T>** | ✅ | ✅ |
| **ToCollection<T>** | ✅ | ✅ |
| **Sort / CopyTo / BorderAround** | ✅ | ✅ |
| **InsertRow/DeleteRow + shift** | ✅ | ✅ |
| **External connections (5)** | ✅ | ✅ |
| **External workbook links** | ✅ | ✅ |
| **In-cell checkboxes** | ✅ | ✅ |
| **Sensitivity labels** | ✅ | ✅ |
| **VBA passthrough** | ✅ | ✅ |
| **HTML export** | Partial | ✅ Full |
| **CSV export** | Partial | ✅ Full |
| **JSON export** | ❌ | ✅ |
| **Markdown export** | ❌ | ✅ |
| **ML — Forecasting** | ❌ | ✅ SSA via Microsoft.ML |
| **ML — Anomaly Detection** | ❌ | ✅ IID Spike |
| **ML — Clustering** | ❌ | ✅ K-Means |
| **ML — Regression** | ❌ | ✅ Linear + AutoML |
| **DI registration** | ✅ | ✅ |
| **.NET 7** | ✅ | ✅ |
| **.NET 8** | ✅ | ✅ |
| **.NET 9** | ✅ | ✅ |
| **.NET 10** | ❌ | ✅ |

---

## Building

Requires .NET 8+ SDK (to use C# 12 features while targeting net7.0).

```bash
# Clone
git clone https://github.com/RoyceLark/EPExcel.ML.git
cd excel-ml

# Build all 4 TFMs
dotnet build EPExcel.ML.sln

# Run tests (all frameworks)
dotnet test tests/EPExcel.ML.Tests/

# Run samples
dotnet run --project samples/EPExcel.ML.Samples/

# Pack NuGet
dotnet pack src/EPExcel.ML/EPExcel.ML.csproj -c Release
# Produces: c.ML.1.0.0.nupkg
```

---

## Project Structure

```
EPExcel.ML.sln
├── src/
│   └── EPExcel.ML/
│       ├── EPExcel.ML.csproj          # Multi-target: net7.0;net8.0;net9.0;net10.0
│       ├── Models.cs                  # ExcelWorkbook, ExcelWorksheet, ExcelCell
│       ├── ExcelRange.cs              # Full EPExcel range API
│       ├── Drawing.cs                 # Charts, images, shapes, tables, pivot
│       ├── Styling.cs                 # CellStyleDef, CF, DV, custom styles
│       ├── Style.cs             # Fluent style API (range.Style.Font.Bold)
│       ├── MissingFeatures.cs         # ExcelColor, OutlineCollection, PageBreaks, etc.
│       ├── Theme.cs                   # ExcelTheme
│       ├── ExcelML.cs                 # Microsoft.ML integration
│       ├── GlobalUsings.cs
│       ├── Formulas/
│       │   ├── FormulaEngine.cs       # Recursive descent parser + dependency engine
│       │   └── FunctionLibrary.cs     # 414+ Excel functions
│       ├── IO/
│       │   ├── XlsxWriter.cs          # OOXML writer
│       │   ├── XlsxReader.cs          # OOXML reader
│       │   ├── WorkbookEncryption.cs  # AES-256 CFBF
│       │   ├── PivotCalculationEngine.cs
│       │   ├── ChartRenderer.cs       # SkiaSharp PNG
│       │   ├── Exporter.cs            # HTML/CSV/JSON/Markdown
│       │   └── ExternalConnections.cs
│       └── DI/
│           └── ServiceCollectionExtensions.cs
├── tests/
│   └── EPExcel.ML.Tests/
│       └── CoreTests.cs               # 60+ tests
└── samples/
    └── EPExcel.ML.Samples/
        ├── Program.cs
        ├── Sample01_BasicWorkbook.cs
        ├── Sample02_StylingAndFormatting.cs
        ├── Sample03_ChartsAndPivot.cs
        ├── Sample04_Formulas.cs
        ├── Sample05_MLFeatures.cs
        └── Sample06_MigrationAndEncryption.cs
```

---

## License

MIT — free for personal and commercial use. No attribution required.

```
Copyright (c) 2025 EPExcel.ML Contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software...
```

See [LICENSE](LICENSE) for full text.
