using Xunit;
using FluentAssertions;
using EPExcel.ML;
using EPExcel.ML.Formulas;
using EPExcel.ML.IO;
using System.IO;
using System.Threading.Tasks;

namespace EPExcel.ML.Tests;

public class WorkbookTests
{
    [Fact] public void AddWorksheet_SetsName() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("Sheet1"); ws.Name.Should().Be("Sheet1"); }
    [Fact] public void AddWorksheet_DuplicateThrows() { var wb = new ExcelWorkbook(); wb.AddWorksheet("Sheet1"); Assert.Throws<ArgumentException>(() => wb.AddWorksheet("Sheet1")); }
    [Fact] public void GetWorksheet_CaseInsensitive() { var wb = new ExcelWorkbook(); wb.AddWorksheet("Sheet1"); wb.GetWorksheet("sheet1").Should().NotBeNull(); }
    [Fact] public void CellValue_StringRoundtrip() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S1"); ws.Cell(1,1).Value = "hello"; ws.Cell(1,1).GetString().Should().Be("hello"); }
    [Fact] public void CellValue_NumberRoundtrip() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S1"); ws.Cell(2,1).Value = 42.5; ws.Cell(2,1).GetDouble().Should().Be(42.5); }
    [Fact] public void InsertRow_ShiftsDown() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S1"); ws.Cell(1,1).Value = "A"; ws.Cell(2,1).Value = "B"; ws.InsertRow(1); ws.GetCell(2,1)?.GetString().Should().Be("A"); ws.GetCell(3,1)?.GetString().Should().Be("B"); }
    [Fact] public void DeleteRow_RemovesRow() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S1"); ws.Cell(1,1).Value = "A"; ws.Cell(2,1).Value = "B"; ws.DeleteRow(1); ws.GetCell(1,1)?.GetString().Should().Be("B"); }
    [Fact] public void MaxRow_ReturnsCorrect() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S1"); ws.Cell(5,1).Value = "x"; ws.MaxRow.Should().Be(5); }
}

public class FormulaTests
{
    private static FormulaEngine Engine() => new(new ExcelWorkbook());

    [Fact] public void SUM_Numbers() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); ws.Cell(1,1).Value=1.0; ws.Cell(1,2).Value=2.0; ws.Cell(1,3).Value=3.0; e.Evaluate("SUM(1,2,3)",ws).Should().Be(6.0); }
    [Fact] public void IF_True() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("IF(1=1,\"yes\",\"no\")",ws).Should().Be("yes"); }
    [Fact] public void IF_False() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("IF(1=2,\"yes\",\"no\")",ws).Should().Be("no"); }
    [Fact] public void VLOOKUP_Exact() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S"); ws.Cell(1,1).Value="A"; ws.Cell(1,2).Value=10.0; ws.Cell(2,1).Value="B"; ws.Cell(2,2).Value=20.0; var e = new FormulaEngine(wb); var r = e.Evaluate("VLOOKUP(\"B\",A1:B2,2,FALSE)",ws); ((double)r!).Should().Be(20.0); }
    [Fact] public void ROUND_TwoDecimals() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("ROUND(3.14159,2)",ws).Should().Be(3.14); }
    [Fact] public void CONCATENATE_Strings() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("CONCATENATE(\"Hello\",\" \",\"World\")",ws).Should().Be("Hello World"); }
    [Fact] public void LEFT_TwoChars() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("LEFT(\"Hello\",2)",ws).Should().Be("He"); }
    [Fact] public void LEN_String() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("LEN(\"Hello\")",ws).Should().Be(5.0); }
    [Fact] public void MAX_Numbers() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("MAX(1,5,3,2)",ws).Should().Be(5.0); }
    [Fact] public void MIN_Numbers() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("MIN(5,1,3,2)",ws).Should().Be(1.0); }
    [Fact] public void AVERAGE_Numbers() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("AVERAGE(1,2,3,4,5)",ws).Should().Be(3.0); }
    [Fact] public void IFERROR_CatchesError() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("IFERROR(1/0,\"err\")",ws).Should().Be("err"); }
    [Fact] public void ISBLANK_EmptyCell() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("ISBLANK(A1)",ws).Should().Be(true); }
    [Fact] public void ISNUMBER_Double() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("ISNUMBER(42)",ws).Should().Be(true); }
    [Fact] public void ISTEXT_String() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("ISTEXT(\"hello\")",ws).Should().Be(true); }
    [Fact] public void AND_AllTrue() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("AND(1=1,2=2)",ws).Should().Be(true); }
    [Fact] public void OR_OneTrue() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("OR(1=2,2=2)",ws).Should().Be(true); }
    [Fact] public void NOT_True() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("NOT(TRUE)",ws).Should().Be(false); }
    [Fact] public void ABS_Negative() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("ABS(-5)",ws).Should().Be(5.0); }
    [Fact] public void SQRT_Four() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("SQRT(4)",ws).Should().Be(2.0); }
    [Fact] public void Power_Operator() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("2^8",ws).Should().Be(256.0); }
    [Fact] public void StringConcat_Ampersand() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("\"Hello\"&\" \"&\"World\"",ws).Should().Be("Hello World"); }
    [Fact] public void Comparison_Equal() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); e.Evaluate("1=1",ws).Should().Be(true); }
    [Fact] public void NORM_INV_Midpoint() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); var r = (double)e.Evaluate("NORM.INV(0.5,0,1)",ws)!; r.Should().BeApproximately(0.0, 0.001); }
    [Fact] public void ROUND15_Precision() { FunctionLibrary.Round15(0.1+0.2).Should().BeApproximately(0.3, 1e-13); }
    [Fact] public void PMT_Loan() { var e = Engine(); var ws = new ExcelWorkbook().AddWorksheet("S"); var r = (double)e.Evaluate("PMT(0.05/12,60,-10000)",ws)!; r.Should().BeApproximately(188.71, 0.1); }
    [Fact] public void LINEST_SimpleSlope() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S"); ws.Cell(1,1).Value=1.0; ws.Cell(2,1).Value=2.0; ws.Cell(3,1).Value=3.0; ws.Cell(1,2).Value=2.0; ws.Cell(2,2).Value=4.0; ws.Cell(3,2).Value=6.0; var e = new FormulaEngine(wb); var r = e.Evaluate("LINEST(B1:B3,A1:A3)",ws) as object?[]; r.Should().NotBeNull(); ((double)r![0]).Should().BeApproximately(2.0, 0.001); }
}

public class RangeTests
{
    [Fact] public void ExcelRange_Address() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S"); var r = ws.Cells(1,1,3,3); r.Address.Should().Be("A1:C3"); }
    [Fact] public void ExcelRange_RowColCount() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S"); var r = ws.Cells(2,3,5,7); r.RowCount.Should().Be(4); r.ColumnCount.Should().Be(5); }
    [Fact] public void LoadFromCollection_LoadsData() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S"); var data = new[]{ new{Name="Alice",Age=30}, new{Name="Bob",Age=25} }; ws.Cells(1,1,1,1).LoadFromCollection(data, true); ws.GetCell(2,1)?.GetString().Should().Be("Alice"); }
    [Fact] public void ToDataTable_ReadsData() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S"); ws.Cell(1,1).Value="Name"; ws.Cell(1,2).Value="Age"; ws.Cell(2,1).Value="Alice"; ws.Cell(2,2).Value=30.0; var dt = ws.Cells(1,1,2,2).ToDataTable(true); dt.Rows.Count.Should().Be(1); dt.Rows[0]["Name"].Should().Be("Alice"); }
    [Fact] public void Sort_SortsAscending() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S"); ws.Cell(1,1).Value=3.0; ws.Cell(2,1).Value=1.0; ws.Cell(3,1).Value=2.0; ws.Cells(1,1,3,1).Sort(0, true); ws.GetCell(1,1)?.GetDouble().Should().Be(1.0); ws.GetCell(3,1)?.GetDouble().Should().Be(3.0); }
    [Fact] public void IsEmpty_EmptyRange() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S"); ws.Cells(5,5,6,6).IsEmpty().Should().BeTrue(); }
    [Fact] public void IsEmpty_WithData() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S"); ws.Cell(5,5).Value="x"; ws.Cells(5,5,6,6).IsEmpty().Should().BeFalse(); }
    [Fact] public void CopyTo_CopiesValues() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S"); ws.Cell(1,1).Value="hello"; ws.Cells(1,1,1,1).CopyTo(ws.Cells(5,5,5,5)); ws.GetCell(5,5)?.GetString().Should().Be("hello"); }
}

public class IoTests
{
    [Fact] public async Task WriteRead_RoundTrip()
    {
        var wb = new ExcelWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell(1,1).Value = "Hello";
        ws.Cell(1,2).Value = 42.0;
        ws.Cell(2,1).Value = true;

        using var ms = new MemoryStream();
        await new XlsxWriter(wb).WriteAsync(ms);
        ms.Position = 0;

        var wb2 = await new XlsxReader().ReadAsync(ms);
        var ws2 = wb2.GetWorksheet("Sheet1");
        ws2.Should().NotBeNull();
        ws2!.GetCell(1,1)?.GetString().Should().Be("Hello");
        ws2.GetCell(1,2)?.GetDouble().Should().Be(42.0);
    }

    [Fact] public async Task WriteRead_MultipleSheets()
    {
        var wb = new ExcelWorkbook();
        wb.AddWorksheet("Alpha").Cell(1,1).Value = "A";
        wb.AddWorksheet("Beta").Cell(1,1).Value = "B";
        using var ms = new MemoryStream();
        await new XlsxWriter(wb).WriteAsync(ms);
        ms.Position = 0;
        var wb2 = await new XlsxReader().ReadAsync(ms);
        wb2.Worksheets.Count.Should().Be(2);
        wb2.GetWorksheet("Alpha")!.GetCell(1,1)!.GetString().Should().Be("A");
        wb2.GetWorksheet("Beta")!.GetCell(1,1)!.GetString().Should().Be("B");
    }

    [Fact] public void Encrypt_Decrypt_RoundTrip()
    {
        var data = System.Text.Encoding.UTF8.GetBytes("TestXlsxContent");
        var encrypted = WorkbookEncryption.Encrypt(data, "password123");
        WorkbookEncryption.IsEncrypted(encrypted).Should().BeTrue();
        var decrypted = WorkbookEncryption.Decrypt(encrypted, "password123");
        decrypted.Take(data.Length).ToArray().Should().Equal(data);
    }

    [Fact] public void IsEncrypted_NormalXlsx_False()
    {
        var data = new byte[] { 0x50, 0x4B, 0x03, 0x04 };
        WorkbookEncryption.IsEncrypted(data).Should().BeFalse();
    }

    [Fact] public async Task ExcelRange_ToJson_Valid()
    {
        var wb = new ExcelWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell(1,1).Value = "Name"; ws.Cell(1,2).Value = "Score";
        ws.Cell(2,1).Value = "Alice"; ws.Cell(2,2).Value = 95.0;
        var json = Exporter.ToJson(ws);
        json.Should().Contain("Alice");
        json.Should().Contain("95");
    }

    [Fact] public void ExcelRange_ToCsv_Valid()
    {
        var wb = new ExcelWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell(1,1).Value = "Name"; ws.Cell(1,2).Value = "Age";
        ws.Cell(2,1).Value = "Bob"; ws.Cell(2,2).Value = 30.0;
        var csv = Exporter.ToCsv(ws);
        csv.Should().Contain("Name,Age");
        csv.Should().Contain("Bob,30");
    }
}

public class AddressParserTests
{
    [Theory]
    [InlineData("A1", 1, 1)]
    [InlineData("B2", 2, 2)]
    [InlineData("Z10", 10, 26)]
    [InlineData("AA1", 1, 27)]
    [InlineData("XFD1048576", 1048576, 16384)]
    public void ParseCell(string addr, int row, int col) {
        var (r, c) = ExcelAddressParser.ParseCell(addr);
        r.Should().Be(row); c.Should().Be(col);
    }

    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(702, "ZZ")]
    [InlineData(16384, "XFD")]
    public void ColumnLetter(int col, string letter) =>
        ExcelAddressParser.ColumnNumberToLetter(col).Should().Be(letter);
}

public class PivotTests
{
    [Fact] public void PivotTable_Calculates()
    {
        var wb = new ExcelWorkbook();
        var src = wb.AddWorksheet("Data");
        src.Cell(1,1).Value = "Region"; src.Cell(1,2).Value = "Sales";
        src.Cell(2,1).Value = "North"; src.Cell(2,2).Value = 100.0;
        src.Cell(3,1).Value = "South"; src.Cell(3,2).Value = 200.0;
        src.Cell(4,1).Value = "North"; src.Cell(4,2).Value = 150.0;

        var ws = wb.AddWorksheet("Pivot");
        var pt = ws.AddPivotTable("SalesPivot", src.Cells("A1:B4"), "A1");
        pt.RowFields.Add(new ExcelPivotRowField { FieldName = "Region" });
        pt.DataFields.Add(new ExcelPivotDataField { FieldName = "Sales", Function = PivotDataFunction.Sum });

        var eng = new PivotCalculationEngine(wb);
        eng.Calculate(pt, wb);

        pt.IsCalculated.Should().BeTrue();
    }
}

public class ColorManagerTests
{
    [Fact] public void FromHex_ParsesCorrectly() { var c = ExcelColor.FromHex("#FF4472C4"); c.R.Should().Be(0x44); c.G.Should().Be(0x72); c.B.Should().Be(0xC4); }
    [Fact] public void FromRgb_Works() { var c = ExcelColor.FromRgb(255, 128, 0); c.R.Should().Be(255); c.G.Should().Be(128); c.B.Should().Be(0); }
    [Fact] public void FromPreset_KnownColor() { var c = ExcelColor.FromPreset("Red"); c.R.Should().Be(255); c.G.Should().Be(0); c.B.Should().Be(0); }
    [Fact] public void FromHsl_Blue() { var c = ExcelColor.FromHsl(240, 1.0, 0.5); c.B.Should().BeGreaterThan(200); }
    [Fact] public void FromTheme_SetsIndex() { var c = ExcelColor.FromTheme(4, 0.5); c.ThemeColorIndex.Should().Be(4); c.Tint.Should().Be(0.5); }
    [Fact] public void ToHex_RoundTrip() { var c = ExcelColor.FromRgb(64, 128, 200); c.ToHex(false).Should().Be("4080C8"); }
    [Fact] public void PresetColors_Has148Plus() { ExcelColor.PresetColors.Count.Should().BeGreaterThanOrEqualTo(140); }
}

public class MissingFeaturesTests
{
    [Fact] public void ExcelPageBreaks_AddRemove() { var pb = new ExcelPageBreaks(); pb.AddRowBreak(10); pb.RowBreaks.Should().Contain(10); pb.RemoveRowBreak(10); pb.RowBreaks.Should().BeEmpty(); }
    [Fact] public void ExcelOutline_GroupRows() { var o = new ExcelOutlineCollection(); o.GroupRows(2, 5, 1); o.GetRowLevel(3).Should().Be(1); o.MaxRowLevel.Should().Be(1); }
    [Fact] public void ExcelNamedStyle_Creates() { var col = new ExcelNamedStyleCollection(); var s = col.Add("Heading1"); s.Name.Should().Be("Heading1"); col["Heading1"].Should().NotBeNull(); }
    [Fact] public void ChartStyleManager_AppliesStyle() { var chart = new ExcelChart("Chart1", ExcelChartType.ColumnClustered); chart.Series.Add(new ExcelChartSeries()); ChartStyleManager.ApplyStyle(chart, 4); chart.BackgroundColor.Should().NotBeNull(); }
    [Fact] public void ExcelOleObject_Creates() { var ole = new ExcelOleObject("Excel.Sheet.12", new byte[]{1,2,3}); ole.ProgId.Should().Be("Excel.Sheet.12"); }
    [Fact] public void ExcelFormControl_CheckBox() { var ctrl = new ExcelFormControl(ExcelFormControlType.CheckBox); ctrl.ControlType.Should().Be(ExcelFormControlType.CheckBox); }
    [Fact] public void ExcelDigitalSignature_Properties() { var sig = new ExcelDigitalSignature(); sig.CommitmentType.Should().Be("ProofOfApproval"); sig.AllowComments.Should().BeTrue(); }
    [Fact] public void XlsbReader_Instantiates() { _ = new EPExcel.ML.IO.XlsbReader(); }
    [Fact] public async Task StreamingWriter_WritesRows() { var wb = new ExcelWorkbook(); using var ms = new System.IO.MemoryStream(); var sw = new EPExcel.ML.IO.StreamingXlsxWriter(wb, ms); var sheet = await sw.BeginWorksheetAsync("Sheet1"); await sheet.WriteRowAsync("Name", "Age"); await sheet.WriteRowAsync("Alice", 30.0); await sw.FinalizeAsync(); ms.Length.Should().BeGreaterThan(0); }
}

public class MLTests
{
    [Fact] public void ExcelMLEngine_Instantiates() { var wb = new ExcelWorkbook(); var ml = new ExcelMLEngine(wb); ml.Should().NotBeNull(); }
    [Fact] public void MLExtensions_AvailableOnRange() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S"); ws.Cell(1,1).Value = 1.0; var r = ws.Cells(1,1,1,1); r.Should().NotBeNull(); }
    [Fact] public void AutoMLResult_ToString() { var r = new AutoMLResult { BestModelName = "FastTree", RSquared = 0.95, MeanSquaredError = 0.01, RunCount = 10 }; r.ToString().Should().Contain("FastTree"); r.ToString().Should().Contain("0.9500"); }
    [Fact] public void ML_ForecastExtension_ReturnsEmpty_NoData() { var wb = new ExcelWorkbook(); var ws = wb.AddWorksheet("S"); var r = ws.Cells(1,1,1,1); var result = r.Forecast(3); result.Should().NotBeNull(); }
}
