namespace EPExcel.ML;

public sealed class ExcelTheme
{
    public string Name { get; set; } = "Office";
    public ExcelThemeFonts Fonts { get; set; } = new();
    public string[] Colors { get; set; } = [
        "FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD",
        "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646",
        "0070C0", "00B050"
    ];

    public static ExcelTheme Office => new() { Name = "Office" };
    public static ExcelTheme Dark => new()
    {
        Name = "Dark",
        Colors = ["000000", "FFFFFF", "1F1F1F", "F2F2F2", "4472C4", "ED7D31", "A9D18E", "7030A0", "00B0F0", "FF0000", "00B050", "FFC000"]
    };
}

public sealed class ExcelThemeFonts
{
    public string HeadingLatin { get; set; } = "Calibri Light";
    public string BodyLatin { get; set; } = "Calibri";
    public string HeadingEastAsian { get; set; } = "";
    public string BodyEastAsian { get; set; } = "";
}
