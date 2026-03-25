using EPExcel.ML.Samples;

Console.WriteLine("╔══════════════════════════════════════════════════════╗");
Console.WriteLine("║          EPExcel.ML — Sample Project                  ║");
Console.WriteLine("║  MIT-licensed Excel library + Microsoft.ML for .NET  ║");
Console.WriteLine("║  Supports: .NET 7 | .NET 8 | .NET 9 | .NET 10        ║");
Console.WriteLine("╚══════════════════════════════════════════════════════╝");

await Sample01_BasicWorkbook.RunAsync();
await Sample02_StylingAndFormatting.RunAsync();
await Sample03_ChartsAndPivot.RunAsync();
await Sample04_Formulas.RunAsync();
await Sample05_MLFeatures.RunAsync();
await Sample06_MigrationAndEncryption.RunAsync();

Console.WriteLine("\n✅ All samples completed.");
Console.WriteLine("   Output files written to current directory.");
