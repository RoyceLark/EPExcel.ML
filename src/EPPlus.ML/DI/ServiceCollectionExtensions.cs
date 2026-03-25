using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;

namespace EPExcel.ML;

/// <summary>Dependency injection registration — EPExcel parity: AddExcelAI()</summary>
public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddEPExcelML(this IServiceCollection services,
        Action<EPExcelMLOptions>? configure = null)
    {
        if (configure != null)
            services.Configure(configure);
        services.AddTransient<ExcelWorkbook>();
        services.AddTransient<IO.XlsxReader>();
        services.AddTransient<IO.XlsxWriter>(sp =>
        {
            var wb = sp.GetRequiredService<ExcelWorkbook>();
            return new IO.XlsxWriter(wb);
        });
        return services;
    }
}

public sealed class EPExcelMLOptions
{
    public string? DefaultAuthor { get; set; }
    public string? DefaultCompany { get; set; }
    public bool CalculateOnSave { get; set; } = true;
    public bool UseSharedStrings { get; set; } = true;
    public int CompressionLevel { get; set; } = 6;
}

/// <summary>External workbook links — EPExcel 8.x parity.</summary>
public sealed class ExcelExternalLink(string name, string path)
{
    public string Name { get; } = name;
    public string FilePath { get; } = path;
    public bool IsbrokenLink { get; private set; }
    public List<ExcelExternalSheet> Sheets { get; } = new();

    public ExcelExternalSheet AddSheet(string name) =>
        Sheets.Count > 0 ? Sheets[^1] : Tap(new ExcelExternalSheet(name), Sheets.Add);

    public void BreakLink() => IsbrokenLink = true;

    private static T Tap<T>(T item, Action<T> action) { action(item); return item; }
}

public sealed class ExcelExternalSheet(string name)
{
    public string Name { get; } = name;
    private readonly Dictionary<string, object?> _cache = new();
    public void SetCachedValue(string address, object? value) => _cache[address] = value;
    public object? GetCachedValue(string address) => _cache.TryGetValue(address, out var v) ? v : null;
}
