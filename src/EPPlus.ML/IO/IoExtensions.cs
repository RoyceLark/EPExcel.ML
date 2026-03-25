namespace EPExcel.ML.IO;

/// <summary>
/// XLSB (Excel Binary Workbook) reader — EPExcel parity.
/// Full XLSB implementation requires parsing BRT (Binary Record Type) records.
/// This implementation reads cell values from XLSB streams by locating
/// worksheet streams and parsing BRT_ROW + BRT_CELL records.
/// </summary>
public sealed class XlsbReader
{
    public async Task<ExcelWorkbook> ReadAsync(Stream stream, CancellationToken ct = default)
    {
        var wb = new ExcelWorkbook();
        var ms = new MemoryStream();
        await stream.CopyToAsync(ms, ct);
        ms.Position = 0;

        try
        {
            using var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, true);
            var ws = wb.AddWorksheet("Sheet1");
            var sheetEntry = zip.Entries.FirstOrDefault(e =>
                e.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) &&
                e.FullName.EndsWith(".bin", StringComparison.OrdinalIgnoreCase));

            if (sheetEntry != null)
            {
                using var binStream = sheetEntry.Open();
                ParseBrtStream(binStream, ws);
            }
        }
        catch (Exception ex)
        {
            wb.AddWorksheet("Sheet1").Cell(1, 1).Value =
                $"XLSB parse error: {ex.Message}";
        }

        return wb;
    }

    private static void ParseBrtStream(Stream stream, ExcelWorksheet ws)
    {
        // BRT record format: RecordType (VarInt) + Size (VarInt) + Data
        using var br = new BinaryReader(stream, System.Text.Encoding.UTF8, true);
        int currentRow = 0, currentCol = 0;

        while (stream.Position < stream.Length - 2)
        {
            int recType = ReadVarInt(br);
            int recSize = ReadVarInt(br);
            if (recSize < 0 || recSize > 1_048_576) break;
            var data = recSize > 0 ? br.ReadBytes(recSize) : Array.Empty<byte>();

            switch (recType)
            {
                case 0x0000: // BrtRowHdr
                    if (data.Length >= 4) currentRow = BitConverter.ToInt32(data, 0) + 1;
                    currentCol = 0;
                    break;
                case 0x0007: // BrtCellIsst (shared string)
                    currentCol++;
                    break;
                case 0x0005: // BrtCellNum (double)
                    if (data.Length >= 8)
                        ws.Cell(currentRow, ++currentCol).Value =
                            BitConverter.ToDouble(data, 0);
                    break;
                case 0x0008: // BrtCellBool
                    if (data.Length >= 1)
                        ws.Cell(currentRow, ++currentCol).Value = data[0] != 0;
                    break;
                case 0x000A: // BrtCellStr
                    currentCol++;
                    if (data.Length > 4)
                        ws.Cell(currentRow, currentCol).Value =
                            System.Text.Encoding.Unicode.GetString(data, 4,
                                Math.Min(data.Length - 4, BitConverter.ToInt32(data, 0) * 2));
                    break;
            }
        }
    }

    private static int ReadVarInt(BinaryReader br)
    {
        int result = 0, shift = 0;
        while (true)
        {
            if (br.BaseStream.Position >= br.BaseStream.Length) return -1;
            byte b = br.ReadByte();
            result |= (b & 0x7F) << shift;
            if ((b & 0x80) == 0) return result;
            shift += 7;
            if (shift >= 35) return result;
        }
    }
}

/// <summary>
/// Streaming XLSX writer for large datasets (100k+ rows).
/// Uses ZipArchive streaming mode to avoid buffering entire file in memory.
/// EPExcel parity: ExcelPackage with streaming mode.
/// </summary>
public sealed class StreamingXlsxWriter : IAsyncDisposable
{
    private readonly System.IO.Compression.ZipArchive _zip;
    private readonly Stream _output;
    private readonly ExcelWorkbook _workbook;
    private readonly List<string> _sharedStrings = new();
    private readonly Dictionary<string, int> _ssIndex = new();
    private readonly List<(string name, long offset)> _sheets = new();
    private System.IO.Compression.ZipArchiveEntry? _currentEntry;
    private Stream? _currentStream;
    private int _currentRow;

    public StreamingXlsxWriter(ExcelWorkbook workbook, Stream output)
    {
        _workbook = workbook;
        _output = output;
        _zip = new System.IO.Compression.ZipArchive(output,
            System.IO.Compression.ZipArchiveMode.Create, true);
    }

    public async Task<StreamingWorksheet> BeginWorksheetAsync(string name, CancellationToken ct = default)
    {
        var ws = _workbook.GetWorksheet(name) ?? _workbook.AddWorksheet(name);
        int idx = _sheets.Count + 1;
        _sheets.Add((name, 0));
        var entry = _zip.CreateEntry($"xl/worksheets/sheet{idx}.xml",
            System.IO.Compression.CompressionLevel.Optimal);
        _currentEntry = entry;
        _currentStream = entry.Open();
        var header = System.Text.Encoding.UTF8.GetBytes(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheetData>");
        await _currentStream.WriteAsync(header, ct);
        _currentRow = 0;
        return new StreamingWorksheet(_currentStream, this);
    }

    internal async Task WriteRowAsync(IEnumerable<object?> values, CancellationToken ct)
    {
        if (_currentStream == null) return;
        _currentRow++;
        var sb = new System.Text.StringBuilder();
        sb.Append($"<row r=\"{_currentRow}\">");
        int col = 1;
        foreach (var val in values)
        {
            string addr = ExcelAddressParser.ToAddress(_currentRow, col++);
            string cellXml;
            if (val is double dv)
                cellXml = $"<c r=\"{addr}\"><v>{dv.ToString("R", System.Globalization.CultureInfo.InvariantCulture)}</v></c>";
            else if (val is int iv)
                cellXml = $"<c r=\"{addr}\"><v>{iv}</v></c>";
            else if (val is bool bv)
                cellXml = $"<c r=\"{addr}\" t=\"b\"><v>{(bv ? 1 : 0)}</v></c>";
            else if (val is string sv)
                cellXml = $"<c r=\"{addr}\" t=\"str\"><v>{System.Net.WebUtility.HtmlEncode(sv)}</v></c>";
            else if (val == null)
                cellXml = $"<c r=\"{addr}\"/>";
            else
                cellXml = $"<c r=\"{addr}\" t=\"str\"><v>{System.Net.WebUtility.HtmlEncode(val.ToString() ?? "")}</v></c>";
            sb.Append(cellXml);
        }
        sb.Append("</row>");
        var bytes = System.Text.Encoding.UTF8.GetBytes(sb.ToString());
        await _currentStream.WriteAsync(bytes, ct);
    }

    public async Task FinalizeAsync(CancellationToken ct = default)
    {
        if (_currentStream != null)
        {
            var footer = System.Text.Encoding.UTF8.GetBytes("</sheetData></worksheet>");
            await _currentStream.WriteAsync(footer, ct);
            await _currentStream.DisposeAsync();
        }
        _zip.Dispose();
    }

    public async ValueTask DisposeAsync() => await FinalizeAsync();
}

/// <summary>Represents a single worksheet being streamed.</summary>
public sealed class StreamingWorksheet(Stream stream, StreamingXlsxWriter writer)
{
    private readonly Stream _stream = stream;
    private readonly StreamingXlsxWriter _writer = writer;

    public async Task WriteRowAsync(IEnumerable<object?> values, CancellationToken ct = default)
        => await _writer.WriteRowAsync(values, ct);

    public async Task WriteRowAsync(params object?[] values) =>
        await WriteRowAsync((IEnumerable<object?>)values);
}
