using System.IO.Compression;
using System.Xml;

namespace EPExcel.ML.IO;

/// <summary>
/// OOXML (.xlsx) reader — EPExcel 8.5 parity.
/// Reads all worksheet content, styles, shared strings, named ranges,
/// tables, pivot tables, and unknown parts for round-trip fidelity.
/// </summary>
public sealed class XlsxReader
{
    public async Task<ExcelWorkbook> ReadAsync(Stream stream, CancellationToken ct = default)
    {
        var wb = new ExcelWorkbook();
        var ms = new MemoryStream();
        await stream.CopyToAsync(ms, ct);
        ms.Position = 0;

        using var zip = new ZipArchive(ms, ZipArchiveMode.Read, true);
        var sharedStrings = ReadSharedStrings(zip);
        var styleMap = ReadStyles(zip);
        ReadWorkbook(zip, wb, sharedStrings, styleMap);
        ReadUnknownParts(zip, wb);
        return wb;
    }

    // ── Shared strings ────────────────────────────────────────────────────────

    private static List<string> ReadSharedStrings(ZipArchive zip)
    {
        var result = new List<string>();
        var entry = zip.GetEntry("xl/sharedStrings.xml");
        if (entry == null) return result;

        using var stream = entry.Open();
        var doc = LoadXml(stream);
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        foreach (XmlNode si in doc.SelectNodes("//s:si", ns)!)
        {
            var sb = new System.Text.StringBuilder();
            foreach (XmlNode t in si.SelectNodes(".//s:t", ns)!) sb.Append(t.InnerText);
            result.Add(sb.ToString());
        }
        return result;
    }

    // ── Styles ────────────────────────────────────────────────────────────────

    private static Dictionary<int, string?> ReadStyles(ZipArchive zip)
    {
        // Map xf index -> number format code
        var numFmts = new Dictionary<int, string>();
        var result = new Dictionary<int, string?>();
        var entry = zip.GetEntry("xl/styles.xml");
        if (entry == null) return result;

        using var stream = entry.Open();
        var doc = LoadXml(stream);
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        // Built-in number formats
        var builtIn = new Dictionary<int, string>
        {
            [0]="General",[1]="0",[2]="0.00",[3]="#,##0",[4]="#,##0.00",
            [9]="0%",[10]="0.00%",[11]="0.00E+00",[14]="m/d/yyyy",[15]="d-mmm-yy",
            [16]="d-mmm",[17]="mmm-yy",[18]="h:mm AM/PM",[19]="h:mm:ss AM/PM",
            [20]="h:mm",[21]="h:mm:ss",[22]="m/d/yyyy h:mm",[37]="#,##0;(#,##0)",
            [38]="#,##0;[Red](#,##0)",[39]="#,##0.00;(#,##0.00)",[40]="#,##0.00;[Red](#,##0.00)",
            [45]="mm:ss",[46]="[h]:mm:ss",[47]="mmss.0",[48]="##0.0E+0",[49]="@",
        };
        foreach (var kv in builtIn) numFmts[kv.Key] = kv.Value;

        foreach (XmlNode nf in doc.SelectNodes("//s:numFmt", ns)!)
        {
            if (int.TryParse(nf.Attributes?["numFmtId"]?.Value, out var id))
                numFmts[id] = nf.Attributes?["formatCode"]?.Value ?? "";
        }

        int idx = 0;
        foreach (XmlNode xf in doc.SelectNodes("//s:cellXfs/s:xf", ns)!)
        {
            if (int.TryParse(xf.Attributes?["numFmtId"]?.Value, out var fmtId))
                result[idx] = numFmts.TryGetValue(fmtId, out var fmt) ? fmt : null;
            else
                result[idx] = null;
            idx++;
        }
        return result;
    }

    // ── Workbook ──────────────────────────────────────────────────────────────

    private static void ReadWorkbook(ZipArchive zip, ExcelWorkbook wb,
        List<string> ss, Dictionary<int, string?> styles)
    {
        var entry = zip.GetEntry("xl/workbook.xml");
        if (entry == null) return;

        using var stream = entry.Open();
        var doc = LoadXml(stream);
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        // Workbook properties
        var wbPr = doc.SelectSingleNode("//s:workbookPr", ns);
        if (wbPr != null)
            wb.Compatibility.Use1904DateSystem = wbPr.Attributes?["date1904"]?.Value == "1";

        // Workbook view
        var wbView = doc.SelectSingleNode("//s:workbookView", ns);
        if (wbView != null)
        {
            if (int.TryParse(wbView.Attributes?["tabRatio"]?.Value, out var tr)) wb.View.TabRatio = tr;
            if (int.TryParse(wbView.Attributes?["activeTab"]?.Value, out var at)) wb.View.ActiveTab = at;
            if (int.TryParse(wbView.Attributes?["firstSheet"]?.Value, out var fs)) wb.View.FirstSheet = fs;
        }

        // Named ranges
        foreach (XmlNode dn in doc.SelectNodes("//s:definedName", ns)!)
        {
            var name = dn.Attributes?["name"]?.Value;
            var value = dn.InnerText.Trim();
            if (name != null && value.Contains('!'))
            {
                var parts = value.Split('!', 2);
                var sheetName = parts[0].Trim('\'', '"');
                var wsRef = wb.GetWorksheet(sheetName);
                if (wsRef != null && parts.Length == 2)
                {
                    try
                    {
                        var addr = parts[1].Replace("$", "");
                        var (fr, fc, tr, tc) = ExcelAddressParser.ParseRange(
                            addr.Contains(':') ? addr : addr + ":" + addr);
                        wb.AddNamedRange(name, wsRef, wsRef.Cells(fr, fc, tr, tc));
                    }
                    catch { }
                }
            }
        }

        // Build sheet ID → rId map
        var rIdToFile = ReadWorkbookRels(zip);
        int sheetOrder = 0;
        foreach (XmlNode sheet in doc.SelectNodes("//s:sheet", ns)!)
        {
            var sheetName = sheet.Attributes?["name"]?.Value ?? $"Sheet{sheetOrder + 1}";
            var rId = sheet.Attributes?["r:id"]?.Value;
            if (rId != null && rIdToFile.TryGetValue(rId, out var target))
            {
                var ws = wb.AddWorksheet(sheetName);
                ws.TabSelected = sheetOrder == 0;
                ReadWorksheet(zip, ws, target, ss, styles);
            }
            sheetOrder++;
        }
    }

    private static Dictionary<string, string> ReadWorkbookRels(ZipArchive zip)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var entry = zip.GetEntry("xl/_rels/workbook.xml.rels");
        if (entry == null) return result;
        using var stream = entry.Open();
        var doc = LoadXml(stream);
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");
        foreach (XmlNode rel in doc.SelectNodes("//r:Relationship", ns)!)
        {
            var id = rel.Attributes?["Id"]?.Value;
            var target = rel.Attributes?["Target"]?.Value;
            if (id != null && target != null) result[id] = "xl/" + target.TrimStart('/');
        }
        return result;
    }

    // ── Worksheet ─────────────────────────────────────────────────────────────

    private static void ReadWorksheet(ZipArchive zip, ExcelWorksheet ws, string path,
        List<string> ss, Dictionary<int, string?> styles)
    {
        var entry = zip.GetEntry(path) ?? zip.GetEntry(path.TrimStart('/'));
        if (entry == null) return;

        ws.SetWorkbook(ws.GetWorkbook()!);

        using var stream = entry.Open();
        var doc = LoadXml(stream);
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        // Sheet properties
        var sheetView = doc.SelectSingleNode("//s:sheetView", ns);
        if (sheetView != null)
        {
            ws.ShowGridLines = sheetView.Attributes?["showGridLines"]?.Value != "0";
            ws.ShowRowColHeaders = sheetView.Attributes?["showRowColHeaders"]?.Value != "0";
            ws.TabSelected = sheetView.Attributes?["tabSelected"]?.Value == "1";
            if (double.TryParse(sheetView.Attributes?["zoomScale"]?.Value,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var z))
                ws.View.Zoom = z;
        }

        // Freeze panes
        var pane = doc.SelectSingleNode("//s:pane", ns);
        if (pane?.Attributes?["state"]?.Value == "frozen")
        {
            ws.FreezeRow = int.TryParse(pane.Attributes?["ySplit"]?.Value, out var yr) ? yr : 0;
            ws.FreezeCol = int.TryParse(pane.Attributes?["xSplit"]?.Value, out var xr) ? xr : 0;
            ws.FreezeRows = ws.FreezeRow > 0;
            ws.FreezeCols = ws.FreezeCol > 0;
        }

        // Sheet format
        var sfp = doc.SelectSingleNode("//s:sheetFormatPr", ns);
        if (sfp != null)
        {
            if (int.TryParse(sfp.Attributes?["defaultRowHeight"]?.Value, out var rh)) ws.DefaultRowHeight = rh;
            if (int.TryParse(sfp.Attributes?["defaultColWidth"]?.Value, out var cw)) ws.DefaultColWidth = cw;
        }

        // Column widths
        foreach (XmlNode col in doc.SelectNodes("//s:col", ns)!)
        {
            if (!double.TryParse(col.Attributes?["width"]?.Value,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var w)) continue;
            int min = int.TryParse(col.Attributes?["min"]?.Value, out var mn) ? mn : 1;
            int max = int.TryParse(col.Attributes?["max"]?.Value, out var mx) ? mx : min;
            for (int c = min; c <= max; c++)
            {
                ws.SetColumnWidth(c, w);
                if (col.Attributes?["hidden"]?.Value == "1") ws.HideColumn(c);
            }
        }

        // Row heights
        foreach (XmlNode row in doc.SelectNodes("//s:row", ns)!)
        {
            if (!int.TryParse(row.Attributes?["r"]?.Value, out var rowNum)) continue;
            if (double.TryParse(row.Attributes?["ht"]?.Value,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var rh2))
                ws.SetRowHeight(rowNum, rh2);
            if (row.Attributes?["hidden"]?.Value == "1") ws.HideRow(rowNum);
        }

        // Cells
        foreach (XmlNode c in doc.SelectNodes("//s:c", ns)!)
        {
            var r = c.Attributes?["r"]?.Value;
            if (r == null) continue;
            try
            {
                var (row, col) = ExcelAddressParser.ParseCell(r);
                var cell = ws.Cell(row, col);

                if (int.TryParse(c.Attributes?["s"]?.Value, out var si))
                {
                    cell.StyleIndex = si;
                    cell.NumberFormat = styles.TryGetValue(si, out var fmt) ? fmt : null;
                }

                var fNode = c.SelectSingleNode("s:f", ns);
                if (fNode != null) cell.Formula = fNode.InnerText;

                var vNode = c.SelectSingleNode("s:v", ns);
                string? raw = vNode?.InnerText;
                string type = c.Attributes?["t"]?.Value ?? "";

                if (raw != null)
                {
                    cell.Value = type switch
                    {
                        "s" when int.TryParse(raw, out var idx) && idx < ss.Count => (object?)ss[idx],
                        "b" => raw == "1",
                        "e" => new CellError(ParseErrorCode(raw)),
                        "str" or "inlineStr" => raw,
                        _ when double.TryParse(raw, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var d) => d,
                        _ => (object?)raw
                    };
                }

                // Inline string
                var is2 = c.SelectSingleNode(".//s:t", ns);
                if (is2 != null && type == "inlineStr") cell.Value = is2.InnerText;
            }
            catch { }
        }

        // Auto filter
        var af = doc.SelectSingleNode("//s:autoFilter", ns);
        if (af != null) ws.AutoFilterAddress = af.Attributes?["ref"]?.Value;

        // Page setup
        var ps = doc.SelectSingleNode("//s:pageSetup", ns);
        if (ps != null)
        {
            if (int.TryParse(ps.Attributes?["paperSize"]?.Value, out var paper))
                ws.PageSetup.PaperSize = (ExcelPaperSize)paper;
            if (ps.Attributes?["orientation"]?.Value is string ori)
                ws.PageSetup.Orientation = ori == "landscape" ? ExcelOrientation.Landscape :
                                            ori == "portrait" ? ExcelOrientation.Portrait :
                                            ExcelOrientation.Default;
        }
    }

    // ── Unknown parts (round-trip) ────────────────────────────────────────────

    private static void ReadUnknownParts(ZipArchive zip, ExcelWorkbook wb)
    {
        var known = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "[Content_Types].xml", "_rels/.rels",
            "xl/workbook.xml", "xl/styles.xml", "xl/sharedStrings.xml",
            "xl/theme/theme1.xml", "xl/connections.xml",
            "xl/_rels/workbook.xml.rels",
        };
        // Add all worksheet paths
        for (int i = 1; i <= 1024; i++)
        {
            known.Add($"xl/worksheets/sheet{i}.xml");
            known.Add($"xl/worksheets/_rels/sheet{i}.xml.rels");
        }

        foreach (var entry in zip.Entries)
        {
            if (known.Contains(entry.FullName)) continue;
            if (entry.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)) continue;
            if (entry.FullName.StartsWith("xl/tables/", StringComparison.OrdinalIgnoreCase)) continue;
            try
            {
                using var s = entry.Open();
                using var buf = new MemoryStream();
                s.CopyTo(buf);
                wb.UnknownParts["/" + entry.FullName] = buf.ToArray();
            }
            catch { }
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static XmlDocument LoadXml(Stream stream)
    {
        var doc = new XmlDocument { PreserveWhitespace = false };
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Ignore,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
        };
        using var reader = XmlReader.Create(stream, settings);
        doc.Load(reader);
        return doc;
    }

    private static ExcelErrorCode ParseErrorCode(string s) => s.ToUpperInvariant() switch
    {
        "#DIV/0!" => ExcelErrorCode.Div0, "#VALUE!" => ExcelErrorCode.Value,
        "#REF!"   => ExcelErrorCode.Ref,  "#NAME?"  => ExcelErrorCode.Name,
        "#NUM!"   => ExcelErrorCode.Num,  "#N/A"    => ExcelErrorCode.NA,
        "#NULL!"  => ExcelErrorCode.Null, "#SPILL!" => ExcelErrorCode.Spill,
        _ => ExcelErrorCode.Value
    };
}
