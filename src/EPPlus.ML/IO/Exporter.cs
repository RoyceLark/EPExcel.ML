using System.Text.Json;

namespace EPExcel.ML.IO;

/// <summary>Export worksheet data to HTML, CSV, and JSON formats.</summary>
public static class Exporter
{
    public static string ToHtml(ExcelWorksheet ws, ExcelRange? range = null)
    {
        int fr = range?.FromRow ?? 1, fc = range?.FromCol ?? 1;
        int tr = range?.ToRow ?? ws.MaxRow, tc = range?.ToCol ?? ws.MaxCol;
        if (tr == 0 || tc == 0) return "<table></table>";
        var sb = new StringBuilder();
        sb.Append("<table border=\"1\" cellspacing=\"0\" cellpadding=\"4\" style=\"border-collapse:collapse;font-family:Calibri,sans-serif;font-size:11pt\">");
        for (int r = fr; r <= tr; r++)
        {
            sb.Append("<tr>");
            for (int c = fc; c <= tc; c++)
            {
                var cell = ws.GetCell(r, c);
                var val = cell?.DisplayValue;
                string content = val switch
                {
                    null or "" => "&nbsp;",
                    double d when cell?.NumberFormat?.Contains("d") == true ||
                                  cell?.NumberFormat?.Contains("m") == true
                        => DateTime.FromOADate(d).ToString("yyyy-MM-dd"),
                    double d => d.ToString("G", System.Globalization.CultureInfo.InvariantCulture),
                    bool b => b ? "TRUE" : "FALSE",
                    CellError e => e.ToString(),
                    _ => System.Net.WebUtility.HtmlEncode(val.ToString() ?? "")
                };
                string tag = r == fr ? "th" : "td";
                string style = "";
                if (cell?.StyleIndex > 0)
                {
                    var st = ws.GetWorkbook()?.Styles.GetStyle(cell.StyleIndex);
                    if (st?.Font.Bold == true) style += "font-weight:bold;";
                    if (st?.Font.Italic == true) style += "font-style:italic;";
                    if (st?.Fill.BackgroundColor != null)
                        style += $"background-color:#{st.Fill.BackgroundColor.TrimStart('#')};";
                    if (st?.Alignment != null && st.Alignment.Horizontal != ExcelHorizontalAlignment.General)
                        style += $"text-align:{st.Alignment.Horizontal.ToString().ToLowerInvariant()};";
                }
                sb.Append(style.Length > 0 ? $"<{tag} style=\"{style}\">{content}</{tag}>" : $"<{tag}>{content}</{tag}>");
            }
            sb.Append("</tr>");
        }
        sb.Append("</table>");
        return sb.ToString();
    }

    public static string ToCsv(ExcelWorksheet ws, ExcelRange? range = null, char delimiter = ',')
    {
        int fr = range?.FromRow ?? 1, fc = range?.FromCol ?? 1;
        int tr = range?.ToRow ?? ws.MaxRow, tc = range?.ToCol ?? ws.MaxCol;
        var sb = new StringBuilder();
        for (int r = fr; r <= tr; r++)
        {
            for (int c = fc; c <= tc; c++)
            {
                if (c > fc) sb.Append(delimiter);
                var val = ws.GetCell(r, c)?.DisplayValue;
                var s = val switch
                {
                    double d => d.ToString("G", System.Globalization.CultureInfo.InvariantCulture),
                    bool b => b ? "TRUE" : "FALSE",
                    CellError e => e.ToString(),
                    null => "",
                    _ => val.ToString() ?? ""
                };
                if (s.Contains(delimiter) || s.Contains('"') || s.Contains('\n'))
                    s = "\"" + s.Replace("\"", "\"\"") + "\"";
                sb.Append(s);
            }
            sb.AppendLine();
        }
        return sb.ToString();
    }

    public static string ToJson(ExcelWorksheet ws, ExcelRange? range = null, bool includeHeaders = true)
    {
        int fr = range?.FromRow ?? 1, fc = range?.FromCol ?? 1;
        int tr = range?.ToRow ?? ws.MaxRow, tc = range?.ToCol ?? ws.MaxCol;
        if (tr == 0 || tc == 0) return "[]";

        var headers = new List<string>();
        int dataStart = fr;
        if (includeHeaders)
        {
            for (int c = fc; c <= tc; c++)
                headers.Add(ws.GetCell(fr, c)?.GetString() ?? $"Col{c - fc + 1}");
            dataStart = fr + 1;
        }
        else
        {
            for (int c = fc; c <= tc; c++) headers.Add($"Col{c - fc + 1}");
        }

        var rows = new List<Dictionary<string, object?>>();
        for (int r = dataStart; r <= tr; r++)
        {
            var row = new Dictionary<string, object?>();
            for (int c = fc; c <= tc; c++)
            {
                var val = ws.GetCell(r, c)?.DisplayValue;
                row[headers[c - fc]] = val switch
                {
                    double d => d, bool b => b, null => null,
                    CellError e => e.ToString(), _ => val.ToString()
                };
            }
            rows.Add(row);
        }
        return JsonSerializer.Serialize(rows, new JsonSerializerOptions { WriteIndented = true });
    }

    public static string ToMarkdown(ExcelWorksheet ws, ExcelRange? range = null)
    {
        int fr = range?.FromRow ?? 1, fc = range?.FromCol ?? 1;
        int tr = range?.ToRow ?? ws.MaxRow, tc = range?.ToCol ?? ws.MaxCol;
        if (tr == 0 || tc == 0) return "";
        var sb = new StringBuilder();

        // Collect all values first to compute column widths
        var cells = new string[tr - fr + 1, tc - fc + 1];
        var widths = new int[tc - fc + 1];
        for (int r = fr; r <= tr; r++)
            for (int c = fc; c <= tc; c++)
            {
                var v = ws.GetCell(r, c)?.DisplayValue?.ToString() ?? "";
                cells[r - fr, c - fc] = v;
                widths[c - fc] = Math.Max(widths[c - fc], v.Length);
            }

        for (int r = 0; r <= tr - fr; r++)
        {
            sb.Append('|');
            for (int c = 0; c <= tc - fc; c++)
                sb.Append(' ').Append(cells[r, c].PadRight(widths[c])).Append(" |");
            sb.AppendLine();
            if (r == 0)
            {
                sb.Append('|');
                for (int c = 0; c <= tc - fc; c++)
                    sb.Append(new string('-', widths[c] + 2)).Append('|');
                sb.AppendLine();
            }
        }
        return sb.ToString();
    }
}
