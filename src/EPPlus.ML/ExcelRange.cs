using System.Data;
using System.Reflection;

namespace EPExcel.ML;

/// <summary>
/// Range of cells — EPExcel ExcelRange parity.
/// All fluent operations, LoadFromCollection, ToCollection, sort, copy, etc.
/// </summary>
public sealed class ExcelRange
{
    public ExcelWorksheet Worksheet { get; }
    public int FromRow { get; }
    public int FromCol { get; }
    public int ToRow { get; }
    public int ToCol { get; }

    public int RowCount => ToRow - FromRow + 1;
    public int ColumnCount => ToCol - FromCol + 1;

    public string Address => ExcelAddressParser.ToRangeAddress(FromRow, FromCol, ToRow, ToCol);
    public string FullAddress => $"'{Worksheet.Name}'!{Address}";

    public ExcelRange(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
    {
        Worksheet = ws;
        FromRow = Math.Min(fromRow, toRow);
        FromCol = Math.Min(fromCol, toCol);
        ToRow = Math.Max(fromRow, toRow);
        ToCol = Math.Max(fromCol, toCol);
    }

    // ── Value access ──────────────────────────────────────────────────────────

    public object? Value
    {
        get => IsSingleCell ? Worksheet.GetCell(FromRow, FromCol)?.Value : null;
        set
        {
            for (int r = FromRow; r <= ToRow; r++)
                for (int c = FromCol; c <= ToCol; c++)
                    Worksheet.Cell(r, c).Value = value;
        }
    }

    public object? this[int row, int col]
    {
        get => Worksheet.GetCell(FromRow + row - 1, FromCol + col - 1)?.Value;
        set => Worksheet.Cell(FromRow + row - 1, FromCol + col - 1).Value = value;
    }

    public string? Formula
    {
        get => IsSingleCell ? Worksheet.GetCell(FromRow, FromCol)?.Formula : null;
        set
        {
            for (int r = FromRow; r <= ToRow; r++)
                for (int c = FromCol; c <= ToCol; c++)
                    Worksheet.Cell(r, c).Formula = value;
        }
    }

    public string? StyleName
    {
        get => null;
        set { /* style name assignment */ }
    }

    // ── EPExcel-compatible fluent style chain ──────────────────────────────────
    // range.Style.Font.Bold = true  (EPExcel parity)
    // range.Style.Fill.BackgroundColor.SetColor(...)

    public ExcelRangeStyle Style => new(this);

    public bool IsSingleCell => FromRow == ToRow && FromCol == ToCol;

    // ── Style shortcuts ───────────────────────────────────────────────────────

    public void SetFont(Action<FontDef> configure)
    {
        for (int r = FromRow; r <= ToRow; r++)
            for (int c = FromCol; c <= ToCol; c++)
            {
                var cell = Worksheet.Cell(r, c);
                var style = Worksheet.GetWorkbook()?.Styles.GetStyle(cell.StyleIndex).Clone() ?? new CellStyleDef();
                configure(style.Font);
                cell.StyleIndex = Worksheet.GetWorkbook()?.Styles.RegisterStyle(style) ?? 0;
            }
    }

    public void SetFill(string? bgColor, ExcelFillPattern pattern = ExcelFillPattern.Solid)
    {
        for (int r = FromRow; r <= ToRow; r++)
            for (int c = FromCol; c <= ToCol; c++)
            {
                var cell = Worksheet.Cell(r, c);
                var style = Worksheet.GetWorkbook()?.Styles.GetStyle(cell.StyleIndex).Clone() ?? new CellStyleDef();
                style.Fill.PatternType = pattern;
                style.Fill.BackgroundColor = bgColor;
                cell.StyleIndex = Worksheet.GetWorkbook()?.Styles.RegisterStyle(style) ?? 0;
            }
    }

    public void SetNumberFormat(string format)
    {
        for (int r = FromRow; r <= ToRow; r++)
            for (int c = FromCol; c <= ToCol; c++)
                Worksheet.Cell(r, c).NumberFormat = format;
    }

    public void SetAlignment(ExcelHorizontalAlignment h, ExcelVerticalCellAlignment v = ExcelVerticalCellAlignment.Bottom)
    {
        for (int r = FromRow; r <= ToRow; r++)
            for (int c = FromCol; c <= ToCol; c++)
            {
                var cell = Worksheet.Cell(r, c);
                var style = Worksheet.GetWorkbook()?.Styles.GetStyle(cell.StyleIndex).Clone() ?? new CellStyleDef();
                style.Alignment.Horizontal = h;
                style.Alignment.Vertical = v;
                cell.StyleIndex = Worksheet.GetWorkbook()?.Styles.RegisterStyle(style) ?? 0;
            }
    }

    public void Merge()
    {
        Worksheet.MergedCells.Add(Address);
    }

    // ── IsEmpty (EPExcel 8 parity) ─────────────────────────────────────────────

    public bool IsEmpty(bool checkValue = true, bool checkFormula = true, bool checkComment = false)
    {
        for (int r = FromRow; r <= ToRow; r++)
            for (int c = FromCol; c <= ToCol; c++)
            {
                var cell = Worksheet.GetCell(r, c);
                if (cell == null) continue;
                if (checkValue && cell.Value != null && cell.Value.ToString() != "") return false;
                if (checkFormula && cell.Formula != null) return false;
                if (checkComment && Worksheet.Comments.Any(cm => cm.Row == r && cm.Col == c)) return false;
            }
        return true;
    }

    // ── BorderAround (EPExcel parity - overrides adjacent cell borders) ────────

    public void BorderAround(BorderLineStyle style, string? color = null)
    {
        // Top row
        for (int c = FromCol; c <= ToCol; c++)
        {
            var cell = Worksheet.Cell(FromRow, c);
            ApplyBorder(cell, "Top", style, color);
        }
        // Bottom row
        for (int c = FromCol; c <= ToCol; c++)
        {
            var cell = Worksheet.Cell(ToRow, c);
            ApplyBorder(cell, "Bottom", style, color);
            // Also set the top border of the cell below (EPExcel override fix)
            if (ToRow + 1 <= 1048576)
            {
                var below = Worksheet.GetCell(ToRow + 1, c);
                if (below != null) ApplyBorder(below, "Top", style, color);
            }
        }
        // Left col
        for (int r = FromRow; r <= ToRow; r++)
        {
            var cell = Worksheet.Cell(r, FromCol);
            ApplyBorder(cell, "Left", style, color);
        }
        // Right col
        for (int r = FromRow; r <= ToRow; r++)
        {
            var cell = Worksheet.Cell(r, ToCol);
            ApplyBorder(cell, "Right", style, color);
        }
    }

    private void ApplyBorder(ExcelCell cell, string side, BorderLineStyle style, string? color)
    {
        var wb = Worksheet.GetWorkbook();
        if (wb == null) return;
        var st = wb.Styles.GetStyle(cell.StyleIndex).Clone();
        var border = side switch
        {
            "Top"    => st.Border.Top,
            "Bottom" => st.Border.Bottom,
            "Left"   => st.Border.Left,
            "Right"  => st.Border.Right,
            _        => st.Border.Top
        };
        border.Style = style;
        if (color != null) border.Color = color;
        cell.StyleIndex = wb.Styles.RegisterStyle(st);
    }

    // ── InsertRange (EPExcel parity - returns the new range) ──────────────────

    public ExcelRange InsertRange(ExcelShiftTypeInsert shiftType)
    {
        if (shiftType == ExcelShiftTypeInsert.Down)
            Worksheet.InsertRow(FromRow, RowCount);
        else
            Worksheet.InsertColumn(FromCol, ColumnCount);
        return new ExcelRange(Worksheet, FromRow, FromCol, ToRow, ToCol);
    }

    // ── CopyTo ────────────────────────────────────────────────────────────────

    public void CopyTo(ExcelRange destination, ExcelRangeCopyOptions options = ExcelRangeCopyOptions.All)
    {
        int rowOff = destination.FromRow - FromRow;
        int colOff = destination.FromCol - FromCol;

        for (int r = FromRow; r <= ToRow; r++)
            for (int c = FromCol; c <= ToCol; c++)
            {
                var src = Worksheet.GetCell(r, c);
                if (src == null) continue;
                var dst = destination.Worksheet.Cell(r + rowOff, c + colOff);

                if (options.HasFlag(ExcelRangeCopyOptions.Values))
                    dst.Value = src.Value;
                if (options.HasFlag(ExcelRangeCopyOptions.Formulas) && src.Formula != null)
                    dst.Formula = RewriteFormula(src.Formula, rowOff, colOff);
                if (options.HasFlag(ExcelRangeCopyOptions.Styles))
                    dst.StyleIndex = src.StyleIndex;
                if (options.HasFlag(ExcelRangeCopyOptions.NumberFormat))
                    dst.NumberFormat = src.NumberFormat;
            }
    }

    internal static string RewriteFormula(string formula, int rowOff, int colOff)
    {
        if (rowOff == 0 && colOff == 0) return formula;
        return System.Text.RegularExpressions.Regex.Replace(
            formula,
            @"(?<colAbs>\$?)(?<col>[A-Z]{1,3})(?<rowAbs>\$?)(?<row>\d{1,7})",
            m =>
            {
                bool colAbs = m.Groups["colAbs"].Value == "$";
                bool rowAbs = m.Groups["rowAbs"].Value == "$";
                int col = ExcelAddressParser.ColumnLetterToNumber(m.Groups["col"].Value);
                int row = int.Parse(m.Groups["row"].Value);
                if (!colAbs) col += colOff;
                if (!rowAbs) row += rowOff;
                if (col < 1 || row < 1) return "#REF!";
                return $"{(colAbs ? "$" : "")}{ExcelAddressParser.ColumnNumberToLetter(col)}{(rowAbs ? "$" : "")}{row}";
            });
    }

    // ── Sort ──────────────────────────────────────────────────────────────────

    public void Sort(int sortColOffset = 0, bool ascending = true, bool hasHeaders = false)
    {
        int dataFromRow = hasHeaders ? FromRow + 1 : FromRow;
        if (dataFromRow > ToRow) return;

        int sortCol = FromCol + sortColOffset;
        var rows = Enumerable.Range(dataFromRow, ToRow - dataFromRow + 1)
            .OrderBy(r =>
            {
                var v = Worksheet.GetCell(r, sortCol)?.Value;
                return v is double d ? d.ToString("R") : v?.ToString() ?? "";
            })
            .ToList();

        if (!ascending) rows.Reverse();

        var snapshot = new Dictionary<(int, int), (object? val, string? fmt, string? formula, int style)>();
        for (int r = dataFromRow; r <= ToRow; r++)
            for (int c = FromCol; c <= ToCol; c++)
            {
                var cell = Worksheet.GetCell(r, c);
                snapshot[(r, c)] = (cell?.Value, cell?.NumberFormat, cell?.Formula, cell?.StyleIndex ?? 0);
            }

        for (int i = 0; i < rows.Count; i++)
        {
            int srcRow = rows[i];
            int dstRow = dataFromRow + i;
            for (int c = FromCol; c <= ToCol; c++)
            {
                var (val, fmt, formula, style) = snapshot[(srcRow, c)];
                var cell = Worksheet.Cell(dstRow, c);
                cell.Value = val;
                cell.NumberFormat = fmt;
                cell.Formula = formula;
                cell.StyleIndex = style;
            }
        }
    }

    // ── LoadFromCollection<T> ─────────────────────────────────────────────────

    public ExcelRange LoadFromCollection<T>(IEnumerable<T> collection,
        bool printHeaders = false, ExcelPrintHeaders header = ExcelPrintHeaders.ColumnCaption)
    {
        var items = collection.ToList();
        if (!items.Any()) return this;

        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead).ToArray();

        int row = FromRow;
        int startRow = FromRow;

        if (printHeaders)
        {
            for (int i = 0; i < props.Length; i++)
                Worksheet.Cell(row, FromCol + i).Value = props[i].Name;
            row++;
            startRow = row;
        }

        foreach (var item in items)
        {
            for (int i = 0; i < props.Length; i++)
            {
                var val = props[i].GetValue(item);
                Worksheet.Cell(row, FromCol + i).Value = val switch
                {
                    DateTime dt => dt.ToOADate(),
                    _ => val
                };
            }
            row++;
        }

        return new ExcelRange(Worksheet, startRow, FromCol, row - 1, FromCol + props.Length - 1);
    }

    public ExcelRange LoadFromArrays(IEnumerable<object?[]> rows)
    {
        int r = FromRow;
        int maxCols = 0;
        foreach (var row in rows)
        {
            for (int c = 0; c < row.Length; c++)
                Worksheet.Cell(r, FromCol + c).Value = row[c];
            maxCols = Math.Max(maxCols, row.Length);
            r++;
        }
        return new ExcelRange(Worksheet, FromRow, FromCol, r - 1, FromCol + maxCols - 1);
    }

    public ExcelRange LoadFromDataTable(DataTable table, bool printHeaders = false)
    {
        int row = FromRow;
        if (printHeaders)
        {
            for (int c = 0; c < table.Columns.Count; c++)
                Worksheet.Cell(row, FromCol + c).Value = table.Columns[c].ColumnName;
            row++;
        }
        foreach (DataRow dr in table.Rows)
        {
            for (int c = 0; c < table.Columns.Count; c++)
            {
                var val = dr[c];
                Worksheet.Cell(row, FromCol + c).Value = val is DBNull ? null : val;
            }
            row++;
        }
        return new ExcelRange(Worksheet, FromRow, FromCol, row - 1, FromCol + table.Columns.Count - 1);
    }

    // ── ToCollection<T> ───────────────────────────────────────────────────────

    public IEnumerable<T> ToCollection<T>(bool hasHeaders = true) where T : new()
    {
        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanWrite).ToDictionary(p => p.Name, StringComparer.OrdinalIgnoreCase);

        int dataStart = hasHeaders ? FromRow + 1 : FromRow;
        string[] headers = new string[ColumnCount];

        if (hasHeaders)
            for (int c = 0; c < ColumnCount; c++)
                headers[c] = Worksheet.GetCell(FromRow, FromCol + c)?.GetString() ?? $"Col{c + 1}";
        else
            for (int c = 0; c < ColumnCount; c++)
                headers[c] = $"Col{c + 1}";

        for (int r = dataStart; r <= ToRow; r++)
        {
            var obj = new T();
            for (int c = 0; c < ColumnCount && c < headers.Length; c++)
            {
                if (!props.TryGetValue(headers[c], out var prop)) continue;
                var cell = Worksheet.GetCell(r, FromCol + c);
                if (cell == null) continue;
                try
                {
                    var val = Convert.ChangeType(cell.Value, prop.PropertyType);
                    prop.SetValue(obj, val);
                }
                catch { /* skip unconvertible */ }
            }
            yield return obj;
        }
    }

    public DataTable ToDataTable(bool hasHeaders = true)
    {
        var dt = new DataTable();
        int dataStart = hasHeaders ? FromRow + 1 : FromRow;

        for (int c = 0; c < ColumnCount; c++)
        {
            var header = hasHeaders
                ? Worksheet.GetCell(FromRow, FromCol + c)?.GetString() ?? $"Col{c + 1}"
                : $"Col{c + 1}";
            dt.Columns.Add(header);
        }

        for (int r = dataStart; r <= ToRow; r++)
        {
            var row = dt.NewRow();
            for (int c = 0; c < ColumnCount; c++)
                row[c] = Worksheet.GetCell(r, FromCol + c)?.Value ?? DBNull.Value;
            dt.Rows.Add(row);
        }
        return dt;
    }

    public List<T> ToList<T>(bool hasHeaders = true) where T : new() =>
        ToCollection<T>(hasHeaders).ToList();

    public object?[,] ToArray()
    {
        var arr = new object?[RowCount, ColumnCount];
        for (int r = 0; r < RowCount; r++)
            for (int c = 0; c < ColumnCount; c++)
                arr[r, c] = Worksheet.GetCell(FromRow + r, FromCol + c)?.Value;
        return arr;
    }

    // ── Range navigation ──────────────────────────────────────────────────────

    public ExcelRange EntireRow => new(Worksheet, FromRow, 1, ToRow, 16384);
    public ExcelRange EntireColumn => new(Worksheet, 1, FromCol, 1048576, ToCol);

    public override string ToString() => Address;
}

[Flags]
public enum ExcelRangeCopyOptions
{
    Values       = 1,
    Formulas     = 2,
    Styles       = 4,
    NumberFormat = 8,
    Comments     = 16,
    Hyperlinks   = 32,
    All          = Values | Formulas | Styles | NumberFormat,
}

public enum ExcelShiftTypeInsert { Down, Right }
public enum ExcelPrintHeaders { None, ColumnCaption, DisplayName }

// ── CellStyleDef clone helper ─────────────────────────────────────────────────

internal static class CellStyleDefExtensions
{
    public static CellStyleDef Clone(this CellStyleDef s) => new()
    {
        NumberFormat = s.NumberFormat,
        WrapText = s.WrapText,
        Locked = s.Locked,
        Hidden = s.Hidden,
        Font = new FontDef
        {
            Name = s.Font.Name, Size = s.Font.Size, Bold = s.Font.Bold,
            Italic = s.Font.Italic, Underline = s.Font.Underline,
            Strikethrough = s.Font.Strikethrough, Color = s.Font.Color,
            ThemeColor = s.Font.ThemeColor, Tint = s.Font.Tint,
        },
        Fill = new FillDef
        {
            PatternType = s.Fill.PatternType,
            BackgroundColor = s.Fill.BackgroundColor,
            ForegroundColor = s.Fill.ForegroundColor,
        },
        Border = new BorderDef
        {
            DiagonalUp = s.Border.DiagonalUp,
            DiagonalDown = s.Border.DiagonalDown,
        },
        Alignment = new AlignmentDef
        {
            Horizontal = s.Alignment.Horizontal,
            Vertical = s.Alignment.Vertical,
            WrapText = s.Alignment.WrapText,
            Indent = s.Alignment.Indent,
            TextRotation = s.Alignment.TextRotation,
            ShrinkToFit = s.Alignment.ShrinkToFit,
        },
    };
}
