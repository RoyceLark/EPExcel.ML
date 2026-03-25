namespace EPExcel.ML.IO;

/// <summary>
/// Live pivot table calculation engine — EPExcel 7.2+ parity.
/// Calculates pivot tables from source data, enabling GETPIVOTDATA formula resolution.
/// </summary>
public sealed class PivotCalculationEngine
{
    private readonly ExcelWorkbook _workbook;

    public PivotCalculationEngine(ExcelWorkbook workbook) => _workbook = workbook;

    public void CalculateAll(bool refreshCache = true)
    {
        foreach (var ws in _workbook.Worksheets)
            foreach (var pivot in ws.PivotTables)
                Calculate(pivot, _workbook, refreshCache);
    }

    public void Calculate(ExcelPivotTable pivot, ExcelWorkbook workbook, bool refreshCache = true)
    {
        if (refreshCache) RefreshCache(pivot, workbook);
        BuildData(pivot);
        pivot.IsCalculated = true;
    }

    private static void RefreshCache(ExcelPivotTable pivot, ExcelWorkbook workbook)
    {
        var srcWs = workbook.GetWorksheet(pivot.DataSheetName);
        if (srcWs == null) return;
        try
        {
            var (fr, fc, tr, tc) = ExcelAddressParser.ParseRange(pivot.DataRange);
            pivot.Fields.Clear();
            for (int c = fc; c <= tc; c++)
            {
                var header = srcWs.GetCell(fr, c)?.GetString()?.Trim() ?? $"Field{c - fc + 1}";
                var field = new ExcelPivotField { Name = header };
                for (int r = fr + 1; r <= tr; r++)
                {
                    var val = srcWs.GetCell(r, c)?.DisplayValue;
                    field.Values.Add(val);
                    if (val is double or int or long) field.IsNumeric = true;
                }
                field.UniqueValues.AddRange(
                    field.Values.Select(v => v?.ToString()).Where(v => v != null)
                        .Distinct().OrderBy(v => v).ToList()!);
                pivot.Fields.Add(field);
            }
        }
        catch { }
    }

    private static void BuildData(ExcelPivotTable pivot)
    {
        pivot.CalculatedData.Clear();
        foreach (var df in pivot.DataFields)
        {
            int fi = pivot.Fields.FindIndex(f => f.Name == df.FieldName);
            if (fi < 0) continue;
            var vals = pivot.Fields[fi].Values;
            var rowIdxs = pivot.RowFields.Select(rf => pivot.Fields.FindIndex(f => f.Name == rf.FieldName)).Where(i => i >= 0).ToList();
            var colIdxs = pivot.ColFields.Select(cf => pivot.Fields.FindIndex(f => f.Name == cf.FieldName)).Where(i => i >= 0).ToList();

            for (int row = 0; row < vals.Count; row++)
            {
                var key = new PivotDataKey(
                    df.FieldName,
                    rowIdxs.Select(i => pivot.Fields[i].Values.Count > row ? pivot.Fields[i].Values[row]?.ToString() ?? "" : "").ToList(),
                    colIdxs.Select(i => pivot.Fields[i].Values.Count > row ? pivot.Fields[i].Values[row]?.ToString() ?? "" : "").ToList()
                );
                if (!pivot.CalculatedData.TryGetValue(key, out var acc))
                {
                    acc = new PivotAccumulator();
                    pivot.CalculatedData[key] = acc;
                }
                acc.Accumulate(Formulas.FunctionLibrary.Num(vals[row]), df.Function);
            }
        }
    }

    public static double QueryPivot(ExcelPivotTable pivot, string dataField,
        string[] fieldNames, string[] fieldValues)
    {
        if (!pivot.CalculatedData.Any()) return 0;
        var criteria = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < Math.Min(fieldNames.Length, fieldValues.Length); i++)
            criteria[fieldNames[i]] = fieldValues[i];

        double total = 0; bool found = false;
        var rowFields = pivot.RowFields.Select(rf => rf.FieldName).ToList();
        var colFields = pivot.ColFields.Select(cf => cf.FieldName).ToList();

        foreach (var (key, acc) in pivot.CalculatedData)
        {
            if (!key.DataField.Equals(dataField, StringComparison.OrdinalIgnoreCase)) continue;
            bool match = true;
            for (int i = 0; i < rowFields.Count && match; i++)
                if (criteria.TryGetValue(rowFields[i], out var exp))
                    if (i >= key.RowValues.Count || !key.RowValues[i].Equals(exp, StringComparison.OrdinalIgnoreCase))
                        match = false;
            for (int i = 0; i < colFields.Count && match; i++)
                if (criteria.TryGetValue(colFields[i], out var exp))
                    if (i >= key.ColValues.Count || !key.ColValues[i].Equals(exp, StringComparison.OrdinalIgnoreCase))
                        match = false;
            if (match) { total += acc.Result; found = true; }
        }
        return found ? total : 0;
    }
}
