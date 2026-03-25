using System.IO.Compression;
using System.Xml;

namespace EPExcel.ML.IO;

/// <summary>OOXML (.xlsx) writer — EPExcel 8.5 parity.</summary>
public sealed class XlsxWriter(ExcelWorkbook workbook)
{
    private readonly ExcelWorkbook _wb = workbook;
    private readonly List<string> _sharedStrings = new();
    private readonly Dictionary<string, int> _ssIndex = new();

    public async Task WriteAsync(Stream output, CancellationToken ct = default)
    {
        using var zip = new ZipArchive(output, ZipArchiveMode.Create, true);
        WriteContentTypes(zip);
        WriteRels(zip);
        WriteWorkbook(zip);
        WriteStyles(zip);
        WriteSharedStrings(zip);
        WriteWorksheets(zip);
        WriteWorkbookRels(zip);
        WriteConnections(zip);
        WriteTheme(zip);
        foreach (var (path, data) in _wb.UnknownParts)
        {
            var entry = zip.CreateEntry(path.TrimStart('/'));
            await using var s = entry.Open();
            await s.WriteAsync(data, ct);
        }
    }

    private void WriteContentTypes(ZipArchive zip)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
        sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
        sb.Append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
        sb.Append("<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
        for (int i = 0; i < _wb.Worksheets.Count; i++)
            sb.Append($"<Override PartName=\"/xl/worksheets/sheet{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        if (_wb.Connections.Connections.Any())
            sb.Append("<Override PartName=\"/xl/connections.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml\"/>");
        if (_wb.IsMacroEnabled)
            sb.Append("<Override PartName=\"/xl/vbaProject.bin\" ContentType=\"application/vnd.ms-office.activeX+xml\"/>");
        sb.Append("</Types>");
        WriteEntry(zip, "[Content_Types].xml", sb.ToString());
    }

    private static void WriteRels(ZipArchive zip) =>
        WriteEntry(zip, "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
            "</Relationships>");

    private void WriteWorkbookRels(ZipArchive zip)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        for (int i = 0; i < _wb.Worksheets.Count; i++)
            sb.Append($"<Relationship Id=\"rId{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i + 1}.xml\"/>");
        int nextRid = _wb.Worksheets.Count + 1;
        sb.Append($"<Relationship Id=\"rId{nextRid++}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
        sb.Append($"<Relationship Id=\"rId{nextRid++}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
        sb.Append($"<Relationship Id=\"rId{nextRid++}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>");
        if (_wb.Connections.Connections.Any())
            sb.Append($"<Relationship Id=\"rId{nextRid}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections\" Target=\"connections.xml\"/>");
        sb.Append("</Relationships>");
        WriteEntry(zip, "xl/_rels/workbook.xml.rels", sb.ToString());
    }

    private void WriteWorkbook(ZipArchive zip)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        sb.Append("<fileVersion appName=\"EPExcel.ML\" lastEdited=\"7\" lowestEdited=\"7\" rupBuild=\"22228\"/>");
        var compat = _wb.Compatibility;
        sb.Append($"<workbookPr date1904=\"{(compat.Use1904DateSystem ? 1 : 0)}\" showInkAnnotation=\"0\" autoCompressPictures=\"0\"/>");
        sb.Append($"<bookViews>{_wb.View.ToOoxml()}</bookViews>");
        sb.Append("<sheets>");
        for (int i = 0; i < _wb.Worksheets.Count; i++)
        {
            var ws = _wb.Worksheets[i];
            sb.Append($"<sheet name=\"{XE(ws.Name)}\" sheetId=\"{i + 1}\" r:id=\"rId{i + 1}\"/>");
        }
        sb.Append("</sheets>");
        if (_wb.NamedRanges.Any())
        {
            sb.Append("<definedNames>");
            foreach (var (name, nr) in _wb.NamedRanges)
                sb.Append($"<definedName name=\"{XE(name)}\">{XE("'" + nr.Worksheet.Name + "'!" + nr.Range.Address)}</definedName>");
            sb.Append("</definedNames>");
        }
        sb.Append($"<calcPr calcId=\"191028\" fullCalcOnLoad=\"{(compat.ForceFullCalcOnLoad ? 1 : 0)}\" calcMode=\"{compat.CalcMode.ToString().ToLowerInvariant()}\" iterate=\"{(compat.Iterate ? 1 : 0)}\" iterateCount=\"{compat.IterateCount}\" iterateDelta=\"{compat.IterateDelta.ToString("R", System.Globalization.CultureInfo.InvariantCulture)}\"/>");
        if (_wb.Protection.LockStructure || _wb.Protection.LockWindows)
        {
            sb.Append($"<workbookProtection lockStructure=\"{(_wb.Protection.LockStructure ? 1 : 0)}\" lockWindows=\"{(_wb.Protection.LockWindows ? 1 : 0)}\"");
            if (_wb.Protection.PasswordHash != null)
                sb.Append($" workbookHashValue=\"{_wb.Protection.PasswordHash}\" workbookSaltValue=\"\" workbookSpinCount=\"100000\" workbookAlgorithmName=\"SHA-512\"");
            sb.Append("/>");
        }
        sb.Append("</workbook>");
        WriteEntry(zip, "xl/workbook.xml", sb.ToString());
    }

    private void WriteStyles(ZipArchive zip)
    {
        var styles = _wb.Styles;
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

        // Number formats
        var custFmts = styles.CustomNumberFormats;
        if (custFmts.Any())
        {
            sb.Append($"<numFmts count=\"{custFmts.Count}\">");
            foreach (var (fmt, id) in custFmts)
                sb.Append($"<numFmt numFmtId=\"{id}\" formatCode=\"{XE(fmt)}\"/>");
            sb.Append("</numFmts>");
        }

        // Fonts
        sb.Append($"<fonts count=\"{styles.AllStyles.Count + 1}\">");
        sb.Append("<font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>");
        foreach (var s in styles.AllStyles)
        {
            sb.Append("<font>");
            if (s.Font.Bold)        sb.Append("<b/>");
            if (s.Font.Italic)      sb.Append("<i/>");
            if (s.Font.Underline)   sb.Append("<u/>");
            if (s.Font.Strikethrough) sb.Append("<strike/>");
            sb.Append($"<sz val=\"{s.Font.Size.ToString(System.Globalization.CultureInfo.InvariantCulture)}\"/>");
            if (s.Font.Color != null)
                sb.Append($"<color rgb=\"{NormColor(s.Font.Color)}\"/>");
            sb.Append($"<name val=\"{XE(s.Font.Name)}\"/>");
            sb.Append("</font>");
        }
        sb.Append("</fonts>");

        // Mapping styles to IDs
        var fillIds = new Dictionary<CellStyleDef, int>();
        var borderIds = new Dictionary<CellStyleDef, int>();

        // Fills
        var allFills = new List<string>();
        allFills.Add("<fill><patternFill patternType=\"none\"/></fill>");
        allFills.Add("<fill><patternFill patternType=\"gray125\"/></fill>");
        foreach (var s in styles.AllStyles)
        {
            if (s.Fill.Gradient != null)
            {
                var grad = s.Fill.Gradient;
                var stops = string.Join("", grad.Stops.Select(st =>
                    $"<a:gs pos=\"{(int)(st.Position * 100000)}\"><a:srgbClr val=\"{st.Color.TrimStart('#')}\"/></a:gs>"));
                fillIds[s] = allFills.Count;
                allFills.Add($"<fill><gradientFill degree=\"{grad.Degree.ToString(System.Globalization.CultureInfo.InvariantCulture)}\">{stops}</gradientFill></fill>");
            }
            else if (s.Fill.PatternType != ExcelFillPattern.None)
            {
                var pt = s.Fill.PatternType.ToString().ToLowerInvariant();
                string bgPart = "", fgPart = "";
                if (s.Fill.PatternType == ExcelFillPattern.Solid)
                {
                    // Solid fills in OOXML use fgColor for the actual color
                    fgPart = s.Fill.BackgroundColor != null ? $"<fgColor rgb=\"{NormColor(s.Fill.BackgroundColor)}\"/>" : "";
                }
                else
                {
                    bgPart = s.Fill.BackgroundColor != null ? $"<bgColor rgb=\"{NormColor(s.Fill.BackgroundColor)}\"/>" : "";
                    fgPart = s.Fill.ForegroundColor != null ? $"<fgColor rgb=\"{NormColor(s.Fill.ForegroundColor)}\"/>" : "";
                }
                fillIds[s] = allFills.Count;
                allFills.Add($"<fill><patternFill patternType=\"{pt}\">{fgPart}{bgPart}</patternFill></fill>");
            }
            else fillIds[s] = 0;
        }
        sb.Append($"<fills count=\"{allFills.Count}\">");
        foreach (var fill in allFills) sb.Append(fill);
        sb.Append("</fills>");

        // Borders
        static string BorderSideXml(string tag, BorderSideDef side)
        {
            if (side.Style == BorderLineStyle.None) return $"<{tag}/>";
            var colorPart = side.Color != null ? $"<color rgb=\"{NormColorStatic(side.Color)}\"/>" : "";
            return $"<{tag} style=\"{side.Style.ToString().ToLowerInvariant()}\">{colorPart}</{tag}>";
        }

        var allBorders = new List<string>();
        allBorders.Add("<border><left/><right/><top/><bottom/><diagonal/></border>");
        foreach (var s in styles.AllStyles)
        {
            bool hasBorder = s.Border.Top.Style    != BorderLineStyle.None ||
                             s.Border.Bottom.Style  != BorderLineStyle.None ||
                             s.Border.Left.Style    != BorderLineStyle.None ||
                             s.Border.Right.Style   != BorderLineStyle.None;
            if (hasBorder)
            {
                var diagAttr = "";
                if (s.Border.DiagonalDown) diagAttr += " diagonalDown=\"1\"";
                if (s.Border.DiagonalUp)   diagAttr += " diagonalUp=\"1\"";
                borderIds[s] = allBorders.Count;
                allBorders.Add($"<border{diagAttr}>{BorderSideXml("left", s.Border.Left)}{BorderSideXml("right", s.Border.Right)}{BorderSideXml("top", s.Border.Top)}{BorderSideXml("bottom", s.Border.Bottom)}{BorderSideXml("diagonal", s.Border.Diagonal)}</border>");
            }
            else borderIds[s] = 0;
        }
        sb.Append($"<borders count=\"{allBorders.Count}\">");
        foreach (var brd in allBorders) sb.Append(brd);
        sb.Append("</borders>");

        // Cell style xfs
        sb.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

        // Cell xfs
        sb.Append($"<cellXfs count=\"{styles.AllStyles.Count + 1}\">");
        sb.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
        int fontIdx = 1;
        foreach (var s in styles.AllStyles)
        {
            int numFmtId = s.NumberFormat != null
                ? (custFmts.TryGetValue(s.NumberFormat, out var nfid) ? nfid : 0) : 0;
            int fillId = fillIds[s];
            int borderId = borderIds[s];

            var alignPart = "";
            if (s.Alignment.Horizontal != ExcelHorizontalAlignment.General ||
                s.Alignment.Vertical   != ExcelVerticalCellAlignment.Bottom ||
                s.WrapText || s.Alignment.TextRotation != 0 || s.Alignment.Indent != 0)
            {
                alignPart  = $"<alignment horizontal=\"{s.Alignment.Horizontal.ToString().ToLowerInvariant()}\"";
                alignPart += $" vertical=\"{s.Alignment.Vertical.ToString().ToLowerInvariant()}\"";
                if (s.WrapText || s.Alignment.WrapText) alignPart += " wrapText=\"1\"";
                if (s.Alignment.TextRotation != 0) alignPart += $" textRotation=\"{s.Alignment.TextRotation}\"";
                if (s.Alignment.Indent != 0) alignPart += $" indent=\"{s.Alignment.Indent}\"";
                alignPart += "/>";
            }

            int applyFont  = s.Font.Bold || s.Font.Italic || s.Font.Color != null || s.Font.Size != 11 ? 1 : 0;
            int applyFill  = fillId > 0 ? 1 : 0;
            int applyBorder = borderId > 0 ? 1 : 0;
            int applyAlign = alignPart.Length > 0 ? 1 : 0;
            sb.Append($"<xf numFmtId=\"{numFmtId}\" fontId=\"{fontIdx}\" fillId=\"{fillId}\" borderId=\"{borderId}\" xfId=\"0\" applyFont=\"{applyFont}\" applyFill=\"{applyFill}\" applyBorder=\"{applyBorder}\" applyAlignment=\"{applyAlign}\">{alignPart}</xf>");
            fontIdx++;
        }
        sb.Append("</cellXfs>");

        sb.Append("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");

        // DXF table for CF
        var allCfRules = _wb.Worksheets.SelectMany(ws => ws.ConditionalFormattings).ToList();
        if (allCfRules.Any())
        {
            sb.Append($"<dxfs count=\"{allCfRules.Count}\">");
            foreach (var rule in allCfRules)
            {
                var rs = rule.Style;
                sb.Append("<dxf>");
                if (rs.Font.Bold || rs.Font.Italic || rs.Font.Color != null)
                {
                    sb.Append("<font>");
                    if (rs.Font.Bold)   sb.Append("<b/>");
                    if (rs.Font.Italic) sb.Append("<i/>");
                    if (rs.Font.Color != null) sb.Append($"<color rgb=\"{NormColor(rs.Font.Color)}\"/>");
                    sb.Append("</font>");
                }
                if (rs.Fill.BackgroundColor != null)
                    sb.Append($"<fill><patternFill><bgColor rgb=\"{NormColor(rs.Fill.BackgroundColor)}\"/></patternFill></fill>");
                sb.Append("</dxf>");
            }
            sb.Append("</dxfs>");
        }

        sb.Append("</styleSheet>");
        WriteEntry(zip, "xl/styles.xml", sb.ToString());
    }

    private void WriteSharedStrings(ZipArchive zip)
    {
        foreach (var ws in _wb.Worksheets)
            foreach (var cell in ws.AllCells())
            {
                var v = cell.DisplayValue;
                if (v is string s && !string.IsNullOrEmpty(s) && !_ssIndex.ContainsKey(s))
                {
                    _ssIndex[s] = _sharedStrings.Count;
                    _sharedStrings.Add(s);
                }
            }

        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{_sharedStrings.Count}\" uniqueCount=\"{_sharedStrings.Count}\">");
        foreach (var str in _sharedStrings)
        {
            var preserve = str.StartsWith(' ') || str.EndsWith(' ') ? " xml:space=\"preserve\"" : "";
            sb.Append($"<si><t{preserve}>{XE(Formulas.FunctionLibrary.SanitizeXml(str))}</t></si>");
        }
        sb.Append("</sst>");
        WriteEntry(zip, "xl/sharedStrings.xml", sb.ToString());
    }

    private void WriteWorksheets(ZipArchive zip)
    {
        for (int i = 0; i < _wb.Worksheets.Count; i++)
        {
            WriteWorksheet(zip, _wb.Worksheets[i], i + 1);
            WriteWorksheetRels(zip, _wb.Worksheets[i], i + 1);
        }
    }

    private void WriteWorksheet(ZipArchive zip, ExcelWorksheet ws, int index)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

        // Sheet properties
        string tabColorPart = ws.TabColor != null ? $"<tabColor rgb=\"{NormColor(ws.TabColor)}\"/>" : "";
        sb.Append($"<sheetPr>{tabColorPart}<outlinePr summaryBelow=\"{(ws.OutlineSymbolsBelow ? 1 : 0)}\" summaryRight=\"{(ws.OutlineSymbolsRight ? 1 : 0)}\"/></sheetPr>");

        if (ws.MaxRow > 0 && ws.MaxCol > 0)
            sb.Append($"<dimension ref=\"{ExcelAddressParser.ToRangeAddress(1, 1, ws.MaxRow, ws.MaxCol)}\"/>");

        // Sheet views
        sb.Append("<sheetViews>");
        sb.Append("<sheetView");
        if (ws.TabSelected) sb.Append(" tabSelected=\"1\"");
        sb.Append($" workbookViewId=\"0\" showGridLines=\"{(ws.ShowGridLines ? 1 : 0)}\" showRowColHeaders=\"{(ws.ShowRowColHeaders ? 1 : 0)}\" zoomScale=\"{(int)ws.View.Zoom}\">");
        if (ws.FreezeRow > 0 || ws.FreezeCol > 0)
        {
            string topLeft = ExcelAddressParser.ToAddress(ws.FreezeRow + 1, ws.FreezeCol + 1);
            sb.Append($"<pane ySplit=\"{ws.FreezeRow}\" xSplit=\"{ws.FreezeCol}\" topLeftCell=\"{topLeft}\" activePane=\"bottomRight\" state=\"frozen\"/>");
            sb.Append($"<selection pane=\"bottomRight\" activeCell=\"{topLeft}\" sqref=\"{topLeft}\"/>");
        }
        else
        {
            sb.Append("<selection activeCell=\"A1\" sqref=\"A1\"/>");
        }
        sb.Append("</sheetView></sheetViews>");

        sb.Append($"<sheetFormatPr defaultRowHeight=\"{ws.DefaultRowHeight}\" defaultColWidth=\"{ws.DefaultColWidth}\" customHeight=\"1\"/>");

        // Column widths
        if (ws.ColWidths.Any())
        {
            sb.Append("<cols>");
            foreach (var kv in ws.ColWidths.OrderBy(kv => kv.Key))
            {
                int col    = kv.Key;
                double width = kv.Value;
                int colOutline = ws.Outline.GetColLevel(col);
                string hiddenAttr   = ws.IsColHidden(col) ? " hidden=\"1\"" : "";
                string outlineAttr  = colOutline > 0 ? $" outlineLevel=\"{colOutline}\"" : "";
                string widthStr     = width.ToString(System.Globalization.CultureInfo.InvariantCulture);
                sb.Append($"<col min=\"{col}\" max=\"{col}\" width=\"{widthStr}\" customWidth=\"1\"{hiddenAttr}{outlineAttr}/>");
            }
            sb.Append("</cols>");
        }

        // Sheet data
        sb.Append("<sheetData>");
        int maxRow = ws.MaxRow;
        for (int r = 1; r <= maxRow; r++)
        {
            bool hasRowData = ws.AllCells().Any(c => c.Row == r);
            if (!hasRowData && !ws.RowHeights.ContainsKey(r)) continue;

            int outlineLv = ws.Outline.GetRowLevel(r);
            sb.Append($"<row r=\"{r}\"");
            if (ws.RowHeights.TryGetValue(r, out var rowH))
                sb.Append($" ht=\"{rowH.ToString(System.Globalization.CultureInfo.InvariantCulture)}\" customHeight=\"1\"");
            if (ws.IsRowHidden(r)) sb.Append(" hidden=\"1\"");
            if (outlineLv > 0) sb.Append($" outlineLevel=\"{outlineLv}\"");
            sb.Append(">");

            foreach (var cell in ws.AllCells().Where(c => c.Row == r).OrderBy(c => c.Col))
            {
                string addr  = ExcelAddressParser.ToAddress(cell.Row, cell.Col);
                string sAttr = cell.StyleIndex > 0 ? $" s=\"{cell.StyleIndex}\"" : "";
                var displayVal = cell.DisplayValue;

                if (cell.Formula != null)
                {
                    string fVal = displayVal is double dv
                        ? $"<v>{dv.ToString("R", System.Globalization.CultureInfo.InvariantCulture)}</v>"
                        : "";
                    sb.Append($"<c r=\"{addr}\"{sAttr} t=\"str\"><f>{XE(cell.Formula)}</f>{fVal}</c>");
                }
                else if (displayVal is double d)
                {
                    sb.Append($"<c r=\"{addr}\"{sAttr}><v>{d.ToString("R", System.Globalization.CultureInfo.InvariantCulture)}</v></c>");
                }
                else if (displayVal is bool b)
                {
                    sb.Append($"<c r=\"{addr}\"{sAttr} t=\"b\"><v>{(b ? 1 : 0)}</v></c>");
                }
                else if (displayVal is CellError ce)
                {
                    sb.Append($"<c r=\"{addr}\"{sAttr} t=\"e\"><v>{XE(ce.ToString())}</v></c>");
                }
                else if (displayVal is string sv && !string.IsNullOrEmpty(sv))
                {
                    int ssi = _ssIndex.TryGetValue(sv, out var idx) ? idx : 0;
                    sb.Append($"<c r=\"{addr}\"{sAttr} t=\"s\"><v>{ssi}</v></c>");
                }
                else if (displayVal != null)
                {
                    sb.Append($"<c r=\"{addr}\"{sAttr} t=\"str\"><v>{XE(displayVal.ToString() ?? "")}</v></c>");
                }
            }
            sb.Append("</row>");
        }
        sb.Append("</sheetData>");

        // Merge cells
        WriteMergeCells(sb, ws);

        // Conditional formatting
        WriteConditionalFormatting(sb, ws);

        // Data validations
        WriteDataValidations(sb, ws);

        // Auto filter
        if (ws.AutoFilterAddress != null)
            sb.Append($"<autoFilter ref=\"{XE(ws.AutoFilterAddress)}\"/>");

        // Sheet protection
        if (ws.Protected)
        {
            sb.Append("<sheetProtection sheet=\"1\"");
            if (ws.PasswordHash != null) sb.Append($" password=\"{ws.PasswordHash}\"");
            sb.Append("/>");
        }

        // Page breaks
        if (ws.PageBreaks.HasRowBreaks) sb.Append(ws.PageBreaks.ToRowXml());
        if (ws.PageBreaks.HasColBreaks) sb.Append(ws.PageBreaks.ToColXml());

        // Page setup
        var ps = ws.PageSetup;
        sb.Append($"<pageSetup paperSize=\"{(int)ps.PaperSize}\" orientation=\"{ps.Orientation.ToString().ToLowerInvariant()}\" copies=\"{ps.Copies}\" scale=\"{ps.Scale}\" fitToWidth=\"{ps.FitToWidth}\" fitToHeight=\"{ps.FitToHeight}\" r:id=\"rId1\"/>");

        // Header/footer
        if (!string.IsNullOrEmpty(ps.OddHeader) || !string.IsNullOrEmpty(ps.OddFooter))
        {
            sb.Append("<headerFooter>");
            if (!string.IsNullOrEmpty(ps.OddHeader)) sb.Append($"<oddHeader>{XE(ps.OddHeader)}</oddHeader>");
            if (!string.IsNullOrEmpty(ps.OddFooter)) sb.Append($"<oddFooter>{XE(ps.OddFooter)}</oddFooter>");
            sb.Append("</headerFooter>");
        }

        // Tables
        if (ws.Tables.Any())
        {
            sb.Append($"<tableParts count=\"{ws.Tables.Count}\">");
            for (int ti = 0; ti < ws.Tables.Count; ti++)
                sb.Append($"<tablePart r:id=\"rId{10 + ti}\"/>");
            sb.Append("</tableParts>");
        }

        // Sparklines
        if (ws.SparklineGroups.Any())
        {
            sb.Append("<extLst><ext uri=\"{05C60535-1F16-4fd2-B633-E4A46CF9E7E1}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">");
            sb.Append("<x14:sparklineGroups xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">");
            foreach (var sg in ws.SparklineGroups)
            {
                sb.Append($"<x14:sparklineGroup type=\"{sg.Type.ToString().ToLowerInvariant()}\" displayEmptyCellsAs=\"gap\">");
                sb.Append($"<x14:sparklines><x14:sparkline><xm:f>{XE(sg.DataRange)}</xm:f><xm:sqref>{XE(sg.LocationRange)}</xm:sqref></x14:sparkline></x14:sparklines>");
                sb.Append("</x14:sparklineGroup>");
            }
            sb.Append("</x14:sparklineGroups></ext></extLst>");
        }

        sb.Append("</worksheet>");
        WriteEntry(zip, $"xl/worksheets/sheet{index}.xml", sb.ToString());

        for (int ti = 0; ti < ws.Tables.Count; ti++)
            WriteTable(zip, ws.Tables[ti], index, ti + 1);
    }

    private static void WriteMergeCells(StringBuilder sb, ExcelWorksheet ws)
    {
        if (!ws.MergedCells.Any()) return;
        sb.Append($"<mergeCells count=\"{ws.MergedCells.Count}\">");
        foreach (var m in ws.MergedCells) sb.Append($"<mergeCell ref=\"{XE(m)}\"/>");
        sb.Append("</mergeCells>");
    }

    private static void WriteConditionalFormatting(StringBuilder sb, ExcelWorksheet ws)
    {
        if (!ws.ConditionalFormattings.Any()) return;
        var byAddr = ws.ConditionalFormattings.GroupBy(r => r.Address);
        foreach (var grp in byAddr)
        {
            sb.Append($"<conditionalFormatting sqref=\"{XE(grp.Key)}\">");
            int dxfIdx = 0;
            foreach (var rule in grp.OrderBy(r => r.Priority))
            {
                string typeStr = rule.Type switch
                {
                    ConditionalFormattingType.CellValue     => "cellIs",
                    ConditionalFormattingType.Expression    => "expression",
                    ConditionalFormattingType.ColorScale    => "colorScale",
                    ConditionalFormattingType.DataBar       => "dataBar",
                    ConditionalFormattingType.IconSet       => "iconSet",
                    ConditionalFormattingType.Top10         => "top10",
                    ConditionalFormattingType.DuplicateValues => "duplicateValues",
                    ConditionalFormattingType.UniqueValues  => "uniqueValues",
                    ConditionalFormattingType.ContainsText  => "containsText",
                    _                                       => "expression"
                };
                sb.Append($"<cfRule type=\"{typeStr}\" dxfId=\"{dxfIdx++}\" priority=\"{rule.Priority}\" stopIfTrue=\"{(rule.StopIfTrue ? 1 : 0)}\"");
                if (rule.Operator != null) sb.Append($" operator=\"{rule.Operator}\"");
                sb.Append(">");
                if (rule.Value1 != null) sb.Append($"<formula>{XE(FormatCfValue(rule.Value1))}</formula>");
                if (rule.Value2 != null) sb.Append($"<formula>{XE(FormatCfValue(rule.Value2))}</formula>");
                if (rule.Text   != null) sb.Append($"<formula>NOT(ISERROR(SEARCH(\"{XE(rule.Text)}\",A1)))</formula>");
                if (rule.ColorScale != null)
                {
                    var cs = rule.ColorScale;
                    sb.Append($"<colorScale><cfvo type=\"{cs.MinType}\"/><cfvo type=\"{cs.MidType}\" val=\"{cs.MidValue.ToString(System.Globalization.CultureInfo.InvariantCulture)}\"/><cfvo type=\"{cs.MaxType}\"/><color rgb=\"{cs.MinColor}\"/><color rgb=\"{cs.MidColor}\"/><color rgb=\"{cs.MaxColor}\"/></colorScale>");
                }
                if (rule.DataBar != null)
                    sb.Append($"<dataBar showValue=\"{(rule.DataBar.ShowValue ? 1 : 0)}\"><cfvo type=\"{rule.DataBar.MinType}\"/><cfvo type=\"{rule.DataBar.MaxType}\"/><color rgb=\"{rule.DataBar.Color}\"/></dataBar>");
                sb.Append("</cfRule>");
            }
            sb.Append("</conditionalFormatting>");
        }
    }

    private static void WriteDataValidations(StringBuilder sb, ExcelWorksheet ws)
    {
        if (!ws.DataValidations.Any()) return;
        sb.Append($"<dataValidations count=\"{ws.DataValidations.Count}\">");
        foreach (var dv in ws.DataValidations)
        {
            string dvType = dv.Type switch
            {
                DataValidationType.TextLength => "textLength",
                DataValidationType.None       => "none",
                _                             => dv.Type.ToString().ToLowerInvariant()
            };
            string dvOp  = dv.Operator.ToString();
            // convert enum to camelCase Excel operator string
            string dvOpStr = dvOp.Length > 0
                ? char.ToLowerInvariant(dvOp[0]) + dvOp.Substring(1)
                : dvOp;

            sb.Append($"<dataValidation type=\"{dvType}\" sqref=\"{XE(dv.Address)}\" allowBlank=\"{(dv.AllowBlank ? 1 : 0)}\" showInputMessage=\"{(dv.ShowInputMessage ? 1 : 0)}\" showErrorAlert=\"{(dv.ShowErrorAlert ? 1 : 0)}\" errorStyle=\"{dv.ErrorStyle.ToString().ToLowerInvariant()}\" operator=\"{dvOpStr}\"");
            if (dv.PromptTitle != null) sb.Append($" promptTitle=\"{XE(dv.PromptTitle)}\"");
            if (dv.Prompt      != null) sb.Append($" prompt=\"{XE(dv.Prompt)}\"");
            if (dv.ErrorTitle  != null) sb.Append($" errorTitle=\"{XE(dv.ErrorTitle)}\"");
            if (dv.Error       != null) sb.Append($" error=\"{XE(dv.Error)}\"");
            if (dv.Type == DataValidationType.List && dv.InCellDropdown) sb.Append(" showDropDown=\"0\"");
            sb.Append(">");
            if (dv.Formula1 != null) sb.Append($"<formula1>{XE(dv.Formula1)}</formula1>");
            if (dv.Formula2 != null) sb.Append($"<formula2>{XE(dv.Formula2)}</formula2>");
            sb.Append("</dataValidation>");
        }
        sb.Append("</dataValidations>");
    }

    private static void WriteWorksheetRels(ZipArchive zip, ExcelWorksheet ws, int index)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        sb.Append("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"\" TargetMode=\"External\"/>");
        for (int ti = 0; ti < ws.Tables.Count; ti++)
            sb.Append($"<Relationship Id=\"rId{10 + ti}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table{index}_{ti + 1}.xml\"/>");
        sb.Append("</Relationships>");
        WriteEntry(zip, $"xl/worksheets/_rels/sheet{index}.xml.rels", sb.ToString());
    }

    private static void WriteTable(ZipArchive zip, ExcelTable table, int wsIdx, int tableIdx)
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        int id = wsIdx * 100 + tableIdx;
        string displayName = XE(table.DisplayName ?? table.Name);
        sb.Append($"<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"{id}\" name=\"{XE(table.Name)}\" displayName=\"{displayName}\" ref=\"{XE(table.Address)}\" headerRowCount=\"{(table.ShowHeader ? 1 : 0)}\" totalsRowCount=\"{(table.ShowTotals ? 1 : 0)}\" insertRow=\"0\" insertRowShift=\"0\">");
        if (!string.IsNullOrEmpty(table.StyleName))
            sb.Append($"<tableStyleInfo name=\"{XE(table.StyleName)}\" showFirstColumn=\"{(table.ShowFirstColumn ? 1 : 0)}\" showLastColumn=\"{(table.ShowLastColumn ? 1 : 0)}\" showRowStripes=\"{(table.ShowRowStripes ? 1 : 0)}\" showColumnStripes=\"{(table.ShowColumnStripes ? 1 : 0)}\"/>");
        if (table.ShowFilter)
            sb.Append($"<autoFilter ref=\"{XE(table.Address)}\"/>");
        sb.Append($"<tableColumns count=\"{Math.Max(1, table.Columns.Count)}\">");
        if (table.Columns.Any())
        {
            int cid = 1;
            foreach (var col in table.Columns)
            {
                sb.Append($"<tableColumn id=\"{cid++}\" name=\"{XE(col.Name)}\">");
                if (col.Formula != null) sb.Append($"<calculatedColumnFormula>{XE(col.Formula)}</calculatedColumnFormula>");
                sb.Append("</tableColumn>");
            }
        }
        else
        {
            sb.Append("<tableColumn id=\"1\" name=\"Column1\"/>");
        }
        sb.Append("</tableColumns></table>");
        WriteEntry(zip, $"xl/tables/table{wsIdx}_{tableIdx}.xml", sb.ToString());
    }

    private void WriteTheme(ZipArchive zip)
    {
        var theme = _wb.Theme;
        string c(int i) => theme.Colors.Length > i ? theme.Colors[i] : "000000";
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"{XE(theme.Name)}\">");
        sb.Append("<a:themeElements><a:clrScheme name=\"Office\">");
        sb.Append("<a:dk1><a:sysClr lastClr=\"000000\" val=\"windowText\"/></a:dk1>");
        sb.Append("<a:lt1><a:sysClr lastClr=\"FFFFFF\" val=\"window\"/></a:lt1>");
        sb.Append($"<a:dk2><a:srgbClr val=\"{c(3)}\"/></a:dk2>");
        sb.Append($"<a:lt2><a:srgbClr val=\"{c(2)}\"/></a:lt2>");
        sb.Append($"<a:accent1><a:srgbClr val=\"{c(4)}\"/></a:accent1>");
        sb.Append($"<a:accent2><a:srgbClr val=\"{c(5)}\"/></a:accent2>");
        sb.Append($"<a:accent3><a:srgbClr val=\"{c(6)}\"/></a:accent3>");
        sb.Append($"<a:accent4><a:srgbClr val=\"{c(7)}\"/></a:accent4>");
        sb.Append($"<a:accent5><a:srgbClr val=\"{c(8)}\"/></a:accent5>");
        sb.Append($"<a:accent6><a:srgbClr val=\"{c(9)}\"/></a:accent6>");
        sb.Append("<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink>");
        sb.Append("<a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>");
        sb.Append("</a:clrScheme>");
        sb.Append("<a:fontScheme name=\"Office\">");
        sb.Append($"<a:majorFont><a:latin typeface=\"{XE(theme.Fonts.HeadingLatin)}\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>");
        sb.Append($"<a:minorFont><a:latin typeface=\"{XE(theme.Fonts.BodyLatin)}\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>");
        sb.Append("</a:fontScheme>");
        sb.Append("<a:fmtScheme name=\"Office\">");
        sb.Append("<a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>");
        sb.Append("<a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>");
        sb.Append("<a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst>");
        sb.Append("<a:lnStyleLst><a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln></a:lnStyleLst>");
        sb.Append("<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>");
        sb.Append("<a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:bgFillStyleLst>");
        sb.Append("</a:fmtScheme></a:themeElements></a:theme>");
        WriteEntry(zip, "xl/theme/theme1.xml", sb.ToString());
    }

    private void WriteConnections(ZipArchive zip)
    {
        if (!_wb.Connections.Connections.Any()) return;
        WriteEntry(zip, "xl/connections.xml", _wb.Connections.ToXml());
    }

    private static void WriteEntry(ZipArchive zip, string path, string content)
    {
        var entry = zip.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        var bytes = Encoding.UTF8.GetBytes(content);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static string XE(string? s) =>
        (s ?? "")
            .Replace("&",  "&amp;")
            .Replace("<",  "&lt;")
            .Replace(">",  "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'",  "&apos;");

    private static string NormColor(string? c)
    {
        if (c == null) return "FF000000";
        var h = c.TrimStart('#');
        if (h.Length == 6)  return "FF" + h.ToUpperInvariant();
        if (h.Length == 8)  return h.ToUpperInvariant();
        return "FF000000";
    }

    private static string NormColorStatic(string? c) => NormColor(c);

    private static string FormatCfValue(object? v)
    {
        if (v is double d) return d.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
        return v?.ToString() ?? "";
    }
}
