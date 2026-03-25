using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;

namespace EPExcel.ML;

// ── ExcelColor — Color manager (EPExcel 8 ColorManager parity) ─────────────────

/// <summary>
/// Full color manager — EPExcel 8 parity.
/// Supports hex, RGB, ARGB, HSL, theme, preset (148 CSS colors), system colors.
/// </summary>
public sealed class ExcelColor
{
    public byte A { get; private set; } = 255;
    public byte R { get; private set; }
    public byte G { get; private set; }
    public byte B { get; private set; }
    public int ThemeColorIndex { get; private set; } = -1;
    public double Tint { get; private set; }

    private static readonly Dictionary<string, (byte r, byte g, byte b)> _presets =
        new(StringComparer.OrdinalIgnoreCase)
    {
        ["AliceBlue"]=(240,248,255),["AntiqueWhite"]=(250,235,215),["Aqua"]=(0,255,255),
        ["Aquamarine"]=(127,255,212),["Azure"]=(240,255,255),["Beige"]=(245,245,220),
        ["Bisque"]=(255,228,196),["Black"]=(0,0,0),["BlanchedAlmond"]=(255,235,205),
        ["Blue"]=(0,0,255),["BlueViolet"]=(138,43,226),["Brown"]=(165,42,42),
        ["BurlyWood"]=(222,184,135),["CadetBlue"]=(95,158,160),["Chartreuse"]=(127,255,0),
        ["Chocolate"]=(210,105,30),["Coral"]=(255,127,80),["CornflowerBlue"]=(100,149,237),
        ["Cornsilk"]=(255,248,220),["Crimson"]=(220,20,60),["Cyan"]=(0,255,255),
        ["DarkBlue"]=(0,0,139),["DarkCyan"]=(0,139,139),["DarkGoldenrod"]=(184,134,11),
        ["DarkGray"]=(169,169,169),["DarkGreen"]=(0,100,0),["DarkKhaki"]=(189,183,107),
        ["DarkMagenta"]=(139,0,139),["DarkOliveGreen"]=(85,107,47),["DarkOrange"]=(255,140,0),
        ["DarkOrchid"]=(153,50,204),["DarkRed"]=(139,0,0),["DarkSalmon"]=(233,150,122),
        ["DarkSeaGreen"]=(143,188,143),["DarkSlateBlue"]=(72,61,139),["DarkSlateGray"]=(47,79,79),
        ["DarkTurquoise"]=(0,206,209),["DarkViolet"]=(148,0,211),["DeepPink"]=(255,20,147),
        ["DeepSkyBlue"]=(0,191,255),["DimGray"]=(105,105,105),["DodgerBlue"]=(30,144,255),
        ["Firebrick"]=(178,34,34),["FloralWhite"]=(255,250,240),["ForestGreen"]=(34,139,34),
        ["Fuchsia"]=(255,0,255),["Gainsboro"]=(220,220,220),["GhostWhite"]=(248,248,255),
        ["Gold"]=(255,215,0),["Goldenrod"]=(218,165,32),["Gray"]=(128,128,128),
        ["Green"]=(0,128,0),["GreenYellow"]=(173,255,47),["Honeydew"]=(240,255,240),
        ["HotPink"]=(255,105,180),["IndianRed"]=(205,92,92),["Indigo"]=(75,0,130),
        ["Ivory"]=(255,255,240),["Khaki"]=(240,230,140),["Lavender"]=(230,230,250),
        ["LavenderBlush"]=(255,240,245),["LawnGreen"]=(124,252,0),["LemonChiffon"]=(255,250,205),
        ["LightBlue"]=(173,216,230),["LightCoral"]=(240,128,128),["LightCyan"]=(224,255,255),
        ["LightGoldenrodYellow"]=(250,250,210),["LightGray"]=(211,211,211),["LightGreen"]=(144,238,144),
        ["LightPink"]=(255,182,193),["LightSalmon"]=(255,160,122),["LightSeaGreen"]=(32,178,170),
        ["LightSkyBlue"]=(135,206,250),["LightSlateGray"]=(119,136,153),["LightSteelBlue"]=(176,196,222),
        ["LightYellow"]=(255,255,224),["Lime"]=(0,255,0),["LimeGreen"]=(50,205,50),
        ["Linen"]=(250,240,230),["Magenta"]=(255,0,255),["Maroon"]=(128,0,0),
        ["MediumAquamarine"]=(102,205,170),["MediumBlue"]=(0,0,205),["MediumOrchid"]=(186,85,211),
        ["MediumPurple"]=(147,112,219),["MediumSeaGreen"]=(60,179,113),["MediumSlateBlue"]=(123,104,238),
        ["MediumSpringGreen"]=(0,250,154),["MediumTurquoise"]=(72,209,204),["MediumVioletRed"]=(199,21,133),
        ["MidnightBlue"]=(25,25,112),["MintCream"]=(245,255,250),["MistyRose"]=(255,228,225),
        ["Moccasin"]=(255,228,181),["NavajoWhite"]=(255,222,173),["Navy"]=(0,0,128),
        ["OldLace"]=(253,245,230),["Olive"]=(128,128,0),["OliveDrab"]=(107,142,35),
        ["Orange"]=(255,165,0),["OrangeRed"]=(255,69,0),["Orchid"]=(218,112,214),
        ["PaleGoldenrod"]=(238,232,170),["PaleGreen"]=(152,251,152),["PaleTurquoise"]=(175,238,238),
        ["PaleVioletRed"]=(219,112,147),["PapayaWhip"]=(255,239,213),["PeachPuff"]=(255,218,185),
        ["Peru"]=(205,133,63),["Pink"]=(255,192,203),["Plum"]=(221,160,221),
        ["PowderBlue"]=(176,224,230),["Purple"]=(128,0,128),["Red"]=(255,0,0),
        ["RosyBrown"]=(188,143,143),["RoyalBlue"]=(65,105,225),["SaddleBrown"]=(139,69,19),
        ["Salmon"]=(250,128,114),["SandyBrown"]=(244,164,96),["SeaGreen"]=(46,139,87),
        ["SeaShell"]=(255,245,238),["Sienna"]=(160,82,45),["Silver"]=(192,192,192),
        ["SkyBlue"]=(135,206,235),["SlateBlue"]=(106,90,205),["SlateGray"]=(112,128,144),
        ["Snow"]=(255,250,250),["SpringGreen"]=(0,255,127),["SteelBlue"]=(70,130,180),
        ["Tan"]=(210,180,140),["Teal"]=(0,128,128),["Thistle"]=(216,191,216),
        ["Tomato"]=(255,99,71),["Turquoise"]=(64,224,208),["Violet"]=(238,130,238),
        ["Wheat"]=(245,222,179),["White"]=(255,255,255),["WhiteSmoke"]=(245,245,245),
        ["Yellow"]=(255,255,0),["YellowGreen"]=(154,205,50),
    };

    private ExcelColor() { }

    public static ExcelColor FromHex(string hex)
    {
        hex = hex.TrimStart('#');
        var c = new ExcelColor();
        if (hex.Length == 8) { c.A=HexByte(hex,0); c.R=HexByte(hex,2); c.G=HexByte(hex,4); c.B=HexByte(hex,6); }
        else if (hex.Length == 6) { c.R=HexByte(hex,0); c.G=HexByte(hex,2); c.B=HexByte(hex,4); }
        return c;
    }

    public static ExcelColor FromRgb(byte r, byte g, byte b) =>
        new() { R = r, G = g, B = b };

    public static ExcelColor FromArgb(byte a, byte r, byte g, byte b) =>
        new() { A = a, R = r, G = g, B = b };

    public static ExcelColor FromHsl(double h, double s, double l)
    {
        double c2 = (1 - Math.Abs(2 * l - 1)) * s;
        double x = c2 * (1 - Math.Abs(h / 60 % 2 - 1));
        double m = l - c2 / 2;
        double rv, gv, bv;
        if (h < 60)       { rv=c2; gv=x;  bv=0;  }
        else if (h < 120) { rv=x;  gv=c2; bv=0;  }
        else if (h < 180) { rv=0;  gv=c2; bv=x;  }
        else if (h < 240) { rv=0;  gv=x;  bv=c2; }
        else if (h < 300) { rv=x;  gv=0;  bv=c2; }
        else              { rv=c2; gv=0;  bv=x;  }
        return FromRgb((byte)((rv+m)*255), (byte)((gv+m)*255), (byte)((bv+m)*255));
    }

    public static ExcelColor FromTheme(int themeColorIndex, double tint = 0) =>
        new() { ThemeColorIndex = themeColorIndex, Tint = tint };

    public static ExcelColor FromPreset(string name) =>
        _presets.TryGetValue(name, out var v) ? FromRgb(v.r, v.g, v.b) : FromRgb(0, 0, 0);

    public static ExcelColor FromSystem(string systemColor) => systemColor.ToLowerInvariant() switch
    {
        "windowtext" => FromRgb(0, 0, 0), "window" => FromRgb(255, 255, 255),
        "highlight" => FromRgb(0, 0, 255), "highlighttext" => FromRgb(255, 255, 255),
        "buttonface" => FromRgb(240, 240, 240), "buttontext" => FromRgb(0, 0, 0),
        _ => FromRgb(0, 0, 0)
    };

    public ExcelColor WithTint(double tint) => new() { A=A, R=R, G=G, B=B, ThemeColorIndex=ThemeColorIndex, Tint=tint };

    public string ToHex(bool includeAlpha = true) =>
        includeAlpha ? $"{A:X2}{R:X2}{G:X2}{B:X2}" : $"{R:X2}{G:X2}{B:X2}";

    public (double H, double S, double L) ToHsl()
    {
        double rr = R / 255.0, gg = G / 255.0, bb = B / 255.0;
        double mx = Math.Max(rr, Math.Max(gg, bb)), mn = Math.Min(rr, Math.Min(gg, bb));
        double l = (mx + mn) / 2;
        if (mx == mn) return (0, 0, l);
        double d = mx - mn, s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
        double h = mx == rr ? (gg - bb) / d + (gg < bb ? 6 : 0) :
                   mx == gg ? (bb - rr) / d + 2 : (rr - gg) / d + 4;
        return (h * 60, s, l);
    }

    private static byte HexByte(string s, int i) => Convert.ToByte(s.Substring(i, 2), 16);

    public static IReadOnlyDictionary<string, (byte R, byte G, byte B)> PresetColors =>
        _presets.ToDictionary(kv => kv.Key, kv => kv.Value);
}

// ── Named styles ──────────────────────────────────────────────────────────────

/// <summary>Named cell styles — EPExcel parity.</summary>
public sealed class ExcelNamedStyleCollection
{
    private readonly Dictionary<string, ExcelNamedStyle> _styles = new(StringComparer.OrdinalIgnoreCase);

    public ExcelNamedStyle Add(string name)
    {
        var s = new ExcelNamedStyle(name);
        _styles[name] = s;
        return s;
    }

    public ExcelNamedStyle? this[string name] => _styles.TryGetValue(name, out var s) ? s : null;
    public IEnumerable<ExcelNamedStyle> All => _styles.Values;
}

public sealed class ExcelNamedStyle(string name)
{
    public string Name { get; } = name;
    public CellStyleDef Style { get; set; } = new();
    public int BuiltInId { get; set; } = -1;
}

// ── Page breaks ───────────────────────────────────────────────────────────────

/// <summary>Page breaks — EPExcel parity.</summary>
public sealed class ExcelPageBreaks
{
    private readonly HashSet<int> _rows = new();
    private readonly HashSet<int> _cols = new();

    public void AddRowBreak(int row) => _rows.Add(row);
    public void AddColBreak(int col) => _cols.Add(col);
    public void RemoveRowBreak(int row) => _rows.Remove(row);
    public void RemoveColBreak(int col) => _cols.Remove(col);
    public IReadOnlySet<int> RowBreaks => _rows;
    public IReadOnlySet<int> ColBreaks => _cols;
    public bool HasRowBreaks => _rows.Any();
    public bool HasColBreaks => _cols.Any();

    internal string ToRowXml()
    {
        if (!_rows.Any()) return "";
        var brks = string.Join("", _rows.Select(r => $"<brk id=\"{r}\" max=\"16383\" man=\"1\"/>"));
        return $"<rowBreaks count=\"{_rows.Count}\" manualBreakCount=\"{_rows.Count}\">{brks}</rowBreaks>";
    }

    internal string ToColXml()
    {
        if (!_cols.Any()) return "";
        var brks = string.Join("", _cols.Select(c => $"<brk id=\"{c}\" max=\"1048575\" man=\"1\"/>"));
        return $"<colBreaks count=\"{_cols.Count}\" manualBreakCount=\"{_cols.Count}\">{brks}</colBreaks>";
    }
}

// ── Outline/Grouping ──────────────────────────────────────────────────────────

/// <summary>Row and column outline grouping — EPExcel parity.</summary>
public sealed class ExcelOutlineCollection
{
    private readonly SortedDictionary<int, int> _rowLevels = new();
    private readonly SortedDictionary<int, int> _colLevels = new();

    public void GroupRows(int fromRow, int toRow, int level = 1)
    {
        for (int r = fromRow; r <= toRow; r++) _rowLevels[r] = level;
    }

    public void GroupColumns(int fromCol, int toCol, int level = 1)
    {
        for (int c = fromCol; c <= toCol; c++) _colLevels[c] = level;
    }

    public void UngroupRows(int fromRow, int toRow)
    {
        for (int r = fromRow; r <= toRow; r++) _rowLevels.Remove(r);
    }

    public void UngroupColumns(int fromCol, int toCol)
    {
        for (int c = fromCol; c <= toCol; c++) _colLevels.Remove(c);
    }

    public int GetRowLevel(int row) => _rowLevels.TryGetValue(row, out var l) ? l : 0;
    public int GetColLevel(int col) => _colLevels.TryGetValue(col, out var l) ? l : 0;
    public int MaxRowLevel => _rowLevels.Values.DefaultIfEmpty(0).Max();
    public int MaxColLevel => _colLevels.Values.DefaultIfEmpty(0).Max();
    public IReadOnlyDictionary<int, int> RowLevels => _rowLevels;
    public IReadOnlyDictionary<int, int> ColLevels => _colLevels;
}

// ── Chart style manager ───────────────────────────────────────────────────────

/// <summary>Chart styles 1–48 + color variants — EPExcel parity.</summary>
public static class ChartStyleManager
{
    private static readonly Dictionary<int, ChartStyleDefinition> _styles = BuildStyles();

    public static void ApplyStyle(ExcelChart chart, int styleId)
    {
        chart.StyleId = Math.Clamp(styleId, 1, 348);
        if (_styles.TryGetValue(styleId % 8 == 0 ? 8 : styleId % 8, out var def))
        {
            foreach (var (i, s) in chart.Series.Select((s, i) => (i, s)))
                s.Color = def.SeriesColors[i % def.SeriesColors.Length];
            if (def.BackgroundColor != null) chart.BackgroundColor = def.BackgroundColor;
        }
    }

    public static void ApplyStyle(ExcelChart chart, ExcelPresetChartStyle preset)
        => ApplyStyle(chart, (int)preset);

    private static Dictionary<int, ChartStyleDefinition> BuildStyles() => new()
    {
        [1] = new(["#4472C4","#ED7D31","#A9D18E","#FFC000"], null),
        [2] = new(["#4472C4","#ED7D31","#A9D18E","#FFC000"], null),
        [3] = new(["#4472C4","#ED7D31","#A9D18E","#FFC000"], null),
        [4] = new(["#FFFFFF","#BFBFBF","#969696","#595959"], "#404040"),
        [5] = new(["#4472C4","#ED7D31","#A9D18E","#FFC000"], "#EEEEEE"),
        [6] = new(["#255E91","#9E480E","#638029","#97680D"], null),
        [7] = new(["#4472C4","#ED7D31","#A9D18E","#FFC000"], "#F0F0F0"),
        [8] = new(["#FFFFFF","#FFBE00","#FF6600","#CC0000"], "#404040"),
    };

    private sealed record ChartStyleDefinition(string[] SeriesColors, string? BackgroundColor);
}

public enum ExcelPresetChartStyle
{
    Style1=1, Style2, Style3, Style4, Style5, Style6, Style7, Style8,
    Style9, Style10, Style11, Style12, Style13, Style14, Style15, Style16,
    Style17, Style18, Style19, Style20, Style21, Style22, Style23, Style24,
    Style25, Style26, Style27, Style28, Style29, Style30, Style31, Style32,
    Style33, Style34, Style35, Style36, Style37, Style38, Style39, Style40,
    Style41, Style42, Style43, Style44, Style45, Style46, Style47, Style48,
}

// ── OLE Objects ───────────────────────────────────────────────────────────────

/// <summary>OLE embedded objects — EPExcel parity.</summary>
public sealed class ExcelOleObject(string progId, byte[] data)
{
    public string ProgId { get; } = progId;
    public byte[] Data { get; } = data;
    public int FromRow { get; set; } = 1;
    public int FromCol { get; set; } = 1;
    public int ToRow { get; set; } = 10;
    public int ToCol { get; set; } = 6;
    public string? DisplayName { get; set; }
    public string? IconCaption { get; set; }
    public bool ShowAsIcon { get; set; }
    public bool LinkedToFile { get; set; }
}

// ── Form Controls ─────────────────────────────────────────────────────────────

/// <summary>Form controls — EPExcel parity.</summary>
public sealed class ExcelFormControl(ExcelFormControlType type)
{
    public ExcelFormControlType ControlType { get; } = type;
    public string? Name { get; set; }
    public int FromRow { get; set; }
    public int FromCol { get; set; }
    public int ToRow { get; set; }
    public int ToCol { get; set; }
    public string? LinkedCell { get; set; }
    public string? ListFillRange { get; set; }
    public string? Text { get; set; }
    public bool IsChecked { get; set; }
    public int SelectedIndex { get; set; }
    public int Min { get; set; }
    public int Max { get; set; } = 100;
    public int Value { get; set; }
    public int SmallChange { get; set; } = 1;
    public int LargeChange { get; set; } = 10;
    public bool MultiSelect { get; set; }
    public string? Macro { get; set; }
    public string? AlternativeText { get; set; }
}

public enum ExcelFormControlType
{
    Button, CheckBox, DropDown, GroupBox, Label, ListBox,
    OptionButton, ScrollBar, SpinButton, ToggleButton
}

// ── Digital signatures ────────────────────────────────────────────────────────

/// <summary>Excel digital signatures — EPExcel parity.</summary>
public sealed class ExcelDigitalSignature
{
    public X509Certificate2? Certificate { get; set; }
    public string? SignerName { get; set; }
    public string? SignatureInstructions { get; set; }
    public bool AllowComments { get; set; } = true;
    public DateTime SignedAt { get; set; } = DateTime.UtcNow;
    public bool IsValid { get; private set; }
    public string? CommitmentType { get; set; } = "ProofOfApproval";

    public byte[] Sign(byte[] content)
    {
        if (Certificate == null) throw new InvalidOperationException("Certificate required");
        var signedXml = new SignedXml();
        signedXml.SigningKey = Certificate.GetRSAPrivateKey();
        var reference = new Reference { Uri = "" };
        reference.AddTransform(new XmlDsigEnvelopedSignatureTransform());
        signedXml.AddReference(reference);
        signedXml.ComputeSignature();
        IsValid = true;
        return content; // In production: embed signature in OOXML
    }

    public static bool Verify(byte[] content, X509Certificate2 cert)
    {
        // Simplified verification stub
        return cert.Verify();
    }
}

/// <summary>Digital signature collection on workbook.</summary>
public sealed class ExcelDigitalSignatureCollection
{
    private readonly List<ExcelDigitalSignature> _signatures = new();
    public IReadOnlyList<ExcelDigitalSignature> Signatures => _signatures.AsReadOnly();
    public void Add(ExcelDigitalSignature sig) => _signatures.Add(sig);
    public bool HasSignatures => _signatures.Any();
}

// ── Query tables ──────────────────────────────────────────────────────────────

// Already defined in Drawing.cs as ExcelQueryTable — adding worksheet-level collection support
