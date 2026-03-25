namespace EPExcel.ML.IO;

/// <summary>External connection collection — EPExcel 8.3 parity.</summary>
public sealed class ExcelConnectionCollection
{
    private readonly List<ExcelConnection> _connections = new();
    private int _nextId = 1;

    public IReadOnlyList<ExcelConnection> Connections => _connections.AsReadOnly();

    public ExcelPowerQueryConnection AddPowerQuery(string name, string mQuery)
    {
        var c = new ExcelPowerQueryConnection(_nextId++) { Name = name, MQuery = mQuery };
        _connections.Add(c); return c;
    }

    public ExcelDatabaseConnection AddDatabase(string name, string connectionString,
        ExcelConnectionType type = ExcelConnectionType.Odbc)
    {
        var c = new ExcelDatabaseConnection(_nextId++) { Name = name, ConnectionString = connectionString, Type = type };
        _connections.Add(c); return c;
    }

    public ExcelOlapConnection AddOlap(string name, string connectionString, string commandText)
    {
        var c = new ExcelOlapConnection(_nextId++) { Name = name, ConnectionString = connectionString, CommandText = commandText };
        _connections.Add(c); return c;
    }

    public ExcelWebConnection AddWeb(string name, string url)
    {
        var c = new ExcelWebConnection(_nextId++) { Name = name, Url = url };
        _connections.Add(c); return c;
    }

    public ExcelTextConnection AddText(string name, string filePath)
    {
        var c = new ExcelTextConnection(_nextId++) { Name = name, FilePath = filePath };
        _connections.Add(c); return c;
    }

    public bool Remove(string name)
    {
        var c = _connections.FirstOrDefault(x => x.Name == name);
        return c != null && _connections.Remove(c);
    }

    public ExcelConnection? this[string name] =>
        _connections.FirstOrDefault(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

    internal string ToXml()
    {
        if (!_connections.Any()) return "";
        var sb = new StringBuilder();
        sb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
        sb.AppendLine("""<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">""");
        foreach (var c in _connections) sb.AppendLine(c.ToXml());
        sb.AppendLine("</connections>");
        return sb.ToString();
    }
}

public abstract class ExcelConnection(int id)
{
    public int Id { get; } = id;
    public string Name { get; set; } = "";
    public string? Description { get; set; }
    public bool RefreshOnLoad { get; set; } = true;
    public bool Background { get; set; }
    internal abstract string ToXml();
    protected static string X(string? s) =>
        (s ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
}

public sealed class ExcelPowerQueryConnection(int id) : ExcelConnection(id)
{
    public string MQuery { get; set; } = "";
    public string? TableName { get; set; }
    internal override string ToXml() =>
        $"""  <connection id="{Id}" name="{X(Name)}" type="101" refreshedVersion="8" refreshOnLoad="{(RefreshOnLoad ? 1 : 0)}" background="{(Background ? 1 : 0)}" saveData="1"><dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={X(Name)}" command="SELECT * FROM [{X(Name)}]"/></connection>""";
}

public sealed class ExcelDatabaseConnection(int id) : ExcelConnection(id)
{
    public string ConnectionString { get; set; } = "";
    public string? CommandText { get; set; }
    public ExcelConnectionType Type { get; set; } = ExcelConnectionType.Odbc;
    internal override string ToXml() =>
        $"""  <connection id="{Id}" name="{X(Name)}" type="{(int)Type}" refreshedVersion="8" refreshOnLoad="{(RefreshOnLoad ? 1 : 0)}"><dbPr connection="{X(ConnectionString)}" command="{X(CommandText ?? "")}"/></connection>""";
}

public sealed class ExcelOlapConnection(int id) : ExcelConnection(id)
{
    public string ConnectionString { get; set; } = "";
    public string CommandText { get; set; } = "";
    internal override string ToXml() =>
        $"""  <connection id="{Id}" name="{X(Name)}" type="5" refreshedVersion="8" refreshOnLoad="{(RefreshOnLoad ? 1 : 0)}"><dbPr connection="{X(ConnectionString)}" command="{X(CommandText)}" commandType="2"/><olapPr sendLocale="1" rowDrillCount="1000"/></connection>""";
}

public sealed class ExcelWebConnection(int id) : ExcelConnection(id)
{
    public string Url { get; set; } = "";
    public bool HtmlTables { get; set; } = true;
    internal override string ToXml() =>
        $"""  <connection id="{Id}" name="{X(Name)}" type="7" refreshedVersion="8" refreshOnLoad="{(RefreshOnLoad ? 1 : 0)}"><webPr url="{X(Url)}" htmlTables="{(HtmlTables ? 1 : 0)}" xml="0" sourceData="0"/></connection>""";
}

public sealed class ExcelTextConnection(int id) : ExcelConnection(id)
{
    public string FilePath { get; set; } = "";
    public char Delimiter { get; set; } = ',';
    public bool HasHeaders { get; set; } = true;
    internal override string ToXml() =>
        $"""  <connection id="{Id}" name="{X(Name)}" type="8" refreshedVersion="8" refreshOnLoad="{(RefreshOnLoad ? 1 : 0)}"><textPr sourceFile="{X(FilePath)}" firstRow="{(HasHeaders ? 2 : 1)}"><textFields count="1"><textField type="general"/></textFields></textPr></connection>""";
}

public enum ExcelConnectionType { Odbc = 1, OleDb = 2, WebQuery = 7, TextFile = 8, Olap = 5, PowerQuery = 101 }
