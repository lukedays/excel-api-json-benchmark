using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace CsAddIn;

/// <summary>Macro exposta ao Excel: baixa a tabela JSON da API e grava numa tabela nomeada.</summary>
public static class TableCommands
{
    private static readonly HttpClient Http = new();

    /// <param name="friendlyNames">"chave=Rótulo|chave=Rótulo"; chaves sem rótulo mantêm o nome original.</param>
    [ExcelCommand(Name = "CsLoadTable")]
    public static void LoadTable(string url, string sheetName, string tableName, string friendlyNames) =>
        Load(url, sheetName, tableName, friendlyNames, BodyWriters.Com);

    /// <summary>Mesma carga, mas o corpo vai pela C API do Excel em vez de COM.</summary>
    [ExcelCommand(Name = "CsLoadTableCApi")]
    public static void LoadTableCApi(string url, string sheetName, string tableName, string friendlyNames) =>
        Load(url, sheetName, tableName, friendlyNames, BodyWriters.CApi);

    private static void Load(string url, string sheetName, string tableName, string friendlyNames, Action<Range, object?[,]> writeBody)
    {
        var json = Http.GetByteArrayAsync(url).GetAwaiter().GetResult();
        var (keys, body) = JsonTableParser.Parse(json);
        var labels = ParseFriendlyNames(friendlyNames);
        var headers = keys.Select(k => (object?)labels.GetValueOrDefault(k, k)).ToArray();

        var app = (Application)ExcelDnaUtil.Application;
        var sheet = (Worksheet)app.ActiveWorkbook.Worksheets[sheetName];
        TableWriter.Write(sheet, tableName, headers, body, writeBody);
    }

    private static Dictionary<string, string> ParseFriendlyNames(string friendlyNames) =>
        friendlyNames
            .Split('|', StringSplitOptions.RemoveEmptyEntries)
            .Select(pair => pair.Split('=', 2))
            .Where(parts => parts.Length == 2)
            .ToDictionary(parts => parts[0], parts => parts[1]);
}
