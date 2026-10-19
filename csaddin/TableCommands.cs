using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace CsAddIn;

/// <summary>Macro exposta ao Excel: baixa a tabela JSON da API e grava numa tabela nomeada.</summary>
public static class TableCommands
{
    private static readonly HttpClient Http = new(new SocketsHttpHandler
    {
        AutomaticDecompression = System.Net.DecompressionMethods.All,
    });

    /// <param name="friendlyNames">"chave=Rótulo|chave=Rótulo"; chaves sem rótulo mantêm o nome original.</param>
    [ExcelCommand(Name = "CsLoadTable")]
    public static void LoadTable(string url, string sheetName, string tableName, string friendlyNames)
    {
        PhaseTimer.Start();
        var json = Http.GetByteArrayAsync(url).GetAwaiter().GetResult();
        PhaseTimer.Mark("http");
        var parsed = JsonTableParser.Parse(json);
        PhaseTimer.Mark("split");
        var labels = ParseFriendlyNames(friendlyNames);
        var headers = parsed.Keys.Select(k => (object)labels.GetValueOrDefault(k, k)).ToArray();

        var app = (Application)ExcelDnaUtil.Application;
        var sheet = (Worksheet)app.ActiveWorkbook.Worksheets[sheetName];
        TableWriter.Write(sheet, tableName, headers, parsed.RowCount, body => BodyWriter.Write(body, parsed));
        PhaseTimer.Stop();
    }

    /// <summary>Tempo de cada fase da última carga, para diagnóstico.</summary>
    [ExcelFunction(Name = "CsLastTimings")]
    public static string LastTimings() => PhaseTimer.LastReport;

    private static Dictionary<string, string> ParseFriendlyNames(string friendlyNames) =>
        friendlyNames
            .Split('|', StringSplitOptions.RemoveEmptyEntries)
            .Select(pair => pair.Split('=', 2))
            .Where(parts => parts.Length == 2)
            .ToDictionary(parts => parts[0], parts => parts[1]);
}
