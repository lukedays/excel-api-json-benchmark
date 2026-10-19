using System.Text.Json;

namespace CsAddIn;

/// <summary>Converte um array JSON de objetos planos em cabeçalhos + matriz de valores para o Excel.</summary>
public static class JsonTableParser
{
    public static (string[] Keys, object?[,] Body) Parse(byte[] utf8Json)
    {
        using var document = JsonDocument.Parse(utf8Json);
        var rows = document.RootElement;
        var rowCount = rows.GetArrayLength();
        if (rowCount == 0)
        {
            return ([], new object?[0, 0]);
        }

        var keys = rows[0].EnumerateObject().Select(p => p.Name).ToArray();
        var columnOf = keys.Select((key, index) => (key, index)).ToDictionary(p => p.key, p => p.index);
        var body = new object?[rowCount, keys.Length];

        var r = 0;
        foreach (var row in rows.EnumerateArray())
        {
            foreach (var property in row.EnumerateObject())
            {
                if (columnOf.TryGetValue(property.Name, out var c))
                {
                    body[r, c] = ToCellValue(property.Value);
                }
            }
            r++;
        }
        return (keys, body);
    }

    private static object? ToCellValue(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.String => value.GetString(),
        JsonValueKind.Number => value.GetDouble(),
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Null or JsonValueKind.Undefined => null,
        _ => value.GetRawText(),
    };
}
