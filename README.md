# API JSON para Excel: qual é o caminho mais rápido

Benchmark de seis formas de levar uma tabela JSON de uma API (com parâmetros na query string) para uma
**tabela nomeada** do Excel. A tabela cresce e encolhe com o número de linhas da API, e as colunas
recebem nomes amigáveis.

**Resposta curta:** Power Query com `Table.FromRows`, acionado pelo VBA. Use o exemplo pronto em
[`exemplo/PowerQueryApi.xlsm`](exemplo/PowerQueryApi.xlsm) (módulo [`vba/ApiTabelas.bas`](vba/ApiTabelas.bas)).

## Resultados

Mediana de 3 execuções, 10 colunas, API local (Node), Excel 365 64 bits no Windows 11.

| Método | 1 mil | 10 mil | 100 mil |
|---|---|---|---|
| **Power Query otimizado (`ApiTabelas.bas`)** | | | **2,4–2,6 s** |
| Power Query `Table.FromRows` | 0,10 s | 0,37 s | 3,5–4,0 s |
| Power Query `Table.FromRecords` | 0,10 s | 0,36 s | 4,4 s |
| Add-in C# (Excel-DNA), gravando pela C API | 0,08 s | 0,71 s | 6,8 s |
| Add-in C# (Excel-DNA), gravando por COM | 0,08 s | 0,80 s | 7,8–9,5 s |
| Python + COM (pywin32) | 0,13 s | 1,0 s | 7,9–10 s |
| VBA com parser próprio | 0,64 s | 2,1 s | 11 s |
| VBA + [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) | 0,29 s | 2,6 s | 28 s |

- O gargalo dos métodos via COM e C API é o Excel gravando as células, não o download nem o parse:
  gravar 1 milhão de células num intervalo comum levou 14 s pelo COM. O Power Query grava por um caminho interno mais rápido.
- A primeira atualização do Power Query na sessão custa cerca de 2,5 s a mais (motor frio).
- Várias tabelas em paralelo (`CarregarVarias`): com 1,5 s de latência por chamada, 3 tabelas caíram
  de 11,7 s para 7,8 s. Com API instantânea não há ganho.

## Usar o exemplo

```vba
' Uma tabela; parâmetros em pares nome, valor da query string.
CarregarTabela "tblVendas", "rows", "100000"

' Várias tabelas em paralelo.
CarregarVarias Array( _
    Array("tblVendas", "rows", "100000"), _
    Array("tblVendasMes", "rows", "30000"))

' Criar uma tabela nova ligada à API (uma vez).
CriarTabelaApi "tblClientes", Worksheets("Clientes").Range("A1")
```

Ajuste `API_BASE`, `API_PATH` e `COLUNAS` no topo de `ApiTabelas.bas`. `COLUNAS` lista os nomes
amigáveis na ordem em que as chaves vêm no JSON. As otimizações estão comentadas no início do módulo.

## Rodar o benchmark

Requer Excel para Windows, Node, [uv](https://docs.astral.sh/uv/) e o SDK do .NET 10.

```sh
node server.mjs                                    # API de teste: /table?rows=N[&delay=ms]
dotnet build csaddin -c Release                    # add-in C#
uv run python bench.py                             # todos os métodos -> results.json
uv run python build_example.py                     # gera e mede exemplo/PowerQueryApi.xlsm
```

Variáveis: `BENCH_SIZES` (padrão `1000,10000,100000`), `BENCH_RUNS` (`3`),
`BENCH_METHODS` (`py,vbaflat,vbajson,cs,cscapi,pqrecords,pqrows`), `API_DELAY_MS` (`0`).

Os scripts ligam temporariamente "Confiar no acesso ao modelo de objeto do VBA" (`AccessVBOM`) para
injetar os módulos e desligam ao terminar.
