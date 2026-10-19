# API JSON para Excel: qual é o caminho mais rápido

Benchmark de formas de levar uma tabela JSON de uma API (com parâmetros na query string) para uma
**tabela nomeada** do Excel. A tabela cresce e encolhe com o número de linhas da API, e as colunas
recebem nomes amigáveis.

**Resposta curta:** o add-in C# otimizado ([`csaddin/`](csaddin), Excel-DNA). Ele carrega 100 mil linhas
em cerca de 1 s e 1 mil linhas em 0,15 s, sem o custo fixo de 0,1 a 1,8 s do Power Query.

## Resultados

10 colunas, API local (Node), Excel 365 64 bits, Windows 11. Valores em segundos, do pedido à tabela pronta.

| Método | 1 mil | 10 mil | 30 mil | 100 mil |
|---|---|---|---|---|
| **Add-in C# otimizado (`CsLoadTable`)** | **0,11–0,18** | **0,12** | **0,31–0,45** | **0,94–1,01** |
| Power Query otimizado ([`vba/ApiTabelas.bas`](vba/ApiTabelas.bas)) | 0,17–1,9 | 0,36–0,39 | | 2,4–2,8 |
| Power Query `Table.FromRecords` | 0,10 | 0,36 | | 4,4 |
| Python + COM (pywin32) | 0,13 | 1,0 | | 7,9–10 |
| VBA com parser próprio | 0,64 | 2,1 | | 11 |
| VBA + [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) | 0,29 | 2,6 | | 28 |

- O add-in C# foi medido num Excel limpo (`bench_csharp.py`). As demais linhas vêm de `bench.py`,
  numa pasta de trabalho que acumula várias tabelas de 100 mil linhas e deixa o Excel mais lento.
  Nesse cenário, o add-in C# na versão anterior (sem os blocos paralelos) levava 6,8 s.
- O Power Query tem custo fixo: depois de alguns segundos ocioso, o Excel encerra os processos que
  executam as consultas, e a próxima carga gasta cerca de 1,5 s em "Connecting to datasource".

## O que deixa o add-in C# rápido

Fases medidas em 100 mil linhas: HTTP 0,03 s, divisão do JSON 0,02 s, gravação 0,87 s. A gravação das
células é o piso do próprio Excel; o resto foi otimizado até ficar abaixo do ruído.

1. **Gravação pela C API (`xlSet`)**, não pelo COM `Range.Value2`: 0,9 s contra 1,35 s em 1 milhão de células.
2. **Parse paralelo sobreposto à gravação.** Uma varredura linear acha o início de cada linha (0,02 s).
   Os blocos são interpretados em paralelo com `Utf8JsonReader`, sem montar DOM, cerca de dois por núcleo.
   O Excel grava o bloco 1 enquanto os seguintes ainda estão sendo interpretados.
3. **Matriz pronta para a C API.** Vazios já saem como `ExcelEmpty` e booleanos usam instâncias
   compartilhadas, então não há cópia nem conversão antes de gravar.
4. **A tabela é dimensionada, não recriada.** `ListObject.Resize` custa milissegundos. Desfazer e recriar
   a tabela custou 3,4 s.
5. **Cabeçalho e corpo em gravações separadas.** Um bloco único gravado sobre a tabela inteira faz o
   Excel desfazê-la.
6. **Sobras apagadas pela C API** quando a tabela encolhe: `xlSet` com um único valor vazio vale para o
   intervalo todo (0,08 s contra 0,13 s do `Clear`).
7. **HTTP com descompressão automática** (gzip, deflate, brotli), que importa em APIs remotas.

O que não ajudou, medido: gravar em blocos menores pelo COM, desligar tela/eventos/cálculo durante a
gravação via C API, tamanhos de bloco de 4 mil a 65 mil linhas (diferenças dentro do ruído).

## Usar o add-in

```sh
dotnet build csaddin -c Release
```

Carregue `csaddin\bin\Release\net10.0-windows\CsAddIn-AddIn64.xll` no Excel (Arquivo > Opções >
Suplementos > Suplementos do Excel) e chame pelo VBA:

```vba
Application.Run "CsLoadTable", _
    "https://sua.api/table?rows=100000", _   ' URL com os parâmetros
    "Vendas", _                              ' planilha
    "tblVendas", _                           ' tabela nomeada (criada se não existir)
    "id=ID|preco=Preço Unitário"             ' nomes amigáveis: chave=Rótulo
```

`=CsLastTimings()` mostra o tempo de cada fase da última carga.

## Rodar os benchmarks

Requer Excel para Windows, Node, [uv](https://docs.astral.sh/uv/) e o SDK do .NET 10.

```sh
node server.mjs                       # API de teste: /table?rows=N[&delay=ms]
dotnet build csaddin -c Release       # add-in C#
uv run python bench_csharp.py         # add-in C# num Excel limpo, com validação dos dados
uv run python bench.py                # todos os métodos -> results.json
uv run python build_example.py        # gera e mede exemplo/PowerQueryApi.xlsm
```

Variáveis: `BENCH_SEQUENCE` (`bench_csharp.py`, padrão `1000,10000,100000,100000,1000`),
`BENCH_SIZES` (`1000,10000,100000`), `BENCH_RUNS` (`3`), `BENCH_METHODS`
(`py,vbaflat,vbajson,cs,pqrecords,pqrows`), `API_DELAY_MS` (`0`).

Os scripts ligam temporariamente "Confiar no acesso ao modelo de objeto do VBA" (`AccessVBOM`) para
injetar os módulos e desligam ao terminar.

## Power Query via VBA

Se preferir não instalar add-in, [`exemplo/PowerQueryApi.xlsm`](exemplo/PowerQueryApi.xlsm) traz
`CarregarTabela`, `CarregarVarias` (paralelo, útil quando a API demora) e `CriarTabelaApi`, com as
otimizações comentadas no topo de [`vba/ApiTabelas.bas`](vba/ApiTabelas.bas).
