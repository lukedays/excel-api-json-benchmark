Attribute VB_Name = "ApiTabelas"
Option Explicit

' Carrega tabelas JSON de uma API em tabelas nomeadas via Power Query, chamado pelo VBA.
'
' Por que é rápido:
'   1. Power Query grava as células por um caminho interno; COM/VBA/C# gravando Range.Value
'      levam ~2x mais em 100 mil linhas (ver results.json).
'   2. Table.FromRows com a lista de colunas fixa: mais rápido que Table.FromRecords e já
'      nasce com os nomes amigáveis, sem passo de RenameColumns. Exige que a API devolva as
'      chaves sempre na mesma ordem.
'   3. Sem conversão de tipos no M: números do JSON já chegam como número.
'   4. AdjustColumnWidth = False: o autoajuste de colunas varre todas as linhas a cada carga.
'   5. Parâmetros entram trocando a fórmula da consulta, não lendo célula de dentro do M
'      (evita exigir configuração de níveis de privacidade). A fórmula só é regravada se mudou.
'   6. Várias tabelas: CarregarVarias atualiza em segundo plano e espera todas juntas, então as
'      esperas pela API se sobrepõem. Com 1,5 s de latência por chamada, 3 tabelas caíram de
'      11,7 s para 7,8 s. Com API instantânea não há ganho: a gravação no Excel é sequencial.
'   7. ScreenUpdating, EnableEvents e cálculo automático desligados durante a carga.
'   8. A tabela cresce e encolhe sozinha com o número de linhas devolvido pela API.

Private Const API_BASE As String = "http://localhost:8765"
Private Const API_PATH As String = "table"
' Nomes amigáveis, na ordem em que as chaves vêm no JSON.
Private Const COLUNAS As String = "ID|Código|Cliente|Cidade|Data|Quantidade|Preço Unitário|Desconto (%)|Ativo|Categoria"

Private Type EstadoExcel
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    Calculation As XlCalculation
End Type

' Exemplo: uma tabela, depois três tabelas em paralelo com parâmetros diferentes.
Public Sub Exemplo()
    CarregarTabela "tblVendas", "rows", "100000"

    CarregarVarias Array( _
        Array("tblVendas", "rows", "100000"), _
        Array("tblVendasMes", "rows", "30000"), _
        Array("tblVendasDia", "rows", "1000"))
End Sub

' Carrega uma tabela e espera terminar. parametros: pares nome, valor da query string.
Public Sub CarregarTabela(ByVal tabela As String, ParamArray parametros() As Variant)
    Dim lista As Variant
    lista = parametros
    Dim estado As EstadoExcel
    estado = DesligarExcel()
    On Error GoTo Falha

    AtualizarFormula tabela, lista
    With TabelaPorNome(tabela).QueryTable
        .BackgroundQuery = False
        .Refresh BackgroundQuery:=False
    End With
    RestaurarExcel estado
    Exit Sub

Falha:
    Dim numero As Long, origem As String, descricao As String
    numero = Err.Number: origem = Err.Source: descricao = Err.Description
    RestaurarExcel estado
    Err.Raise numero, origem, descricao
End Sub

' Carrega várias tabelas em paralelo. cargas: Array(Array(tabela, nome, valor, ...), ...).
Public Sub CarregarVarias(ByVal cargas As Variant)
    Dim estado As EstadoExcel
    estado = DesligarExcel()
    On Error GoTo Falha

    Dim carga As Variant
    For Each carga In cargas
        AtualizarFormula carga(0), ParametrosDaCarga(carga)
    Next carga
    For Each carga In cargas
        With TabelaPorNome(carga(0)).QueryTable
            .BackgroundQuery = True
            .Refresh BackgroundQuery:=True
        End With
    Next carga
    Application.CalculateUntilAsyncQueriesDone
    RestaurarExcel estado
    Exit Sub

Falha:
    Dim numero As Long, origem As String, descricao As String
    numero = Err.Number: origem = Err.Source: descricao = Err.Description
    RestaurarExcel estado
    Err.Raise numero, origem, descricao
End Sub

' Cria a consulta e a tabela de saída uma única vez. destino: célula do canto superior esquerdo.
Public Sub CriarTabelaApi(ByVal tabela As String, ByVal destino As Range)
    ThisWorkbook.Queries.Add tabela, FormulaApi(Array("rows", "1"))
    Dim saida As ListObject
    Set saida = destino.Worksheet.ListObjects.Add( _
        SourceType:=xlSrcExternal, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & tabela & ";Extended Properties=""""", _
        Destination:=destino)
    saida.Name = tabela
    With saida.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & tabela & "]")
        .AdjustColumnWidth = False
        .PreserveFormatting = True
        .BackgroundQuery = False
        .Refresh BackgroundQuery:=False
    End With
End Sub

' A consulta tem o mesmo nome da tabela.
Private Sub AtualizarFormula(ByVal tabela As String, ByVal parametros As Variant)
    Dim formula As String
    formula = FormulaApi(parametros)
    With ThisWorkbook.Queries(tabela)
        If .Formula <> formula Then .Formula = formula
    End With
End Sub

Private Function FormulaApi(ByVal parametros As Variant) As String
    FormulaApi = _
        "let" & vbLf & _
        "    Origem = Json.Document(Web.Contents(" & TextoM(API_BASE) & ", [RelativePath = " & TextoM(API_PATH) & _
        ", Query = " & RegistroM(parametros) & "]))," & vbLf & _
        "    Tabela = Table.FromRows(List.Transform(Origem, Record.FieldValues), " & ListaM(Split(COLUNAS, "|")) & ")" & vbLf & _
        "in" & vbLf & _
        "    Tabela"
End Function

' Pares nome, valor -> [#"nome" = "valor", ...]
Private Function RegistroM(ByVal parametros As Variant) As String
    Dim i As Long, campos As String
    For i = LBound(parametros) To UBound(parametros) - 1 Step 2
        If Len(campos) > 0 Then campos = campos & ", "
        campos = campos & "#" & TextoM(CStr(parametros(i))) & " = " & TextoM(CStr(parametros(i + 1)))
    Next i
    RegistroM = "[" & campos & "]"
End Function

Private Function ListaM(ByVal itens As Variant) As String
    Dim i As Long, partes() As String
    ReDim partes(LBound(itens) To UBound(itens))
    For i = LBound(itens) To UBound(itens)
        partes(i) = TextoM(CStr(itens(i)))
    Next i
    ListaM = "{" & Join(partes, ", ") & "}"
End Function

Private Function TextoM(ByVal valor As String) As String
    TextoM = """" & Replace(valor, """", """""") & """"
End Function

Private Function DesligarExcel() As EstadoExcel
    With Application
        DesligarExcel.ScreenUpdating = .ScreenUpdating
        DesligarExcel.EnableEvents = .EnableEvents
        DesligarExcel.Calculation = .Calculation
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
End Function

Private Sub RestaurarExcel(ByRef estado As EstadoExcel)
    With Application
        .Calculation = estado.Calculation
        .EnableEvents = estado.EnableEvents
        .ScreenUpdating = estado.ScreenUpdating
    End With
End Sub

' Array(tabela, nome, valor, ...) -> Array(nome, valor, ...)
Private Function ParametrosDaCarga(ByVal carga As Variant) As Variant
    Dim i As Long, resultado() As Variant
    ReDim resultado(0 To UBound(carga) - LBound(carga) - 1)
    For i = LBound(carga) + 1 To UBound(carga)
        resultado(i - LBound(carga) - 1) = carga(i)
    Next i
    ParametrosDaCarga = resultado
End Function

Private Function TabelaPorNome(ByVal tabela As String) As ListObject
    Dim planilha As Worksheet
    For Each planilha In ThisWorkbook.Worksheets
        On Error Resume Next
        Set TabelaPorNome = planilha.ListObjects(tabela)
        On Error GoTo 0
        If Not TabelaPorNome Is Nothing Then Exit Function
    Next planilha
    Err.Raise vbObjectError + 1, , "Tabela não encontrada: " & tabela
End Function
