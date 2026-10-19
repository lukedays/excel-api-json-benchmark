Attribute VB_Name = "BenchFetch"
Option Explicit

' Baixa o corpo da API como texto (UTF-8).
Private Function HttpGetText(ByVal url As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.send
    HttpGetText = http.responseText
End Function

' Variante 1: VBA-JSON (biblioteca padrão de mercado) + escrita em bloco via array.
Public Sub LoadWithVbaJson(ByVal url As String, ByVal sheetName As String, ByVal tableName As String, ByVal friendlyNames As String)
    Dim rows As Collection, row As Dictionary, headers As Variant
    Set rows = JsonConverter.ParseJson(HttpGetText(url))
    headers = rows(1).Keys

    Dim colCount As Long, rowCount As Long, r As Long, c As Long
    colCount = UBound(headers) + 1
    rowCount = rows.Count

    Dim cells() As Variant
    ReDim cells(0 To rowCount - 1, 0 To colCount - 1)
    r = 0
    For Each row In rows
        For c = 0 To colCount - 1
            cells(r, c) = row(headers(c))
        Next c
        r = r + 1
    Next row

    WriteToTable ThisWorkbook.Worksheets(sheetName), tableName, headers, cells, rowCount, colCount, friendlyNames
End Sub

' Variante 2: parser próprio para array de objetos planos, varrendo o texto como bytes UTF-16.
Public Sub LoadWithFlatParser(ByVal url As String, ByVal sheetName As String, ByVal tableName As String, ByVal friendlyNames As String)
    Dim cells As Variant, headers As Variant, rowCount As Long, colCount As Long
    cells = ParseFlatJsonTable(HttpGetText(url), headers, rowCount, colCount)
    WriteToTable ThisWorkbook.Worksheets(sheetName), tableName, headers, cells, rowCount, colCount, friendlyNames
End Sub

' Grava cabeçalhos (vetor) e corpo (matriz) numa tabela nomeada, crescendo ou encolhendo para caber exatamente.
' friendlyNames: "chave=Rótulo|chave=Rótulo"; chaves sem rótulo mantêm o nome original.
' Cabeçalho e corpo vão em gravações separadas: um bloco único sobre a tabela inteira a desfaz.
Private Sub WriteToTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal headers As Variant, ByRef cells As Variant, _
                         ByVal rowCount As Long, ByVal colCount As Long, ByVal friendlyNames As String)
    Dim headerRow() As Variant
    headerRow = FriendlyHeaderRow(headers, colCount, friendlyNames)

    Dim table As ListObject, oldRange As Range, newRange As Range
    On Error Resume Next
    Set table = ws.ListObjects(tableName)
    On Error GoTo 0

    Dim anchor As Range
    If table Is Nothing Then Set anchor = ws.Range("A1") Else Set anchor = table.Range.Cells(1, 1)
    ' Tabela precisa de ao menos uma linha de dados.
    Set newRange = anchor.Resize(IIf(rowCount = 0, 2, rowCount + 1), colCount)

    If table Is Nothing Then
        newRange.Rows(1).Value = headerRow
        If rowCount > 0 Then newRange.Offset(1).Resize(rowCount).Value = cells
        Set table = ws.ListObjects.Add(xlSrcRange, newRange, , xlYes)
        table.Name = tableName
        Exit Sub
    End If

    Set oldRange = table.Range
    If table.ShowAutoFilter Then If table.AutoFilter.FilterMode Then table.AutoFilter.ShowAllData
    table.Resize newRange
    table.HeaderRowRange.Value = headerRow
    If rowCount = 0 Then
        table.DataBodyRange.ClearContents
    Else
        table.DataBodyRange.Resize(rowCount).Value = cells
    End If
    ClearLeftovers ws, oldRange, newRange
End Sub

' Limpa as células que a tabela ocupava antes e não ocupa mais (linhas abaixo e colunas à direita).
Private Sub ClearLeftovers(ByVal ws As Worksheet, ByVal oldRange As Range, ByVal newRange As Range)
    If oldRange.Rows.Count > newRange.Rows.Count Then
        oldRange.Offset(newRange.Rows.Count).Resize(oldRange.Rows.Count - newRange.Rows.Count).Clear
    End If
    If oldRange.Columns.Count > newRange.Columns.Count Then
        oldRange.Offset(, newRange.Columns.Count).Resize(, oldRange.Columns.Count - newRange.Columns.Count).Clear
    End If
End Sub

Private Function FriendlyHeaderRow(ByVal headers As Variant, ByVal colCount As Long, ByVal friendlyNames As String) As Variant()
    Dim labels As New Dictionary, pair As Variant, parts() As String, c As Long
    For Each pair In Split(friendlyNames, "|")
        parts = Split(pair, "=", 2)
        If UBound(parts) = 1 Then labels(parts(0)) = parts(1)
    Next pair
    Dim headerRow() As Variant
    ReDim headerRow(1 To 1, 1 To colCount)
    For c = 0 To colCount - 1
        If labels.Exists(headers(c)) Then headerRow(1, c + 1) = labels(headers(c)) Else headerRow(1, c + 1) = headers(c)
    Next c
    FriendlyHeaderRow = headerRow
End Function

' Converte [{"k":v,...},...] numa matriz de valores e devolve as chaves em headers. Valores aninhados não são suportados.
Private Function ParseFlatJsonTable(ByRef json As String, ByRef headersOut As Variant, ByRef rowCount As Long, ByRef colCount As Long) As Variant
    Dim b() As Byte
    b = json
    Dim n As Long: n = UBound(b) - 1

    ' Limite superior de linhas: quantidade de "{" no texto.
    Dim maxRows As Long, p As Long
    p = InStr(1, json, "{")
    Do While p > 0
        maxRows = maxRows + 1
        p = InStr(p + 1, json, "{")
    Loop
    If maxRows = 0 Then Exit Function

    Const MAX_COLS As Long = 256
    Dim headers(0 To MAX_COLS - 1) As String
    Dim cells() As Variant
    ReDim cells(0 To maxRows - 1, 0 To MAX_COLS - 1)

    Dim i As Long, ch As Long, key As String, col As Long, r As Long, k As Long
    Dim inFirstRow As Boolean
    r = 0
    Do While i <= n
        ch = b(i)
        If ch = 123 Then ' {
            r = r + 1
            col = 0
            inFirstRow = (r = 1)
            i = i + 2
        ElseIf ch = 34 Then ' chave
            key = ReadJsonString(json, b, i)
            i = SkipToValue(b, i)
            If inFirstRow Then
                headers(col) = key
                colCount = col + 1
            ElseIf headers(col) <> key Then
                For k = 0 To colCount - 1
                    If headers(k) = key Then col = k: Exit For
                Next k
            End If
            cells(r - 1, col) = ReadJsonValue(json, b, i, n)
            col = col + 1
        Else
            i = i + 2
        End If
    Loop

    rowCount = r
    ReDim Preserve cells(0 To maxRows - 1, 0 To colCount - 1)
    Dim keys() As Variant, c As Long
    ReDim keys(0 To colCount - 1)
    For c = 0 To colCount - 1
        keys(c) = headers(c)
    Next c
    headersOut = keys
    ParseFlatJsonTable = cells
End Function

Private Function SkipToValue(ByRef b() As Byte, ByVal i As Long) As Long
    Do While b(i) <> 58 ' :
        i = i + 2
    Loop
    i = i + 2
    Do While b(i) = 32 Or b(i) = 9 Or b(i) = 10 Or b(i) = 13
        i = i + 2
    Loop
    SkipToValue = i
End Function

' Lê o valor em b(i) e avança i para depois dele. json e b são o mesmo texto: b para varrer, json para recortar.
Private Function ReadJsonValue(ByRef json As String, ByRef b() As Byte, ByRef i As Long, ByVal n As Long) As Variant
    Dim ch As Long: ch = b(i)
    If ch = 34 Then
        ReadJsonValue = ReadJsonString(json, b, i)
    ElseIf ch = 116 Then ' true
        ReadJsonValue = True: i = i + 8
    ElseIf ch = 102 Then ' false
        ReadJsonValue = False: i = i + 10
    ElseIf ch = 110 Then ' null
        ReadJsonValue = Empty: i = i + 8
    Else
        Dim start As Long: start = i
        Do While i <= n
            ch = b(i)
            If ch = 44 Or ch = 125 Or ch = 93 Or ch <= 32 Then Exit Do
            i = i + 2
        Loop
        ReadJsonValue = Val(MidB$(json, start + 1, i - start))
    End If
End Function

' Lê uma string JSON começando na aspa em b(i); i termina depois da aspa de fechamento.
Private Function ReadJsonString(ByRef json As String, ByRef b() As Byte, ByRef i As Long) As String
    Dim start As Long, hasEscape As Boolean, text As String
    i = i + 2
    start = i
    Do While b(i) <> 34 Or b(i + 1) <> 0
        If b(i) = 92 And b(i + 1) = 0 Then hasEscape = True: i = i + 2
        i = i + 2
    Loop
    text = MidB$(json, start + 1, i - start)
    i = i + 2
    If hasEscape Then text = UnescapeJson(text)
    ReadJsonString = text
End Function

Private Function UnescapeJson(ByVal s As String) As String
    Dim out As String, p As Long, ch As String
    p = 1
    Do While p <= Len(s)
        ch = Mid$(s, p, 1)
        If ch = "\" Then
            p = p + 1
            Select Case Mid$(s, p, 1)
                Case "n": out = out & vbLf
                Case "t": out = out & vbTab
                Case "r": out = out & vbCr
                Case "b": out = out & Chr$(8)
                Case "f": out = out & Chr$(12)
                Case "u": out = out & ChrW$(CLng("&H" & Mid$(s, p + 1, 4))): p = p + 4
                Case Else: out = out & Mid$(s, p, 1)
            End Select
        Else
            out = out & ch
        End If
        p = p + 1
    Loop
    UnescapeJson = out
End Function
