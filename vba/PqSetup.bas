Attribute VB_Name = "PqSetup"
Option Explicit

' Cria a tabela nomeada que recebe a saída da consulta do Power Query (carregamento em planilha).
Public Sub CreateQueryTable(ByVal sheetName As String, ByVal queryName As String, ByVal tableName As String)
    Dim table As ListObject
    Set table = ThisWorkbook.Worksheets(sheetName).ListObjects.Add( _
        SourceType:=xlSrcExternal, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & ";Extended Properties=""""", _
        Destination:=ThisWorkbook.Worksheets(sheetName).Range("A1"))
    table.Name = tableName
    With table.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & queryName & "]")
        .BackgroundQuery = False
    End With
End Sub
