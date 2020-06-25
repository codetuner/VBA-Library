Attribute VB_Name = "CsvModule"
Option Explicit
Option Compare Database

Public Sub ImportCsvAskFile(ByVal tableName As String)
    Dim v As Variant
    With Application.FileDialog(1)
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "CSV files", "*.csv"
        .Filters.Add "All files", "*.*"
        .FilterIndex = 0
        .Title = "Import CSV file into [" & tableName & "]"
        If .Show() <> 0 Then
            For Each v In .SelectedItems
                ImportCsv tableName, v
            Next
        End If
    End With

End Sub

Public Sub ImportCsv(ByVal tableName As String, ByVal fileName As String)

    Dim t As DAO.Recordset
    Dim separator As String
    Dim mappedColumns As New Collection
    Dim mappedColumnIndexes As New Collection
    Dim data As Collection
    Dim v As Variant
    Dim n As Integer
    
    Set t = CurrentDb.TableDefs(tableName).OpenRecordset(dbOpenTable)

    '' Detect separator:
    Open fileName For Input As #1
    separator = AutoDetectSeparator(1)
    Close #1
    
    '' Read header line and map columns:
    Open fileName For Input As #1
    Set data = ReadCsvLineFromFile(1, separator)
    On Error Resume Next
    n = 0
    For Each v In data
        n = n + 1
        mappedColumns.Add t.Fields(v)
        If Err.Number = 3265 Then
            mappedColumns.Add Nothing
            Err.Clear
        Else
            mappedColumnIndexes.Add n
        End If
    Next
    On Error GoTo 0
    
    '' Proceed with other lines:
    Do While Not EOF(1)
        Set data = ReadCsvLineFromFile(1, separator)
        t.AddNew
        For Each v In mappedColumnIndexes
            If data(v) <> "" Then mappedColumns(v).Value = data(v)
        Next
        t.Update
    Loop

    Close #1
    
    t.Close
    Set t = Nothing

End Sub

Private Function AutoDetectSeparator(ByVal fileNumber As Integer) As String
    Dim linestr As String
    Dim septab As Integer, sepcom As Integer, sepsem As Integer
    Line Input #fileNumber, linestr
    septab = InStr(linestr, vbTab)
    If septab = 0 Then septab = Len(linestr)
    sepcom = InStr(linestr, ",")
    If sepcom = 0 Then sepcom = Len(linestr)
    sepsem = InStr(linestr, ";")
    If sepsem = 0 Then sepsem = Len(linestr)
    If septab < sepcom Then
        AutoDetectSeparator = vbTab
    ElseIf sepcom < sepsem Then
        AutoDetectSeparator = ","
    Else
        AutoDetectSeparator = ";"
    End If
End Function

Private Function ReadCsvLineFromFile(ByVal fileNumber As Integer, ByVal separator As String) As Collection
    Dim result As New Collection
    Dim str As String
    
    Line Input #fileNumber, str
    
    Set ReadCsvLineFromFile = ReadCsvLine(str, separator)
    
End Function

Private Function ReadCsvLine(ByVal str As String, ByVal separator As String) As Collection
    Dim result As New Collection
    Dim startpos As Integer, seppos As Integer, endpos As Integer
    
    '' Skip UTF BOM:
    If Left(str, 3) = "﻿" Then str = Mid(str, 4)
    
    startpos = 1
    Do
        If Mid(str, startpos, 1) = """" Then
            endpos = startpos
            Do
                endpos = InStr(endpos + 1, str, """")
                If endpos = Len(str) Then Exit Do
                If Mid(str, endpos + 1, 1) = separator Then Exit Do
            Loop
            result.Add Replace(Mid(str, startpos + 1, endpos - startpos - 1), """""", """")
            startpos = endpos + 1 + Len(separator)
        Else
            endpos = InStr(startpos, str, separator)
            If endpos = 0 Then
                result.Add Mid(str, startpos)
                startpos = Len(str)
            Else
                result.Add Mid(str, startpos, endpos - startpos)
                startpos = endpos + 1
            End If
        End If
    Loop While startpos < Len(str)
    
    Set ReadCsvLine = result

End Function







