Attribute VB_Name = "ExcelUtils"
'========================================================================'
'== EXCELUTILS                                                         =='
'==                                                                    =='
'== © Copyright 2018 Rudi Breedenraedt - rudi@breedenraedt.be          =='
'========================================================================'
Option Explicit

'------------------------------------------------------------------------'
'Function FirstColumnOf : returns the first column of the given range.   '
'------------------------------------------------------------------------'
Function FirstColumnOf(ByVal r As Range) As Range  
    Set FirstColumnOf = r.Columns(1)
End Function

'------------------------------------------------------------------------'
'Function FirstRowOf : returns the first row of the given range.         '
'------------------------------------------------------------------------'
Function FirstRowOf(ByVal r As Range) As Range 
    Set FirstRowOf = r.Rows(1)
End Function

'------------------------------------------------------------------------'
'Function LastColumnOf : returns the last column of the given range.     '
'------------------------------------------------------------------------'
Function LastColumnOf(ByVal r As Range) As Range 
    Set LastColumnOf = r.Columns(r.Columns.Count)
End Function

'------------------------------------------------------------------------'
'Function LastRowOf : returns the last row of the given range.           '
'------------------------------------------------------------------------'
Function LastRowOf(ByVal r As Range) As Range  
    Set LastRowOf = r.Rows(r.Rows.Count)
End Function

'------------------------------------------------------------------------'
'Function BSearchRow : binary search for row by value.                   '
'  ws         : worksheet to search                                      '
'  startrow   : rownumber to start searching                             '
'  endrow     : rownumber to end searching                               '
'  forValue   : value to search for                                      '
'  inColumn   : index of the column to search for the value              '
'Returns the row index where the value was found, or 0 if not found.     '
'Note that the values must appear in sorted order in the table for binary'
'search to work.
'------------------------------------------------------------------------'
Function BSearchRow(ByVal ws As Worksheet, ByVal startrow As Integer, ByVal endrow As Integer, ByVal forValue As Variant, ByVal inColumn As Integer) As Integer
    Dim r As Range
    Dim rindex As Integer

    Do
        rindex = (startrow + endrow) / 2
        Set r = ws.Cells(rindex, columnindex)
        If r.value = value Then
            BSearchRow = rindex
            Exit Function
        ElseIf r.value < value Then
            startrow = rindex + 1
        Else
            endrow = rindex - 1
        End If
        If startrow > endrow Then Exit Function
    Loop

End Function
