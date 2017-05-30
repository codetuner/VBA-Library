Attribute VB_Name = "DataFunctions"
Option Compare Database
Option Explicit

Public Function QueryValue(ByVal field As String, ByVal QueryString As String, ParamArray Args() As Variant) As Variant
    Dim rs As Recordset
    
    Set rs = CurrentDb.OpenRecordset(QStrInternal(QueryString, Args))
    If rs.EOF Then
        QueryValue = Null
    Else
        QueryValue = rs(field).Value
    End If
    rs.Close

End Function

Public Function QueryValues(ByVal field As String, ByVal QueryString As String, ParamArray Args() As Variant) As Collection
    Dim rs As Recordset
    Dim result As New Collection
    
    Set rs = CurrentDb.OpenRecordset(QStrInternal(QueryString, Args))
    While Not rs.EOF
        result.Add rs(field).Value
        rs.MoveNext
    Wend
    rs.Close
    Set QueryValues = result

End Function

Public Function QueryRow(ByVal QueryString As String, ParamArray Args() As Variant) As Collection
    Dim rs As Recordset
    Dim result As Collection
    Dim fld As field
    
    Set rs = CurrentDb.OpenRecordset(QStrInternal(QueryString, Args))
    If rs.EOF Then
        Set QueryRow = Nothing
    Else
        Set result = New Collection
        For Each fld In rs.Fields
            result.Add fld.Value, fld.Name
        Next
        Set QueryRow = result
    End If
    rs.Close

End Function

Public Function QueryRows(ByVal keyFieldName As String, ByVal QueryString As String, ParamArray Args() As Variant) As Collection
    Dim rs As Recordset
    Dim result As New Collection
    Dim row As Collection
    Dim fld As field
    
    Set rs = CurrentDb.OpenRecordset(QStrInternal(QueryString, Args))
    While Not rs.EOF
        Set row = New Collection
        For Each fld In rs.Fields
            row.Add fld.Value, fld.Name
        Next
        result.Add row, rs(keyFieldName).Value
        rs.MoveNext
    Wend
    rs.Close

    Set QueryRows = result

End Function

Public Function QStr(ByVal str As String, ParamArray Args() As Variant) As String
    QStr = QStrInternal(str, Args)
End Function

Function QStrInternal(ByVal str As String, ParamArray Args() As Variant) As String
    Dim newStr As String
    Dim n As Long
    
    newStr = str
    For n = LBound(Args(0)) To UBound(Args(0))
        newStr = Replace(newStr, "?", Args(0)(n), 1, 1)
    Next
    QStrInternal = newStr

End Function
