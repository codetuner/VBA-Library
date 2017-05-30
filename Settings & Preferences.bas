Attribute VB_Name = "Settings & Preferences"
Option Compare Database
Option Explicit

Public Function GetAppSettingStr(ByVal Named As String, Optional ByVal DefaultValue As Variant = Null) As Variant
    Dim rs As Recordset
    
    Set rs = CurrentDb().OpenRecordset("SELECT Value FROM [_AppSettings] WHERE Name=""" & Named & """", dbOpenSnapshot)
    If rs.EOF Then
        GetAppSettingStr = DefaultValue
    Else
        GetAppSettingStr = rs!value.value
    End If
    rs.Close
    
End Function

Public Function GetSettingStr(ByVal Named As String, Optional ByVal DefaultValue As Variant = Null) As Variant
    Dim rs As Recordset
    
    Set rs = CurrentDb().OpenRecordset("SELECT Value FROM [_SystemSettings] WHERE Name=""" & Named & """", dbOpenSnapshot)
    If rs.EOF Then
        GetSettingStr = DefaultValue
    Else
        GetSettingStr = rs!value.value
    End If
    rs.Close
    
End Function

Public Function GetSettingDbl(ByVal Named As String) As Variant
    Dim v As Variant
    v = GetSettingStr(Named)
    If Not IsNull(v) Then
        GetSettingDbl = val(v)
    Else
        GetSettingDbl = Null
    End If
End Function

Public Function GetSettingLng(ByVal Named As String) As Variant
    Dim v As Variant
    v = GetSettingStr(Named)
    If Not IsNull(v) Then
        GetSettingLng = val(v)
    Else
        GetSettingLng = Null
    End If
End Function

Public Function GetAppSettingLng(ByVal Named As String) As Variant
    Dim v As Variant
    v = GetAppSettingStr(Named)
    If Not IsNull(v) Then
        GetAppSettingLng = val(v)
    Else
        GetAppSettingLng = Null
    End If
End Function

Public Function GetAppSettingDbl(ByVal Named As String) As Variant
    Dim v As Variant
    v = GetAppSettingStr(Named)
    If Not IsNull(v) Then
        GetAppSettingDbl = val(v)
    Else
        GetAppSettingDbl = Null
    End If
End Function

Public Sub SetSetting(ByVal Named As String, ByVal value As Variant)
    Dim rs As Recordset
    
    Set rs = CurrentDb().OpenRecordset("SELECT Value FROM [_SystemSettings] WHERE Name=""" & Named & """", dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!Name = Named
        If IsNull(value) Then
            rs!value = Null
        ElseIf IsNumeric(value) Then
            rs!value = str(value)
        Else
            rs!value = CStr(value)
        End If
        rs.Update
    Else
        rs.Edit
        If IsNull(value) Then
            rs!value = Null
        ElseIf IsNumeric(value) Then
            rs!value = str(value)
        Else
            rs!value = CStr(value)
        End If
        rs.Update
    End If
    rs.Close

End Sub

Public Sub SetAppSetting(ByVal Named As String, ByVal value As Variant)
    Dim rs As Recordset
    
    Set rs = CurrentDb().OpenRecordset("SELECT Value FROM [_AppSettings] WHERE Name=""" & Named & """", dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!Name = Named
        If IsNull(value) Then
            rs!value = Null
        ElseIf IsNumeric(value) Then
            rs!value = str(value)
        Else
            rs!value = CStr(value)
        End If
        rs.Update
    Else
        rs.Edit
        If IsNull(value) Then
            rs!value = Null
        ElseIf IsNumeric(value) Then
            rs!value = str(value)
        Else
            rs!value = CStr(value)
        End If
        rs.Update
    End If
    rs.Close

End Sub
