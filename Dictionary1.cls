VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Keys As New Collection
Private values As New Collection

Public Function At(ByVal key As String, Optional byDefault As Variant = Null) As Variant
    On Error Resume Next
    At = byDefault
    At = values(UCase$(key))
End Function

Public Function AsString() As String
    Dim result As String
    Dim c As Integer
    For c = 1 To Keys.Count
        result = result & Keys(c) & "=" & CStr(values(c)) & Chr$(9)
    Next
    result = result & Chr$(9)
    AsString = result
End Function

Public Function FromString(ByVal str As Variant) As Dictionary
    Dim nextpos As Integer
    Dim current As String
    
    If IsNull(str) Then Exit Function
    
    nextpos = 1
    Do
        nextpos = InStr(str, Chr$(9))
        If nextpos > 1 Then
            current = Left$(str, nextpos - 1)
            Me.Add Left$(current, InStr(current, "=") - 1), Mid$(current, InStr(current, "=") + 1)
            str = Mid$(str, nextpos + 1)
        End If
    Loop While nextpos > 1

End Function

Public Sub Add(ByVal key As String, ByVal item As Variant)
    Keys.Add key
    values.Add item, UCase(key)
End Sub

Public Sub DebugDict()
    Dim c As Integer
    Debug.Print "Dictionary:"
    For c = 1 To Keys.Count
        Debug.Print " - " & Keys(c) & "=" & CStr(values(c))
    Next
End Sub

