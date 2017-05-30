Attribute VB_Name = "DictionaryTools"
Option Explicit

'------------------------------------------------------------------------'
'Function DictParse : builds a collection (with keys) based on a string. '
'See also: DictFastParse.                                                '
'------------------------------------------------------------------------'
Public Function DictParse(ByVal str As String, Optional ByVal ElementSeparator As String = ";", Optional ByVal KeyValueSeparator As String = "=") As Dictionary
    Set DictParse = DictFastParse(str, ElementSeparator, KeyValueSeparator)
End Function

'------------------------------------------------------------------------'
'Function DictFastParse : builds a collection (with keys) based on a     '
'    string.                                                             '
'  The FastParse~ works faster then the Parse~ function, but ignores     '
'  nested collections.                                                   '
'See also: DictParse.                                                    '
'------------------------------------------------------------------------'
Public Function DictFastParse(ByVal str As String, Optional ByVal ElementSeparator As String = ";", Optional ByVal KeyValueSeparator As String = "=") As Dictionary
    Dim Dict As New Dictionary
    Dim element As Variant
    Dim spos As Long
    
    For Each element In PDictParseKeyValuePairs(str, ElementSeparator)
        spos = InStr(element, KeyValueSeparator)
        If spos > 0 Then
            Dict.Add left$(element, spos - 1), Mid$(element, spos + Len(KeyValueSeparator))
        Else
            If Len(element) > 0 Then
                Dict.Add element, True
            End If
        End If
    Next
    Set DictFastParse = Dict
    
End Function

'------------------------------------------------------------------------'
'------------------------------------------------------------------------'
Public Function FormatDictionary(ByVal Dict As Dictionary, Optional ByVal ElementSeparator As String = ";", Optional ByVal KeyValueSeparator As String = "=", Optional ByVal NullValue As String = "#NULL#") As String
    Dim key As Variant
    Dim result As String
    Dim element As Variant
    Dim first As Boolean
    
    first = True
    For Each key In Dict.Keys
        If Not first Then result = result & ElementSeparator Else first = False
        Select Case VarType(key)
        Case vbString
            result = result & key
        Case vbNull
            result = result & NullValue
        Case vbEmpty
            result = result
        Case vbObject
            If key Is Nothing Then
                result = result & NullValue
            ElseIf TypeOf key Is Dictionary Then
                result = result & "(" & FormatDictionary(key, ElementSeparator, KeyValueSeparator, NullValue) & ")"
            ElseIf TypeOf key Is Collection Then
                result = result & "(" & FormatCollection(key, ElementSeparator, NullValue) & ")"
            Else
                result = result & "#" & TypeName(key) & "#"
            End If
        Case Else
            result = result & CStr(key)
        End Select
        If IsObject(Dict.Item(key)) Then
            Set element = Dict.Item(key)
        Else
            element = Dict.Item(key)
        End If
        result = result & KeyValueSeparator
        Select Case VarType(element)
        Case vbString
            result = result & element
        Case vbNull
            result = result & NullValue
        Case vbEmpty
            result = result
        Case vbObject
            If element Is Nothing Then
                result = result & NullValue
            ElseIf TypeOf element Is Dictionary Then
                result = result & "(" & FormatDictionary(element, ElementSeparator, KeyValueSeparator, NullValue) & ")"
            ElseIf TypeOf element Is Collection Then
                result = result & "(" & FormatCollection(element, ElementSeparator, NullValue) & ")"
            Else
                result = result & "#" & TypeName(element) & "#"
            End If
        Case Else
            result = result & CStr(element)
        End Select
    Next
    FormatDictionary = result

End Function

Private Function PDictParseKeyValuePairs(ByVal str As String, Optional ByVal ElementSeparator As String = ";") As Collection
    Dim coll As New Collection
    Dim epos As Long
    
    If Len(str) = 0 Then
        Set PDictParseKeyValuePairs = coll
        Exit Function
    End If
    epos = InStr(str, ElementSeparator)
    Do While epos <> 0
        coll.Add left$(str, epos - 1)
        str = Mid$(str, epos + Len(ElementSeparator))
        epos = InStr(str, ElementSeparator)
    Loop
    coll.Add str
    
    Set PDictParseKeyValuePairs = coll

End Function


