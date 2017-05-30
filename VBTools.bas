Attribute VB_Name = "VBTools"
'========================================================================'
'== VBTOOLS                                                            =='
'==                                                                    =='
'== © Copyright 1997-2001 Rudi Breedenraedt - rudi@breedenraedt.be     =='
'========================================================================'
Option Explicit

'========================================================================'
'Note, 2001-09-25, Rudi Breedenraedt                                     '
'  FileExists, LoadFile and WriteFile moved to FileTools.bas             '
'========================================================================'

'------------------------------------------------------------------------'
'Function Construct : creates a new object and initializes it.           '
'------------------------------------------------------------------------'
Public Function Construct(ByVal Obj As Object, ByVal InitMethod As String, ParamArray Arguments() As Variant) As Object
    Select Case UBound(Arguments) - LBound(Arguments) + 1
        Case 0: CallByName Obj, InitMethod, VbMethod
        Case 1: CallByName Obj, InitMethod, VbMethod, Arguments(0)
        Case 2: CallByName Obj, InitMethod, VbMethod, Arguments(0), Arguments(1)
        Case 3: CallByName Obj, InitMethod, VbMethod, Arguments(0), Arguments(1), Arguments(2)
        Case 4: CallByName Obj, InitMethod, VbMethod, Arguments(0), Arguments(1), Arguments(2), Arguments(3)
        Case 5: CallByName Obj, InitMethod, VbMethod, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4)
        Case 6: CallByName Obj, InitMethod, VbMethod, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5)
        Case 7: CallByName Obj, InitMethod, VbMethod, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6)
        Case 8: CallByName Obj, InitMethod, VbMethod, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7)
        Case 9: CallByName Obj, InitMethod, VbMethod, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8)
        Case 10: CallByName Obj, InitMethod, VbMethod, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9)
        Case 11: CallByName Obj, InitMethod, VbMethod, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10)
        Case 12: CallByName Obj, InitMethod, VbMethod, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(10)
        Case Else: Err.Raise 5, , "Too many arguments."
    End Select
    Set Construct = Obj
End Function

'------------------------------------------------------------------------'
'Function AssignVar : assigns a value to a variant, where the value can  '
'  be a value or an object reference. The function returns the value.    '
'  Typically, use this to avoid If IsObject(x) Then Set y = x Else y = x '
'  constructions which - if y is a function - call y twice.              '
'------------------------------------------------------------------------'
Public Function AssignVar(ByRef Var As Variant, Value As Variant) As Variant
    If IsObject(Value) Then
        Set Var = Value
        Set AssignVar = Value
    Else
        Var = Value
        AssignVar = Value
    End If
End Function

'------------------------------------------------------------------------'
'Function VarComp : Compares two variant values. If they are equal,      '
'  returns True, otherwise returns False. If one of the variants is Null,'
'  Null is returned.                                                     '
'------------------------------------------------------------------------'
Public Function VarComp(Var1 As Variant, Var2 As Variant, Optional Compare As VbCompareMethod) As Variant
    
    VarComp = False
    If IsNumeric(Var1) And IsNumeric(Var2) Then
        VarComp = (Var1 = Var2)
    ElseIf VarType(Var1) = VarType(Var2) Then
        Select Case VarType(Var1)
        Case 0, 1
            VarComp = True
        Case 2 To 7, 11, 14, 17
            VarComp = (Var1 = Var2)
        Case 8
            VarComp = (StrComp(Var1, Var2, Compare) = 0)
        Case 9
            VarComp = (Var1 Is Var2)
        Case Else
            Debug.Print "Datatype not yet supported in VBTools.VarComp."
            Debug.Assert False
        End Select
    End If

End Function

'------------------------------------------------------------------------'
'Function CollComp : Compares two collections. Returns True if both      '
'  collections contain the same values.                                  '
'------------------------------------------------------------------------'
Public Function CollComp(Coll1 As Collection, Coll2 As Collection) As Boolean
    Dim i As Integer
    
    CollComp = False
    If Not (Coll1 Is Coll2) Then
        If Coll1.Count = Coll2.Count Then
            For i = 1 To Coll1.Count
                If (TypeName(Coll1(i)) = "Collection") And (TypeName(Coll2(i)) = "Collection") Then
                    If Not CollComp(Coll1(i), Coll2(i)) Then Exit Function
                Else
                    If VarComp(Coll1(i), Coll2(i)) <> True Then Exit Function
                End If
            Next
        Else
            Exit Function
        End If
    End If
    CollComp = True

End Function

'------------------------------------------------------------------------'
'Function NVL : if Value is Null, returns NullValue, else returns Value  '
'------------------------------------------------------------------------'
Public Function NVL(ByVal Value As Variant, ByVal NullValue As Variant) As Variant
    If IsNull(Value) Then
        AssignVar NVL, NullValue
    Else
        AssignVar NVL, Value
    End If
End Function

'------------------------------------------------------------------------'
'Function IIF : Internal-IF, if condition is True, returns IfTrue, else  '
'  returns IfFalse                                                       '
'------------------------------------------------------------------------'
Public Function IIf(ByVal Condition As Boolean, ByVal IfTrue As Variant, ByVal IfFalse As Variant) As Variant
    If Condition Then
        AssignVar IIf, IfTrue
    Else
        AssignVar IIf, IfFalse
    End If
End Function

'------------------------------------------------------------------------'
'Function ICase : Internal-Case, matches a variable to a value and       '
'  returns the corresponding value.                                      '
'E.g: MsgBox ICase(x, 1, "One", 2, "Two", 3, "Three", "More than 3")     '
'------------------------------------------------------------------------'
Public Function ICase(ByVal Value As Variant, ParamArray Cases() As Variant)
    Dim n As Integer
    Dim casecount As Integer
    
    casecount = ((UBound(Cases) - LBound(Cases)) + 1) \ 2
    If (LBound(Cases) + (casecount * 2)) = UBound(Cases) Then
        AssignVar ICase, Cases(UBound(Cases))
    Else
        ICase = Null
    End If
    If IsNull(Value) Then
        For n = 1 To casecount
            If IsNull(Cases(n * 2 - 2 + LBound(Cases))) Then
                AssignVar ICase, Cases(n * 2 - 1 + LBound(Cases))
                Exit Function
            End If
        Next
    Else
        For n = 1 To casecount
            If VarComp(Value, Cases(n * 2 - 2 + LBound(Cases))) = True Then
                AssignVar ICase, Cases(n * 2 - 1 + LBound(Cases))
                Exit Function
            End If
        Next
    End If

End Function

'------------------------------------------------------------------------'
'Function Inside : Checks if a value matches one of the others (like an  '
'  IN operator).                                                         '
'E.g: If Inside(x, 1, 2, 3) Then ...                                     '
'------------------------------------------------------------------------'
Public Function Inside(ByVal arg As Variant, ParamArray values() As Variant) As Boolean
    Dim n As Long
    Inside = False
    For n = LBound(values) To UBound(values)
        If arg = values(n) Then Inside = True
    Next
End Function

'------------------------------------------------------------------------'
'Function Between: Checks if a value lays inside an interval.            '
'E.g: If Between(x, 0, 100) Then MsgBox "'0 <= x <= 100' is True"        '
'------------------------------------------------------------------------'
Public Function Between(ByVal arg As Variant, ByVal Max As Variant, ByVal Min As Variant) As Boolean
    If arg >= Min And arg <= Max Then Between = True Else Between = False
End Function

'------------------------------------------------------------------------'
'------------------------------------------------------------------------'
Function Min(ParamArray values() As Variant) As Variant
    Dim v As Variant
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If i = LBound(values) Then
            v = values(i)
        ElseIf values(i) < v Then
            v = values(i)
        End If
    Next
    Min = v
End Function

'------------------------------------------------------------------------'
'------------------------------------------------------------------------'
Function Max(ParamArray values() As Variant) As Variant
    Dim v As Variant
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If i = LBound(values) Then
            v = values(i)
        ElseIf values(i) > v Then
            v = values(i)
        End If
    Next
    Max = v
End Function

'------------------------------------------------------------------------'
'------------------------------------------------------------------------'
Function MinIndex(ParamArray values() As Variant) As Variant
    Dim v As Long
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If i = LBound(values) Then
            v = i + 1
        ElseIf values(i) < values(v - 1) Then
            v = i + 1
        End If
    Next
    MinIndex = v
End Function

'------------------------------------------------------------------------'
'------------------------------------------------------------------------'
Function MaxIndex(ParamArray values() As Variant) As Variant
    Dim v As Long
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If i = LBound(values) Then
            v = i + 1
        ElseIf values(i) > values(v - 1) Then
            v = i + 1
        End If
    Next
    MaxIndex = v
End Function

'------------------------------------------------------------------------'
'Function FlagIsSet : Checks if a flag is set in a binary coded field    '
'E.g: If FlagIsSet(aFile.Attributes, FLAG_READONLY) Then...              '
'------------------------------------------------------------------------'
Public Function FlagIsSet(ByVal Flags As Long, ByVal Flag As Long) As Boolean
    FlagIsSet = ((Flags And Flag) <> 0)
End Function

'------------------------------------------------------------------------'
'Function ToString : Converts anything into a string. Usefull for        '
'  debugging.                                                            '
'------------------------------------------------------------------------'
Public Function ToString(Anything As Variant) As String
    Dim s As String
    Dim v As Variant
    Dim i As Long
    
    On Error Resume Next
    ToString = "<Unknown>"
    If VarType(Anything) < vbArray Then
        Select Case VarType(Anything)
        Case vbEmpty
            ToString = "{Empty}"
        Case vbNull
            ToString = "{Null}"
        Case 2 To 8, 11, 14, 17
            ToString = CStr(Anything)
        Case vbObject
            If Anything Is Nothing Then
                ToString = "{Nothing}"
            Else
                Select Case TypeName(Anything)
                Case "Collection"
                    s = ""
                    For Each v In Anything
                        s = s + ToString(v) + ", "
                    Next
                    If Len(s) > 0 Then
                        ToString = "{Coll:[" + Left$(s, Len(s) - 2) + "]}"
                    Else
                        ToString = "{Coll:[]}"
                    End If
                Case Else
                    ToString = "{" + TypeName(Anything) + "}"
                    ToString = Anything.ToString()
                End Select
            End If
        Case vbDataObject
            ToString = "{DataObject}"
            ToString = "{DataObject:" + TypeName(Anything) + "}"
        Case vbError
            ToString = "{Error}"
        Case Else
            ToString = "{Unknown}"
        End Select
    Else
        s = "{Arr:["
        For i = LBound(Anything) To UBound(Anything)
            s = s & ToString(Anything(i))
            If i < UBound(Anything) Then s = s & ", "
        Next
        s = s & "]}"
        ToString = s
    End If
    
End Function

'------------------------------------------------------------------------'
'Function StrLen : returns the length of a string; supports Null string  '
'------------------------------------------------------------------------'
Public Function StrLen(ByVal str As String) As Long
    StrLen = NVL(Len(str), 0)
End Function
