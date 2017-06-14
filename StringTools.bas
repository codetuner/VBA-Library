Attribute VB_Name = "StringTools"
'========================================================================'
'== STRINGTOOLS                                                        =='
'==                                                                    =='
'== © Copyright 1999-2001 Rudi Breedenraedt - rudi@breedenraedt.be     =='
'========================================================================'
Option Explicit

'------------------------------------------------------------------------'
'Sub StrAdd : appends a string to another                                '
'See also: StrAddLn                                                      '
'------------------------------------------------------------------------'
Public Sub StrAdd(ByRef Target As String, ByVal Source As String)
    Target = Target + Source
End Sub

'------------------------------------------------------------------------'
'Sub StrAddLn : appends a string to another and adds a linefeed          '
'See also: StrAdd                                                        '
'------------------------------------------------------------------------'
Public Sub StrAddLn(ByRef Target As String, Optional ByVal Source As String = "")
    Target = Target + Source + vbCrLf
End Sub

'------------------------------------------------------------------------'
'Function CStrNull : converts the value into a string; if the value is   '
'  Null, NullValue is returned                                           '
'------------------------------------------------------------------------'
Public Function CStrNull(ByVal value As Variant, Optional ByVal NullValue As String = "") As String
    If IsNull(value) Then
        CStrNull = NullValue
    Else
        CStrNull = CStr(value)
    End If
End Function

'------------------------------------------------------------------------'
'Function ButLastStr : returns all but the last chars of a string        '
'  E.g: MsgBox ButLastStr("ABCDEFGH",2)                                  '
'       will display 'ABCDEF'                                            '
'------------------------------------------------------------------------'
Public Function ButLastStr(ByVal str As String, Optional ByVal Last As Long = 1) As String
    ButLastStr = Left$(str, Len(str) - Last)
End Function

'------------------------------------------------------------------------'
'Function FromToStr : returns subset of a string                         '
'  E.g: MsgBox FromToStr("ABCDEFGH",2,4)                                 '
'       will display 'BCD'                                               '
'------------------------------------------------------------------------'
Public Function FromToStr(ByVal str As String, ByVal FromOffset As Long, ByVal ToOffset As Long) As String
    Dim c As Long
    
    FromToStr = ""
    If FromOffset > Len(str) Then
        Exit Function
    ElseIf FromOffset < 1 Then
        Error 5
    End If
    
    If ToOffset > Len(str) Then
        ToOffset = Len(str)
    ElseIf ToOffset < 1 Then
        Error 5
    End If

    If ToOffset < FromOffset Then
        Exit Function
    Else
        FromToStr = Mid$(str, FromOffset, ToOffset - FromOffset + 1)
    End If

End Function

'------------------------------------------------------------------------'
'Function ReplaceFirst : replaces the first occurrence of a substring by '
'  another string                                                        '
'  E.g: MsgBox ReplaceFirst("Hi <Name> !","<Name>","John")               '
'       will display 'Hi John !'                                         '
'See also: ReplaceAll, ReplaceMulti                                      '
'------------------------------------------------------------------------'
Public Function ReplaceFirst(ByVal Target As String, ByVal FindStr As String, ByVal ByStr As String, Optional ByVal start As Long = 1) As String
    Dim i As Integer
    
    i = InStr(start, Target, FindStr)
    If i <> 0 Then
        Target = Left$(Target, i - 1) + ByStr + ButLastStr(Target, i + Len(FindStr) - 1)
    End If
    ReplaceFirst = Target

End Function

'------------------------------------------------------------------------'
'Function ReplaceAll : replaces all occurrences of a string by another   '
'  string                                                                '
'  E.g: MsgBox ReplaceAll("Hi <Name> !","<Name>","John")                 '
'       will display 'Hi John !'                                         '
'See also: ReplaceFirst, ReplaceMulti                                    '
'------------------------------------------------------------------------'
Public Function ReplaceAll(ByVal Target As String, ByVal FindStr As String, ByVal ByStr As String, Optional ByVal start As Long = 1) As String
    Dim i As Integer
    
    i = InStr(start, Target, FindStr)
    While i <> 0
        Target = Left(Target, i - 1) + ByStr + Mid(Target, i + Len(FindStr))
        start = i + Len(ByStr)
        i = InStr(start, Target, FindStr)
    Wend
    ReplaceAll = Target

End Function

 '------------------------------------------------------------------------'
 'Function ReplaceMulti : performs multiple string replaces in a one call '
 ' E.g: str = ReplaceMulti(vbTextCompare, "abcd", "a", "1", "c", "3")     '
 '      will assign "1b3d" to str.
 'See also: ReplaceFirst, ReplaceAll
 '------------------------------------------------------------------------'
 Public Function ReplaceMulti(ByVal compare As VbCompareMethod, ByVal str As String, ParamArray replacements() As Variant) As String

    Dim i As Integer
     Dim rfrom As String, rto As String
     For i = LBound(replacements) + 1 To UBound(replacements) Step 2
         rfrom = "" & replacements(i - 1)
         rto = "" & replacements(i)
         str = Replace(str, rfrom, rto, , , compare)
     Next

    ReplaceMulti = str

End Function

'------------------------------------------------------------------------'
'Function FixStr : the given string is shortended or increased in length '
'  (by appending spaces) to match the given length                       '
'------------------------------------------------------------------------'
Public Function FixStr(ByVal str As String, ByVal length As Long)
    If Len(str) < length Then
        FixStr = str + String$(length - Len(str), " ")
    Else
        FixStr = Left$(str, length)
    End If
End Function

'------------------------------------------------------------------------'
'Function FixStrRight : the given string is shortended or increased in   '
'  length (by inserting spaces before) to match the given length         '
'------------------------------------------------------------------------'
Public Function FixStrRight(ByVal str As String, ByVal length As Long)
    If Len(str) < length Then
        FixStrRight = String$(length - Len(str), " ") + str
    Else
        FixStrRight = Right$(str, length)
    End If
End Function

'------------------------------------------------------------------------'
'Function StrUpTo : returns the string up to a given delimiter           '
'------------------------------------------------------------------------'
Public Function StrUpTo(ByVal str As String, ByVal delimiter As String, Optional ByVal start As Variant)
    Dim istart As Long
    Dim ipos As Long
    
    If IsMissing(start) Then istart = 1 Else istart = start
    ipos = InStr(istart, str, delimiter)
    If ipos = 0 Then
        StrUpTo = Mid(str, istart)
    Else
        StrUpTo = FromToStr(str, istart, ipos - 1)
    End If

End Function

'------------------------------------------------------------------------'
'Function StrBetween : returns subtring between two parts                '
'------------------------------------------------------------------------'
Public Function StrBetween(ByVal str As String, ByVal before As String, ByVal after As String, Optional ByVal start As Long = 1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As String
    Dim pos1 As Long
    Dim pos2 As Long
    Dim beforelen As Long
    
    beforelen = Len(before)
    pos1 = InStr(start, str, before, Compare)
    If pos1 = 0 Then
        StrBetween = ""
    Else
        If after = "" Then pos2 = Len(str) + 1 Else pos2 = InStr(pos1 + beforelen, str, after, Compare)
        If pos2 = 0 Then
            StrBetween = ""
        Else
            StrBetween = FromToStr(str, pos1 + beforelen, pos2 - 1)
        End If
    End If

End Function

'------------------------------------------------------------------------'
'Function RemoveDblSpaces : removes double spaces from string            '
'------------------------------------------------------------------------'
Public Function RemoveDblSpaces(ByVal str As String, Optional ByVal keepTabs As Boolean = True, Optional ByVal keepCrlfs As Boolean = True) As String
    Dim s As String
    
    s = str
    If Not keepTabs Then s = Replace(str, vbTab, " ")
    If Not keepCrlfs Then s = Replace(str, vbCrLf, " ")
    While InStr(s, "  ") <> 0
        s = Replace(s, "  ", " ")
    Wend
    RemoveDblSpaces = s

End Function

'------------------------------------------------------------------------'
'Function FirstInstr : returns the index of the instring that occurs as  '
'  first in the searchstring.                                            '
'------------------------------------------------------------------------'
Public Function FirstInstr(ByVal searchstring As String, ByVal Compare As VbCompareMethod, ParamArray instrings() As Variant)
    Dim i As Long
    Dim p As Long
    Dim poss As New Collection
    Dim cand As New Collection
    
    For i = LBound(instrings) To UBound(instrings)
        p = InStr(1, searchstring, instrings(i), Compare)
        If p > 0 Then
            poss.Add p
            cand.Add i + 1
        End If
    Next

    Dim fp As Long
    Dim fc As Long
    For i = 1 To poss.Count
        If i = 1 Then
            fp = poss(i)
            fc = cand(i)
        ElseIf poss(i) < fp Then
            fp = poss(i)
            fc = cand(i)
        End If
    Next

    FirstInstr = fc

End Function

'------------------------------------------------------------------------'
'Function NthInstr : Locates the n'th occurence of a substring within a  '
'  string. Returns 0 if no n occurences are present.                     '
'------------------------------------------------------------------------'
Public Function NthInstr(ByVal str As String, ByVal str2 As String, ByVal n As Integer, Optional ByVal compare As VbCompareMethod = VbCompareMethod.vbBinaryCompare)

    If n < 1 Then
        Err.Raise 5, "StringTools", "Argument n should be > 0."
    End If

    Dim i As Integer
    Dim lastix As Integer
    
    For i = 1 To n
        lastix = Strings.InStr(lastix + 1, str, str2, compare)
        If lastix = 0 Then Exit For
    Next
    
    NthInstr = lastix
    
End Function

'------------------------------------------------------------------------'
'Function StrFollowing : Returns the string part following the given     '
'  delimiter. If the delimiter is not found, returns an empty string.    '
'------------------------------------------------------------------------'
Public Function StrFollowing(ByVal str As String, ByVal delimiter As String) As String

    Dim ix As Integer
    ix = InStr(str, delimiter)
    If ix = 0 Then
        StrFollowing = ""
    Else
        StrFollowing = Mid(str, ix + Len(delimiter))
    End If

End Function

