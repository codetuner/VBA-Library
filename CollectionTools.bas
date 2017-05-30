Attribute VB_Name = "CollectionTools"
'========================================================================'
'== COLLECTIONTOOLS                                                    =='
'==                                                                    =='
'== © Copyright 1997-2002 Rudi Breedenraedt - rudi@breedenraedt.be     =='
'========================================================================'
Option Explicit

'------------------------------------------------------------------------'
'Function CollClone : Returns a new Collection object containing the same'
'  content as the one passed as argument.                                '
'  Note: Keys on the original collection get lost.                       '
'See also: BuildCollection, MergeCollections                             '
'------------------------------------------------------------------------'
Public Function CollClone(ByVal coll As Collection) As Collection
    Dim newcoll As New Collection
    Dim item As Variant
    
    For Each item In coll
        newcoll.Add item
    Next
    Set CollClone = newcoll
    
End Function

'------------------------------------------------------------------------'
'Function BuildCollection : builds a collection containing all parameters'
'  E.g: Set MyColl = BuildCollection(7,21,24,26,33,39)                   '
'See also: ParseCollection, FastParseCollection, AppendCollection        '
'------------------------------------------------------------------------'
Public Function BuildCollection(ParamArray values() As Variant) As Collection
    Dim n As Integer
    Dim c As New Collection
    
    For n = LBound(values) To UBound(values)
        c.Add values(n)
    Next
    Set BuildCollection = c

End Function

'------------------------------------------------------------------------'
'Function AppendCollection : builds a collection based on an existing    '
'    collection and appends further arguments to it                      '
'  E.g: Set MyColl = AppendCollection(HisColl,"John","Linda")            '
'See also: BuildCollection, MergeCollections                             '
'------------------------------------------------------------------------'
Public Function AppendCollection(ByVal coll As Collection, ParamArray values() As Variant) As Collection
    Dim c As New Collection
    Dim item As Variant
    Dim n As Integer
    
    'Copy elements of original collection:
    For Each item In coll
        c.Add item
    Next
    
    'Append additional elements:
    For n = LBound(values) To UBound(values)
        c.Add values(n)
    Next
    Set AppendCollection = c

End Function

'------------------------------------------------------------------------'
'Function MergeCollections : builds a collection by merging two or more  '
'    existing collections                                                '
'  E.g: Set Everyone = MergeCollections(Employees,Customers,Suppliers)   '
'See also: BuildCollection, AppendCollection                             '
'------------------------------------------------------------------------'
Public Function MergeCollections(ParamArray values() As Variant) As Collection
    Dim n As Integer
    Dim c As New Collection
    Dim s As Collection
    Dim i As Variant
    
    For n = LBound(values) To UBound(values)
        Set s = values(n)
        For Each i In s
            c.Add i
        Next
    Next
    Set MergeCollections = c

End Function

'------------------------------------------------------------------------'
'Function ArrayToColl : Converts an array into a collection.             '
'------------------------------------------------------------------------'
Public Function ArrayToColl(Items As Variant) As Collection
    Dim coll As New Collection
    Dim i As Integer
    
    For i = LBound(Items) To UBound(Items)
        coll.Add Items(i)
    Next
    Set ArrayToColl = coll

End Function

'------------------------------------------------------------------------'
'Function CollToArray : Converts a collection into an array.             '
'------------------------------------------------------------------------'
Public Function CollToArray(coll As Collection) As Variant()
    Dim result() As Variant
    Dim i As Integer
    
    ReDim result(0 To coll.Count - 1) As Variant
    For i = 0 To coll.Count - 1
        If IsObject(coll(i + 1)) Then
            Set result(i) = coll(i + 1)
        Else
            result(i) = coll(i + 1)
        End If
    Next
    CollToArray = result
    
End Function

'------------------------------------------------------------------------'
'Function IndexOf : returns the index of an element in a collection (0 if'
'    not found)                                                          '
'------------------------------------------------------------------------'
Public Function IndexOf(ByVal coll As Collection, ByVal item As Variant, Optional ByVal StartIndex As Long = 1) As Long
    Dim collindex As Long
    Dim collitemtype As Integer
    Dim itemtype As Integer
    
    itemtype = VarType(item)
    For collindex = StartIndex To coll.Count
        collitemtype = VarType(coll(collindex))
        If collitemtype = itemtype Then
            Select Case collitemtype
                Case 0 To 1: IndexOf = collindex: Exit Function
                Case 2 To 8, 11, 14, 17: If coll(collindex) = item Then IndexOf = collindex: Exit Function
                Case 9: If coll(collindex) Is item Then IndexOf = collindex: Exit Function
                Case Else
                    Debug.Print "Unsupported type for CollectionTools.IndexOf."
                    Debug.Assert False
            End Select
        End If
    Next
    IndexOf = 0

End Function

'------------------------------------------------------------------------'
'Function Includes : Returns true if the collection includes the given   '
'  value.                                                                '
'------------------------------------------------------------------------'
Public Function Includes(ByVal coll As Collection, item As Variant) As Boolean
    Dim collitem As Variant
    Dim collitemtype As Integer
    Dim itemtype As Integer
    
    'If we leave this function before it's end, it means we found the item:'
    Includes = True
    
    itemtype = VarType(item)
    For Each collitem In coll
        collitemtype = VarType(collitem)
        If collitemtype = itemtype Then
            Select Case collitemtype
                Case 0 To 1: Exit Function
                Case 2 To 8, 11, 14, 17: If collitem = item Then Exit Function
                Case 9: If collitem Is item Then Exit Function
                Case Else
                    Debug.Print "Unsupported type for CollectionTools.Includes."
                    Debug.Assert False
            End Select
        End If
    Next
    
    'If we didn't leave the function yet, it means we didn't find the item:'
    Includes = False
    
End Function

'------------------------------------------------------------------------'
'Function IncludesKey : Returns true if the collection includes the given'
'  key.                                                                  '
'------------------------------------------------------------------------'
Public Function IncludesKey(ByVal coll As Collection, IndexOrKey As Variant) As Boolean
    Dim i As Integer
    
    On Error Resume Next
    i = VarType(coll(IndexOrKey)) 'Dummy assignment'
    If Err = 5 Or Err = 9 Then
        IncludesKey = False
    Else
        IncludesKey = True
    End If

End Function

'------------------------------------------------------------------------'
'Function CollAt : safely returns an element of a collection             '
'  CollAt returns an element of a collection, based on its index or its  '
'  key. If the element is not found, a default value (or Null) is        '
'  returned.                                                             '
'------------------------------------------------------------------------'
Public Function CollAt(ByVal coll As Collection, ByVal IndexOrKey As Variant, Optional ByVal Default As Variant = Null) As Variant
    On Error Resume Next
    
    If IsObject(coll(IndexOrKey)) Then
        Set CollAt = coll(IndexOrKey)
    Else
        CollAt = coll(IndexOrKey)
    End If
    If Err = 5 Or Err = 9 Then
        If IsObject(Default) Then
            Set CollAt = Default
        Else
            CollAt = Default
        End If
    End If
    
End Function

'------------------------------------------------------------------------'
'Function CollAdd : adds an item to a collection, only if there is       '
'  not yet an item with the same key included. Returns True if the item  '
'  was added.                                                            '
'------------------------------------------------------------------------'
Public Function CollAdd(ByVal coll As Collection, ByVal key As String, ByVal item As Variant) As Boolean
    On Error Resume Next
    
    CollAdd = True
    coll.Add item, key
    If Err = 457 Then
        CollAdd = False
    End If

End Function

'------------------------------------------------------------------------'
'Function CollSet : assigns a value to a particular key/index in a       '
'  collection; if the key is not found, it is created. Returns True on   '
'  success.                                                              '
'------------------------------------------------------------------------'
Public Function CollSet(ByVal coll As Collection, ByVal IndexOrKey As Variant, ByVal item As Variant) As Boolean
    On Error Resume Next

    coll.Remove IndexOrKey
    On Error GoTo 0
    If VarType(IndexOrKey) = vbString Then
        coll.Add item, IndexOrKey
    Else
        If coll.Count >= IndexOrKey Then
            coll.Add item, , IndexOrKey
        Else
            coll.Add item
        End If
    End If
    CollSet = True
    
End Function

'------------------------------------------------------------------------'
'Function SetAdd : Adds a value to a collection only if the value is not '
'  yet included in the collection (the collection is a 'set').           '
'  Returns true if the item was added, false if it was already present.  '
'------------------------------------------------------------------------'
Public Function SetAdd(ByVal coll As Collection, ByVal item As Variant) As Boolean
    If Includes(coll, item) Then
        SetAdd = False
    Else
        SetAdd = True
        coll.Add item
    End If
End Function

'------------------------------------------------------------------------'
'Function RemoveValue : Removes a value from a collection. If            '
'  AllOccurrences is true, all occurrences of the value are removed,     '
'  otherwise (by default) only the first occurrence is removed.          '
'  Returns the number of removed items.                                  '
'------------------------------------------------------------------------'
Public Function RemoveValue(ByVal coll As Collection, value As Variant, Optional AllOccurrences As Boolean = False) As Integer
    Dim collindex As Long
    Dim collitemtype As Integer
    Dim itemtype As Integer
    Dim itemsfound As New Collection
    
    itemtype = VarType(value)
    For collindex = 1 To coll.Count
        collitemtype = VarType(coll(collindex))
        If collitemtype = itemtype Then
            Select Case collitemtype
                Case 0 To 1: GoSub RemoveValue_Found
                Case 2 To 8, 11, 14, 17: If coll(collindex) = value Then GoSub RemoveValue_Found
                Case 9: If coll(collindex) Is value Then GoSub RemoveValue_Found
                Case Else
                    Debug.Print "Unsupported type for CollectionTools.IndexOf."
                    Debug.Assert False
            End Select
        End If
    Next
    
RemoveValue_End:
    'Remove items backward to avoid index-shift problems while removing items'
    For collindex = itemsfound.Count To 1 Step -1
        coll.Remove itemsfound(collindex)
    Next
    'Return number of removed items'
    RemoveValue = itemsfound.Count
Exit Function

RemoveValue_Found:
    itemsfound.Add collindex
    If AllOccurrences Then
        Return
    Else
        GoTo RemoveValue_End
    End If
End Function


'------------------------------------------------------------------------'
'Function ParseCollection : builds a collection based on a string        '
'  The FastParse~ works faster then the Parse~ function, but ignores     '
'  nested collections.                                                   '
'  NOTE: ParseCollection does not yet support nested collections since   '
'    it's implementation points to FastParseCollection. In a future,     '
'    ParseCollection may support nested collections.                     '
'  E.g: Set MyColl = ParseCollection("John;Cathy;Bill;Linda")            '
'       Set MyColl = ParseCollection("C:\My Documents\Word\Letters","\") '
'See also: FastParseCollection, ParseDictionary, FastParseDictionary,    '
'  Tokenize                                                              '
'------------------------------------------------------------------------'
Public Function ParseCollection(ByVal str As String, Optional ByVal ElementSeparator As String = ";") As Collection
    Set ParseCollection = FastParseCollection(str, ElementSeparator)
End Function

'------------------------------------------------------------------------'
'Function FastParseCollection : builds a collection based on a string    '
'  The FastParse~ works faster then the Parse~ function, but ignores     '
'  nested collections.                                                   '
'  E.g: Set MyColl = ParseCollection("John;Cathy;Bill;Linda")            '
'       Set MyColl = ParseCollection("C:\My Documents\Word\Letters","\") '
'See also: ParseCollection, ParseDictionary, FastParseDictionary,        '
'  Tokenize                                                              '
'------------------------------------------------------------------------'
Public Function FastParseCollection(ByVal str As String, Optional ByVal ElementSeparator As String = ";") As Collection
    Dim coll As New Collection
    Dim epos As Long
    
    If Len(str) = 0 Then
        Set FastParseCollection = coll
        Exit Function
    End If
    epos = InStr(str, ElementSeparator)
    Do While epos <> 0
        coll.Add Left$(str, epos - 1)
        str = Mid$(str, epos + Len(ElementSeparator))
        epos = InStr(str, ElementSeparator)
    Loop
    coll.Add str
    
    Set FastParseCollection = coll

End Function

'------------------------------------------------------------------------'
'Function ParseDictionary : builds a collection (with keys) based on a   '
'    string.                                                             '
'  The FastParse~ works faster then the Parse~ function, but ignores     '
'  nested collections.                                                   '
'  NOTE: ParseDictionary does not yet support nested collections since   '
'    it's implementation points to FastParseDictionary. In a future,     '
'    ParseDictionary may support nested collections.                     '
'  E.g: Set MyColl = ParseDictionary("Name=John;Age=25;IsHuman")         '
'See also: FastParseDictionary, ParseCollection, FastParseCollection     '
'------------------------------------------------------------------------'
Public Function ParseDictionary(ByVal str As String, Optional ByVal ElementSeparator As String = ";", Optional ByVal KeyValueSeparator As String = "=") As Collection
    Set ParseDictionary = FastParseDictionary(str, ElementSeparator, KeyValueSeparator)
End Function

'------------------------------------------------------------------------'
'Function FastParseDictionary : builds a collection (with keys) based on '
'    a string.                                                           '
'  The FastParse~ works faster then the Parse~ function, but ignores     '
'  nested collections.                                                   '
'  E.g: Set MyColl = FastParseDictionary("Name=John;Age=25;IsHuman")     '
'See also: ParseDictionary, ParseCollection, FastParseCollection         '
'------------------------------------------------------------------------'
Public Function FastParseDictionary(ByVal str As String, Optional ByVal ElementSeparator As String = ";", Optional ByVal KeyValueSeparator As String = "=") As Collection
    Dim coll As New Collection
    Dim Keys As New Collection
    Dim element As Variant
    Dim spos As Long
    
    For Each element In FastParseCollection(str, ElementSeparator)
        spos = InStr(element, KeyValueSeparator)
        If spos > 0 Then
            coll.Add Mid$(element, spos + Len(KeyValueSeparator)), Left$(element, spos - 1)
            Keys.Add Left$(element, spos - 1)
        Else
            If Len(element) > 0 Then
                coll.Add True, element
                Keys.Add element
            End If
        End If
    Next
    coll.Add Keys, "__Keys"
    Set FastParseDictionary = coll
    
End Function

'------------------------------------------------------------------------'
'Function CollKeys : returns the keys of a collection                    '
'  CollKeys assumes the collection was build with [Fast]ParseDictionary, '
'  if this was not the case, or if the collection was altered afterwards,'
'  a collection with all index is returned                               '
'------------------------------------------------------------------------'
Public Function CollKeys(ByVal coll As Collection) As Collection
    Dim key As Variant
    
    If IncludesKey(coll, "__Keys") Then
        If (coll("__Keys").Count + 1) = coll.Count Then
            Set CollKeys = coll("__Keys")
            Exit Function
        Else
            coll.Remove "__Keys"
        End If
    End If

    'If no __Keys collection found, or collection was altered:
    Dim result As New Collection
    For key = 1 To coll.Count
        result.Add key
    Next
    Set CollKeys = result

End Function

'------------------------------------------------------------------------'
'Function Tokenize : tokenizes a string and returns tokens as collection '
'  Tokinization assumes spaces and tabs are separators                   '
'  NOTE: Tokenize currently does not support tokenization of strings,    '
'    accepting stuff between double quotes as one single token, and      '
'    recognize dates between #-chars is for future releases...           '
'See also: ParseCollection, FastParseCollection, ParseDictionary,        '
'  FastParseDictionary                                                   '
'------------------------------------------------------------------------'
Public Function Tokenize(ByVal str As String) As Collection
    Dim coll As New Collection
    Dim token As String
    Dim c As String
    Dim i As Long
    
    token = ""
    For i = 1 To Len(str)
        c = Mid$(str, i, 1)
        Select Case c
        Case "0" To "9", "a" To "z", "A" To "Z", "_"
            token = token + c
        Case " ", Chr$(9)
            If Len(token) <> 0 Then coll.Add token: token = ""
        Case Chr$(10)
            'Do nothing
        Case Chr$(13)
            If Len(token) <> 0 Then coll.Add token: token = ""
            coll.Add vbCrLf
        Case Else
            If Len(token) <> 0 Then coll.Add token: token = ""
            coll.Add c
        End Select
    Next
    If Len(token) <> 0 Then coll.Add token
    Set Tokenize = coll

End Function

'------------------------------------------------------------------------'
'Function FormatCollection : formats a collection into a string          '
'See also: ParseCollection, FastParseCollection                          '
'------------------------------------------------------------------------'
Public Function FormatCollection(ByVal coll As Collection, Optional ByVal ElementSeparator As String = ";", Optional ByVal NullValue As String = "#NULL#") As String
    Dim result As String
    Dim element As Variant
    Dim first As Boolean
    
    first = True
    For Each element In coll
        If Not first Then result = result & ElementSeparator
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
            ElseIf TypeOf element Is Collection Then
                result = result & "(" & FormatCollection(element, ElementSeparator, NullValue) & ")"
            Else
                result = result & "#" & TypeName(element) & "#"
            End If
        Case Else
            result = result & CStr(element)
        End Select
        If first Then first = False
    Next
    FormatCollection = result

End Function

'------------------------------------------------------------------------'
'Function MidColl : returns subset of a collection                       '
'  Similar to Mid$() on strings.                                         '
'See also: FromToColl, LeftColl, RightColl, ButLastColl                  '
'------------------------------------------------------------------------'
Public Function MidColl(ByVal coll As Collection, ByVal start As Long, Optional ByVal length As Variant) As Collection
    Dim result As New Collection
    Dim Count As Long
    Dim c As Long
    
    Count = coll.Count
    If start < 1 Then Error 5
    If start > Count Then
        Set result = result
    Else
        If IsMissing(length) Then
            For c = start To Count
                result.Add coll(c)
            Next
        ElseIf (length >= (Count + 1 - start)) Then
            For c = start To Count
                result.Add coll(c)
            Next
        ElseIf length < 0 Then
            Error 5
        Else
            For c = start To start + Count - 1
                result.Add coll(c)
            Next
        End If
    End If
    Set MidColl = result

End Function

'------------------------------------------------------------------------'
'Function FromToColl : returns subset of a collection                    '
'See also: MidColl, LeftColl, RightColl                                  '
'------------------------------------------------------------------------'
Public Function FromToColl(ByVal coll As Collection, ByVal FromOffset As Long, ByVal ToOffset As Long) As Collection
    Dim result As New Collection
    Dim c As Long
    
    If FromOffset > coll.Count Then
        Set FromToColl = result
        Exit Function
    ElseIf FromOffset < 1 Then
        Error 5
    End If
    
    If ToOffset > coll.Count Then
        ToOffset = coll.Count
    ElseIf ToOffset < 1 Then
        Error 5
    End If

    If ToOffset < FromOffset Then
        Set FromToColl = result
        Exit Function
    Else
        For c = FromOffset To ToOffset
            result.Add coll(c)
        Next
    End If

    Set FromToColl = result

End Function

'------------------------------------------------------------------------'
'Function LeftColl : returns left-most elements of a collection          '
'  Similar to Left$() on strings.                                        '
'See also: MidColl, RightColl, ButLastColl                               '
'------------------------------------------------------------------------'
Public Function LeftColl(ByVal coll As Collection, ByVal length As Long) As Collection
    Dim result As New Collection
    Dim c As Long
    
    If length > coll.Count Then length = coll.Count
    For c = 1 To length
        result.Add coll(c)
    Next
    Set LeftColl = result

End Function

'------------------------------------------------------------------------'
'Function RightColl : returns right-most elements of a collection        '
'  Similar to Right$() on strings.                                       '
'See also: MidColl, LeftColl                                             '
'------------------------------------------------------------------------'
Public Function RightColl(ByVal coll As Collection, ByVal length As Long) As Collection
    Dim result As New Collection
    Dim c As Long
    
    If length > coll.Count Then length = coll.Count
    For c = coll.Count - length + 1 To coll.Count
        result.Add coll(c)
    Next
    Set RightColl = result

End Function

'------------------------------------------------------------------------'
'Function ButLastColl : returns all but the last elements of a collection'
'See also: LeftColl, MidColl                                             '
'------------------------------------------------------------------------'
Public Function ButLastColl(ByVal coll As Collection, Optional ByVal Last As Long = 1) As Collection
    Set ButLastColl = LeftColl(coll, coll.Count - Last)
End Function

