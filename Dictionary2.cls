VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mKeys As New Collection
Private mItems As New Collection

Private Function IndexOfKey(Key As Variant) As Long
Attribute IndexOfKey.VB_Description = "Returns the internal index position of an item for a given key."
    Dim i As Long
    
    For i = 1 To mKeys.Count
        If mKeys(i) = Key Then
            IndexOfKey = i
            Exit Function
        End If
    Next
    IndexOfKey = 0

End Function

Public Sub Add(Key As Variant, Item As Variant)
Attribute Add.VB_Description = "Adds an item to the dictionary."
    Dim i As Long
    
    i = IndexOfKey(Key)
    If i = 0 Then
        mKeys.Add Key
        mItems.Add Item
    Else
        Error 5
    End If

End Sub

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the dictionary."
    Count = mKeys.Count
End Property

Public Function Exists(Key As Variant) As Boolean
Attribute Exists.VB_Description = "Returns whether an item exists for the given key."
    Exists = (IndexOfKey(Key) <> 0)
End Function

Public Sub Remove(Key As Variant)
Attribute Remove.VB_Description = "Removes the value for a given key."
    Dim i As Long
    
    i = IndexOfKey(Key)
    If i <> 0 Then
        mKeys.Remove i
        mItems.Remove i
    End If
    
End Sub

Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Removes all items in the dictionary."
    Set mKeys = New Collection
    Set mItems = New Collection
End Sub

Public Function Keys() As Collection
Attribute Keys.VB_Description = "Returns a collection with all keys of this dictionary."
    Set Keys = mKeys
End Function

Public Function Items() As Collection
Attribute Items.VB_Description = "Returns a collection with all items of this dictionary."
    Set Items = mItems
End Function

Public Property Get Item(Key As Variant) As Variant
Attribute Item.VB_Description = "Returns/sets the value for a given key."
    Dim i As Long
    
    i = IndexOfKey(Key)
    If i = 0 Then
        'Item is nothing
    Else
        If IsObject(mItems(i)) Then
            Set Item = mItems(i)
        Else
            Item = mItems(i)
        End If
    End If

End Property

Public Property Let Item(Key As Variant, Item As Variant)
    Dim i As Long
    
    i = IndexOfKey(Key)
    If i = 0 Then
        mKeys.Add Key
        mItems.Add Item
    Else
        mItems.Add Item, , , i
        mItems.Remove i
    End If

End Property

Public Property Set Item(Key As Variant, Item As Variant)
    Dim i As Long
    
    i = IndexOfKey(Key)
    If i = 0 Then
        mKeys.Add Key
        mItems.Add Item
    Else
        mItems.Add Item, , , i
        mItems.Remove i
    End If

End Property

Public Property Let Key(NewKey As Variant)
Attribute Key.VB_Description = "Changes a key (write-only property)."
    Dim i As Long
    
    i = IndexOfKey(NewKey)
    If i = 0 Then
        Error 5
    Else
        mKeys.Add NewKey, , , i
        mKeys.Remove i
    End If

End Property

Public Property Set Key(NewKey As Variant)
    Dim i As Long
    
    i = IndexOfKey(NewKey)
    If i = 0 Then
        Error 5
    Else
        mKeys.Add NewKey, , , i
        mKeys.Remove i
    End If

End Property

