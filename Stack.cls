VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mColl As New Collection

Public Function Push(ByVal v As Variant) As Variant
    mColl.Add v
    If IsObject(v) Then
        Set Push = v
    Else
        Let Push = v
    End If
End Function

Public Function Pop(Optional IfMissing As Variant = Nothing) As Variant
    If mColl.Count = 0 Then
        Set Pop = IfMissing
        Exit Function
    ElseIf IsObject(mColl(mColl.Count)) Then
        Set Pop = mColl(mColl.Count)
    Else
        Let Pop = mColl(mColl.Count)
    End If
    mColl.Remove (mColl.Count)
End Function

Public Function Peek(Optional ByVal n As Long = 1, Optional IfMissing As Variant = Nothing) As Variant
    If mColl.Count < n Then
        Set Peek = IfMissing
    ElseIf IsObject(mColl(mColl.Count - n + 1)) Then
        Set Peek = mColl(mColl.Count - n + 1)
    Else
        Let Peek = mColl(mColl.Count - n + 1)
    End If
End Function

Public Function IsEmpty() As Boolean
    IsEmpty = (mColl.Count = 0)
End Function

Public Function Depth() As Long
    Depth = mColl.Count
End Function

Public Sub Restore(ByVal toDepth As Long)
    If Depth < toDepth Then Err.Raise 5
    While Depth > toDepth
        mColl.Remove (mColl.Count)
    Wend
End Sub

