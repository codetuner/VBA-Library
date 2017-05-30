VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InOutFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim mFilehandle As Integer
Dim mFilename As String

Public Property Get Filename() As String
    Filename = mFilename
End Property

Public Property Let Filename(ByVal vNewValue As String)
    mFilename = vNewValue
End Property

Public Sub OpenWrite(Optional ByVal pFilename As Variant)

    If Not IsMissing(pFilename) Then Filename = pFilename
    CloseFile
    
    mFilehandle = FreeFile()
    Open mFilename For Output As mFilehandle

End Sub

Public Sub OpenAppend(Optional ByVal pFilename As Variant)

    If Not IsMissing(pFilename) Then Filename = pFilename
    CloseFile
    
    mFilehandle = FreeFile()
    Open mFilename For Append As mFilehandle

End Sub

Public Sub OpenBinary(Optional ByVal pFilename As Variant)

    If Not IsMissing(pFilename) Then Filename = pFilename
    CloseFile
    
    mFilehandle = FreeFile()
    Open mFilename For Binary As mFilehandle

End Sub

Public Sub OpenRead(Optional ByVal pFilename As Variant)

    If Not IsMissing(pFilename) Then Filename = pFilename
    CloseFile
    
    mFilehandle = FreeFile()
    Open mFilename For Input As mFilehandle

End Sub

Public Sub CloseFile()
    If mFilehandle <> 0 Then Close #mFilehandle
    mFilehandle = 0
End Sub

Public Sub FWrite(ByVal text As String)
    Print #mFilehandle, text;
End Sub

Public Sub FWriteLn(ByVal text As String)
    Print #mFilehandle, text
End Sub

