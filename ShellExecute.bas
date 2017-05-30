Attribute VB_Name = "ShellExecute"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
#Else
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
#End If


Private Const SW_SHOWNORMAL          As Long = 1
Private Const SE_ERR_FNF             As Long = 2
Private Const SE_ERR_PNF             As Long = 3
Private Const SE_ERR_ACCESSDENIED    As Long = 5
Private Const SE_ERR_OOM             As Long = 8
Private Const SE_ERR_DLLNOTFOUND     As Long = 32
Private Const SE_ERR_SHARE           As Long = 26
Private Const SE_ERR_ASSOCINCOMPLETE As Long = 27
Private Const SE_ERR_DDETIMEOUT      As Long = 28
Private Const SE_ERR_DDEFAIL         As Long = 29
Private Const SE_ERR_DDEBUSY         As Long = 30
Private Const SE_ERR_NOASSOC         As Long = 31
Private Const ERROR_BAD_FORMAT       As Long = 11
 
Public Sub Execute(ByVal Document As String, Optional ByVal Command As String = "OPEN", Optional ByVal Parameters As String = "", Optional ByVal ShowError As Boolean = False)
    Dim Scr_hDC As Long
    Dim rtn     As Long
    
    Scr_hDC = GetDesktopWindow()
    rtn = ShellExecute(Scr_hDC, Command, Document, Parameters, "C:\", SW_SHOWNORMAL)
    
    If ShowError Then
       ShowErrorMessage rtn
    End If
End Sub
 
Private Sub ShowErrorMessage(r As Long)
    Dim s As String
    
    If r <= 32 Then
        'There was an error
        Select Case r
            Case SE_ERR_FNF
                s = "File not found"
            Case SE_ERR_PNF
                s = "Path not found"
            Case SE_ERR_ACCESSDENIED
                s = "Access denied"
            Case SE_ERR_OOM
                s = "Out of memory"
            Case SE_ERR_DLLNOTFOUND
                s = "DLL not found"
            Case SE_ERR_SHARE
                s = "A sharing violation occurred"
            Case SE_ERR_ASSOCINCOMPLETE
                s = "Incomplete or invalid file association"
            Case SE_ERR_DDETIMEOUT
                s = "DDE Time out"
            Case SE_ERR_DDEFAIL
                s = "DDE transaction failed"
            Case SE_ERR_DDEBUSY
                s = "DDE busy"
            Case SE_ERR_NOASSOC
                s = "No association for file extension"
            Case ERROR_BAD_FORMAT
                s = "Invalid EXE file or error in EXE image"
            Case Else
                s = "Unknown error"
        End Select
        MsgBox s, vbInformation
    End If
End Sub


