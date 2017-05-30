Attribute VB_Name = "FileTools"
Option Explicit

'Uses CollectionTools

'------------------------------------------------------------------------'
'Function MergePaths : merges two portions of a path, returns the merged '
'  path as a string.                                                     '
'  E.g: filename = MergePaths("C:\Temp", "..\AUTOEXEC.BAT")              '
'------------------------------------------------------------------------'
Function MergePaths(ByVal BasePath As String, ByVal Extension As String) As String
    Dim bPath As Collection
    Dim xPath As Collection
    Dim pItem As Variant
    
    If Len(BasePath) = 0 Then
        MergePaths = Extension
        Exit Function
    End If
    
    If Mid(Extension & " ", 2, 1) = ":" Then
        'Extension is an absolute path and cannot be merged:
        MergePaths = Extension
        Exit Function
    End If

    Set bPath = FastParseCollection(BasePath, "\")
    Set xPath = FastParseCollection(Extension, "\")
    If Len(RightColl(bPath, 1)(1)) = 0 Then Set bPath = LeftColl(bPath, bPath.Count - 1)
    For Each pItem In xPath
        If pItem = "." Then
            'Just skip item
        ElseIf pItem = ".." Then
            If bPath.Count > 1 Then
                Set bPath = LeftColl(bPath, bPath.Count - 1)
            Else
                Err.Raise 5, "MergePaths", "Paths can not be merged."
            End If
        Else
            bPath.Add pItem
        End If
    Next

    MergePaths = FormatCollection(bPath, "\")

End Function

'------------------------------------------------------------------------'
'Function ParentFolder : For a file, returns its folder, for a folder,   '
'  returns its parent folder.                                            '
'------------------------------------------------------------------------'
Function ParentFolder(ByVal Folder As String) As String
    If Right(Folder, 1) = "\" Then Folder = Left(Folder, Len(Folder) - 1)
    ParentFolder = FormatCollection(ButLastColl(FastParseCollection(Folder, "\"), 1), "\")
End Function

'------------------------------------------------------------------------'
'Function FolderExists : checks whether a given folder exists            '
'------------------------------------------------------------------------'
Function FolderExists(ByVal Folder As String) As Boolean
    If Right(Folder, 1) = "\" Then Folder = Left(Folder, Len(Folder) - 1)
    FolderExists = (Dir(Folder, vbDirectory + vbArchive + vbHidden + vbSystem) <> "") And (Dir(Folder, vbArchive + vbHidden + vbSystem) = "")
End Function

'------------------------------------------------------------------------'
'Function FileExists : checks whether a given file exists                '
'------------------------------------------------------------------------'
Public Function FileExists(ByVal filename As String) As Boolean
    If Right(filename, 1) = "\" Then filename = Left(filename, Len(filename) - 1)
    FileExists = (Dir(filename, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "")
End Function

'------------------------------------------------------------------------'
'Function LoadFile : Loads an entire file into a string                  '
'------------------------------------------------------------------------'
Public Function LoadFile(ByVal filename As String) As Variant
    Dim fh As Integer
    Dim Contents As String
    
    If FileExists(filename) Then
        fh = FreeFile()
        Open filename For Binary Access Read As #fh
        Contents = String(LOF(fh), 0)
        Get #fh, , Contents
        Close #fh
        LoadFile = Contents
    Else
        LoadFile = Null
    End If
    
End Function

'------------------------------------------------------------------------'
'Function WriteFile : Writes an entire file based on a string            '
'------------------------------------------------------------------------'
Public Function WriteFile(ByVal filename As String, ByVal Contents As String) As Boolean
    Dim fh As Integer
    
    On Error GoTo WriteFileHandler
    fh = FreeFile()
    Open filename For Output Access Write As #fh
    Close #fh
    Open filename For Binary Access Write As #fh
    Put #fh, , Contents
    Close #fh
    WriteFile = True
Exit Function

WriteFileHandler:
    WriteFile = False
End Function

'------------------------------------------------------------------------'
'Sub MkDirDeep : builds a complete path.                                 '
'  E.g: MkDirDeep "C:\Program Files\MyCompany\MyProgram\Settings"        '
'------------------------------------------------------------------------'
Sub MkDirDeep(ByVal Folder As String)
    If Folder = "" Then
        Exit Sub
    ElseIf Not FolderExists(Folder) Then
        MkDirDeep ParentFolder(Folder)
        MkDir Folder
    End If
End Sub
