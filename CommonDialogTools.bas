Attribute VB_Name = "CommonDialogTools"
Option Explicit

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHOWHELP = &H10
Private Const OFS_MAXPATHNAME = 128
'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not a standard Win95 type.
Private Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS
Private Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
Private Type OPENFILENAME
    nStructSize As Long
#If VBA7 Then
    hwndOwner As LongPtr
    hInstance As LongPtr
#Else
    hwndOwner As Long
    hInstance As Long
#End If
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
#If VBA7 Then
    fnHook As LongPtr
#Else
    fnHook As Long
#End If
    sTemplateName As String
End Type
Private OFN As OPENFILENAME

#If VBA7 Then
Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare PtrSafe Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
#Else
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
#End If



Public Function FileOpen(ByVal DlgTitle As String, ByVal Filetypes As String) As String

  Dim xx1 As Long, xx2 As Long, defaultExtension As String
#If VBA7 Then
  Dim r As LongPtr
#Else
  Dim r As Long
#End If
  Dim sp As Long
  Dim LongName As String
  Dim shortName As String
  Dim ShortSize As Long
  'to keep lines short(er), I've abbreviated a
  'Null$ to n and n2, and the filter$ to f.
  Dim n As String
  Dim n2 As String
  Dim f As String
  n = Chr$(0)
  n2 = n & n
  
  '------------------------------------------------
  'INITIALIZATION
  '------------------------------------------------
  'fill in the size of the OFN structure
  OFN.nStructSize = LenB(OFN)
  'assign the owner of the dialog; this can be null if no owner.
  OFN.hwndOwner = 0
  
  '------------------------------------------------
  'FILTERS
  '------------------------------------------------
  'There are 2 methods of setting filters (patterns) for
  'use in the dropdown combo of the dialog.
  'The first, using OFN.sFilter, fills the combo with the
  'specified filters, and works as the VB common dialog does.
  'These must be in the "Friendly Name"-null$-Extension format,
  'terminating with 2 null strings.
  If Filetypes = "" Then Filetypes = "Microsoft Access Databases (*.mdb)|*.mdb|"
  If Right$(Filetypes, 1) <> "|" Then Filetypes = Filetypes + "|"
  While InStr(Filetypes, "|") <> 0
    Mid$(Filetypes, InStr(Filetypes, "|"), 1) = Chr$(0)
  Wend
  f = Filetypes & "All Files" & n & "*.*" & n2
  OFN.sFilter = f
  xx1 = InStr(f, Chr$(0))
  xx2 = InStr(xx1 + 1, f, Chr$(0))
  defaultExtension = Mid$(f, xx1 + 1, xx2 - xx1 - 1)
  'The second method, uses sCustomFilter and nCustFilterSize to pass
  'the filters to use and the size of the filter string.
  'The operating system copies the strings to the buffer when the
  'user closes the dialog box. The system uses the strings
  'to initialize the user-defined file filter the next time the
  'dialog box is created. If this parameter is NULL, the dialog
  'box lists but does not save user-defined filter strings.
  'To see the difference, comment out the line
  'OFN.sFilter = f above, and uncomment the 2 lines below.
  ' OFN.sCustomFilter = f
  ' OFN.nCustFilterSize = Len(OFN.sCustomFilter)
  'nFilterIndex specifies an index into the buffer pointed to by sFilter.
  'The system uses the index value to obtain a pair of strings to use
  'as the initial filter description and filter pattern for the dialog box.
  'The first pair of strings has an index value of 1. When the user closes
  'the dialog box, the system copies the index of the selected filter strings
  'into this location.
  OFN.nFilterIndex = 1
 
  '------------------------------------------------
  'FILENAME
  '------------------------------------------------
  'sFile points to a buffer that contains a filename used to initialize
  'the File Name edit control. The first character of this buffer must be
  'NULL if initialization is not necessary. When the GetOpenFileName
  'or GetSaveFileName function returns, this buffer contains the drive
  'designator, path, filename, and extension of the selected file.
  'perform no filename initialization (Filename textbox is blank)
  'and initialize the sFile buffer for the return value
  ' OFN.sFile = Chr$(0)
  ' OFN.sFile = Space$(1024)
  'OR
  'pass a default filename and initialize for return value
  OFN.sFile = String$(1024, 0)
  OFN.nFileSize = Len(OFN.sFile)
  'default extension applied to a selected file if it has no extension.
  OFN.sDefFileExt = defaultExtension
  'sFileTitle points to a buffer that receives the title of the
  'selected file. The application should use this string
  'to display the file title. If this member is NULL, the
  'function does not copy the file title.
  OFN.sFileTitle = String$(512, 0)
  OFN.nTitleSize = Len(OFN.sFileTitle)
  'sInitDir is the string that specifies the initial
  'file directory. If this member is NULL, the system
  'uses the current directory as the initial directory.
  OFN.sInitDir = ""
  
  '------------------------------------------------
  'MISC
  '------------------------------------------------
  'sDlgTitle is the title to display in the dialog. If null
  'the default title for the dialog is used.
  OFN.sDlgTitle = DlgTitle
  'flags are the actions and options for the dialog.
  OFN.flags = OFS_FILE_OPEN_FLAGS
  'Finally, show the File Open Dialog
  r = GetOpenFileName(OFN)
  
  '------------------------------------------------
  'RESULTS
  '------------------------------------------------
  If r Then
    FileOpen = Left$(OFN.sFile, InStr(OFN.sFile, Chr$(0)) - 1)
    'Path & File Returned (OFN.sFile):
'     Text1 = OFN.sFile
    'File Path (from OFN.nFileOffset):
'     Text2 = Left$(OFN.sFile, OFN.nFileOffset)
    'File Name (from OFN.nFileOffset):
'     Text3 = Mid$(OFN.sFile, OFN.nFileOffset + 1, Len(OFN.sFile) - OFN.nFileOffset - 1)
    'Extension (from OFN.nFileExt):
'     Text4 = Mid$(OFN.sFile, OFN.nFileExt + 1, Len(OFN.sFile) - OFN.nFileExt)
    'File Name (OFN.sFileTitle):
'     Text5 = OFN.sFileTitle
    'Short 8.3 File Name (using (OFN.sFileTitle):
'     LongName = OFN.sFileTitle
'     shortName = Space$(128)
'     ShortSize = Len(shortName)
'     sp = GetShortPathName(LongName, shortName, ShortSize)
'     Text6 = Left$(shortName, sp)
    'Short 8.3 File Name (using OFN.sFile):
'     LongName = OFN.sFile
'     shortName = Space$(128)
'     ShortSize = Len(shortName)
'     sp = GetShortPathName(LongName, shortName, ShortSize)
'     Text7 = Left$(shortName, sp)
    'User Requested this file be opened as Read Only:
'     chkReadOnly.Value = Abs((OFN.flags And OFN_READONLY) = OFN_READONLY)
  End If

End Function

Public Function FileSave(ByVal DlgTitle As String, ByVal Filetypes As String) As String

  Dim xx1 As Long, xx2 As Long, defaultExtension As String

  Dim r As Long
  Dim sp As Long
  Dim LongName As String
  Dim shortName As String
  Dim ShortSize As Long
  'to keep lines short(er), I've abbreviated a
  'Null$ to n and n2, and the filter$ to f.
  Dim n As String
  Dim n2 As String
  Dim f As String
  n = Chr$(0)
  n2 = n & n
    
  '------------------------------------------------
  'INITIALIZATION
  '------------------------------------------------
  'fill in the size of the OFN structure
  OFN.nStructSize = Len(OFN)
  'assign the owner of the dialog; this can be null if no owner.
  OFN.hwndOwner = 0
  
  '------------------------------------------------
  'FILTERS
  '------------------------------------------------
  'There are 2 methods of setting filters (patterns) for
  'use in the dropdown combo of the dialog.
  'The first, using OFN.sFilter, fills the combo with the
  'specified filters, and works as the VB common dialog does.
  'These must be in the "Friendly Name"-null$-Extension format,
  'terminating with 2 null strings.
  If Filetypes = "" Then Filetypes = "Microsoft Access Databases (*.mdb)|*.mdb|"
  If Right$(Filetypes, 1) <> "|" Then Filetypes = Filetypes + "|"
  While InStr(Filetypes, "|") <> 0
    Mid$(Filetypes, InStr(Filetypes, "|"), 1) = Chr$(0)
  Wend
  f = Filetypes & "All Files" & n & "*.*" & n2
  OFN.sFilter = f
  xx1 = InStr(f, Chr$(0))
  xx2 = InStr(xx1 + 1, f, Chr$(0))
  defaultExtension = Mid$(f, xx1 + 1, xx2 - xx1 - 1)
  'The second method, uses sCustomFilter and nCustFilterSize to pass
  'the filters to use and the size of the filter string.
  'The operating system copies the strings to the buffer when the
  'user closes the dialog box. The system uses the strings
  'to initialize the user-defined file filter the next time the
  'dialog box is created. If this parameter is NULL, the dialog
  'box lists but does not save user-defined filter strings.
  'To see the difference, comment out the line
  'OFN.sFilter = f above, and uncomment the 2 lines below.
  ' OFN.sCustomFilter = f
  ' OFN.nCustFilterSize = Len(OFN.sCustomFilter)
  'nFilterIndex specifies an index into the buffer pointed to by sFilter.
  'The system uses the index value to obtain a pair of strings to use
  'as the initial filter description and filter pattern for the dialog box.
  'The first pair of strings has an index value of 1. When the user closes
  'the dialog box, the system copies the index of the selected filter strings
  'into this location.
  OFN.nFilterIndex = 1
 
  '------------------------------------------------
  'FILENAME
  '------------------------------------------------
  'sFile points to a buffer that contains a filename used to initialize
  'the File Name edit control. The first character of this buffer must be
  'NULL if initialization is not necessary. When the GetOpenFileName
  'or GetSaveFileName function returns, this buffer contains the drive
  'designator, path, filename, and extension of the selected file.
  'perform no filename initialization (Filename textbox is blank)
  'and initialize the sFile buffer for the return value
  ' OFN.sFile = Chr$(0)
  ' OFN.sFile = Space$(1024)
  'OR
  'pass a default filename and initialize for return value
  OFN.sFile = String$(1024, 0)
  OFN.nFileSize = Len(OFN.sFile)
  'default extension applied to a selected file if it has no extension.
  OFN.sDefFileExt = defaultExtension
  'sFileTitle points to a buffer that receives the title of the
  'selected file. The application should use this string
  'to display the file title. If this member is NULL, the
  'function does not copy the file title.
  OFN.sFileTitle = String$(512, 0)
  OFN.nTitleSize = Len(OFN.sFileTitle)
  'sInitDir is the string that specifies the initial
  'file directory. If this member is NULL, the system
  'uses the current directory as the initial directory.
  OFN.sInitDir = ""
  
  '------------------------------------------------
  'MISC
  '------------------------------------------------
  'sDlgTitle is the title to display in the dialog. If null
  'the default title for the dialog is used.
  OFN.sDlgTitle = DlgTitle
  'flags are the actions and options for the dialog.
  OFN.flags = OFS_FILE_SAVE_FLAGS
  'Finally, show the File Open Dialog
  r = GetSaveFileName(OFN)
  
  '------------------------------------------------
  'RESULTS
  '------------------------------------------------
  If r Then
    FileSave = Left$(OFN.sFile, InStr(OFN.sFile, Chr$(0)) - 1)
    'Path & File Returned (OFN.sFile):
'     Text1 = OFN.sFile
    'File Path (from OFN.nFileOffset):
'     Text2 = Left$(OFN.sFile, OFN.nFileOffset)
    'File Name (from OFN.nFileOffset):
'     Text3 = Mid$(OFN.sFile, OFN.nFileOffset + 1, Len(OFN.sFile) - OFN.nFileOffset - 1)
    'Extension (from OFN.nFileExt):
'     Text4 = Mid$(OFN.sFile, OFN.nFileExt + 1, Len(OFN.sFile) - OFN.nFileExt)
    'File Name (OFN.sFileTitle):
'     Text5 = OFN.sFileTitle
    'Short 8.3 File Name (using (OFN.sFileTitle):
'     LongName = OFN.sFileTitle
'     shortName = Space$(128)
'     ShortSize = Len(shortName)
'     sp = GetShortPathName(LongName, shortName, ShortSize)
'     Text6 = Left$(shortName, sp)
    'Short 8.3 File Name (using OFN.sFile):
'     LongName = OFN.sFile
'     shortName = Space$(128)
'     ShortSize = Len(shortName)
'     sp = GetShortPathName(LongName, shortName, ShortSize)
'     Text7 = Left$(shortName, sp)
    'User Requested this file be opened as Read Only:
'     chkReadOnly.Value = Abs((OFN.flags And OFN_READONLY) = OFN_READONLY)
  End If


End Function



