Attribute VB_Name = "IniFileReader"
Option Explicit

'Sample usage, retrieves the value of Name in the section [User] :
'  Dim settings As Collection
'  Set settings = ReadIniFile("my.ini")
'  MsgBox("Hello " + settings("User")("Name"))

Private Declare Function w32_GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function w32_WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ReadIniFile(ByVal filename As String) As Collection
    Dim inifile As New Collection
    Dim section As Collection
    Dim sectionname As Variant
    Dim keyname As Variant
    Dim newcoll As Collection
    
    inifile.Add filename, "__Filename"
    Set newcoll = New Collection: inifile.Add newcoll, "__Sections"
    For Each sectionname In ReadIniSections(filename)
        inifile("__Sections").Add sectionname
        Set section = New Collection
        Set newcoll = New Collection: section.Add newcoll, "__Keys"
        For Each keyname In ReadIniSectionKeys((sectionname), filename)
            section("__Keys").Add keyname
            section.Add ReadIniSettingString((sectionname), (keyname), CStr(True), filename), keyname
        Next
        inifile.Add section, sectionname
    Next
    
    Set ReadIniFile = inifile

End Function

Public Sub WriteIniFile(iniobject As Collection, Optional filename As Variant)
    Dim filename2 As Variant
    Dim section As Variant
    Dim key As Variant
    
    If Not IsMissing(filename) Then
        filename2 = filename
        iniobject("__Filename") = filename
    Else
        filename2 = CollAt(iniobject, "__Filename", Null)
    End If
    If IsNull(filename2) Then
        Err.Raise 5, , "No filename to write INI file to."
        Exit Sub
    End If
    
    For Each section In CollAt(iniobject, "__Sections", New Collection)
        For Each key In CollAt(iniobject(section), "__Keys", New Collection)
            WriteIniSettingString section, key, CStr(CollAt(iniobject(section), key, "")), filename2
        Next
    Next

End Sub

Public Function GetIniSetting(iniobject As Collection, section As String, key As String, Optional ByVal Default As Variant = Null) As Variant
    Dim sec As Collection
    Dim Keys As Collection
    
    If IsObject(Default) Then
        Err.Raise 5, , "No object can be passed to GetIniSetting method as default value."
    End If
    Set sec = CollAt(iniobject, section, Nothing)
    If sec Is Nothing Then
        GetIniSetting = Default
    Else
        GetIniSetting = CollAt(sec, key, Default)
    End If
    
End Function

Public Sub SetIniSetting(iniobject As Collection, section As String, key As String, Value As Variant)
    Dim sec As Collection
    Dim Keys As Collection
    
    If IsObject(Value) Then
        Err.Raise 5, , "No object can be passed to SetIniSetting method as value."
    End If
    If Not Includes(iniobject("__Sections"), section) Then
        Set sec = New Collection
        iniobject.Add sec, section
        iniobject("__Sections").Add section
        Set Keys = New Collection
        sec.Add Keys, "__Keys"
    Else
        Set sec = iniobject(section)
        Set Keys = iniobject(section)("__Keys")
    End If
    If Not Includes(Keys, key) Then
        Keys.Add key
    End If
    
    CollSet sec, key, Value
    
End Sub

Public Function ReadIniSections(Optional pvsFileName As Variant) As Collection
    Dim coll As New Collection
    Dim Items As Collection
    Dim item As Variant
    Dim s As String
    
    s = ReadIniSettingString(vbNullString, vbNullString, "", pvsFileName)
    Set Items = FastParseCollection(s, Chr$(0))
    For Each item In Items
        If Len(item) > 0 Then coll.Add item
    Next
    Set ReadIniSections = coll

End Function

Public Function ReadIniSectionKeys(psSectionName As String, Optional pvsFileName As Variant) As Collection
    Dim coll As New Collection
    Dim Items As Collection
    Dim item As Variant
    Dim s As String
    
    s = ReadIniSettingString(psSectionName, vbNullString, "", pvsFileName)
    Set Items = FastParseCollection(s, Chr$(0))
    For Each item In Items
        If Len(item) > 0 Then coll.Add item
    Next
    Set ReadIniSectionKeys = coll

End Function

Private Function ReadIniSettingString( _
  psSectionName As String, _
  psKeyName As String, _
  Optional pvsDefault As Variant, _
  Optional pvsFileName As Variant) As String
'********************
' Purpose:    Get a string from a private .ini file
' Parameters:
' (Input)
'   psApplicationName - the Application name
'   psKeyName - the key (section) name
'   pvsDefault - Default value if key not found (optional)
'   pvsFileName - the name of the .ini file
' Returns:  The requested value
' Notes:
'   If no value is provided for pvsDefault, a zero-length string is used
'   The file path defaults to the windows directory if not fully qualified
'   If pvsFileName is omitted, win.ini is used
'   If vbNullString is passed for psKeyName, the entire section is returned in
'     the form of a multi-c-string. Use MultiCStringToStringArray to parse it after appending the
'     second null terminator that this function strips. Note that the value returned is all the
'     key names and DOES NOT include all the values. This can be used to setup multiple calls for
'     the values ala the Reg enumeration functions.
'********************

  ' call params
  Dim lpSectionName As String
  Dim lpKeyName As String
  Dim lpDefault As String
  Dim lpReturnedString As String
  Dim nSize As Long
  Dim lpFileName As String
  ' results
  Dim lResult As Long
  Dim sResult As String
  
  sResult = ""
  
  ' setup API call params
  nSize = 256
  lpReturnedString = Space$(nSize)
  lpSectionName = psSectionName
  lpKeyName = psKeyName
  ' check for value in file name
  If Not IsMissing(pvsFileName) Then
    lpFileName = CStr(pvsFileName)
  Else
    lpFileName = "win.ini"
  End If
  ' check for value in optional pvsDefault
  If Not IsMissing(pvsDefault) Then
    lpDefault = CStr(pvsDefault)
  Else
    lpDefault = ""
  End If
  ' call
  ' setup loop to retry if result string too short
  Do
    lResult = w32_GetPrivateProfileString( _
        lpSectionName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFileName)
    ' Note: See docs for GetPrivateProfileString API
    ' the function returns nSize - 1 if a key name is provided but the buffer is too small
    ' the function returns nSize - 2 if no key name is provided and the buffer is too small
    ' we test for those specific cases - this method is a bit of hack, but it works.
    ' the result is that the buffer must be at least three characters longer than the
    ' longest string(s)
    If (lResult = nSize - 1) Or (lResult = nSize - 2) Then
      nSize = nSize * 2
      lpReturnedString = Space$(nSize)
    Else
      sResult = Left$(lpReturnedString, lResult)
      Exit Do
    End If
  Loop
  
  ReadIniSettingString = sResult
  
End Function

Private Function WriteIniSettingString( _
  ByVal psSectionName As String, _
  ByVal psKeyName As String, _
  ByVal psValue As String, _
  ByVal psFileName As String) As Boolean
'********************
' Purpose:    Write a string to an ini file
' Parameters: (Input Only)
'   psApplicationName - the ini section name
'   psKeyName - the ini key name
'   psValue - the value to write to the key
'   psFileName - the ini file name
' Returns:    True if successful
' Notes:
'   Path defaults to windows directory if the file name
'   is not fully qualified
'********************

  Dim lResult As Long
  Dim fRV As Boolean
  
  lResult = w32_WritePrivateProfileString( _
      psSectionName, _
      psKeyName, _
      psValue, _
      psFileName)
  If lResult <> 0 Then
    fRV = True
  Else
    fRV = False
  End If
  
  WriteIniSettingString = fRV
  
End Function

Private Function StringUpTo(str As String, UpTo As String) As String
    Dim i As Integer
    
    i = InStr(str, UpTo)
    If i = 0 Then
        StringUpTo = str
    Else
        StringUpTo = Left$(str, i - 1)
    End If

End Function
