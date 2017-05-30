VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPreviousInChain As Registry
Private mRoot As RegRoot
Private mPath As String
Private tempKey As Long

Public Enum RegRoot
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

Public Enum RegValueType
    REG_NONE = (0)                         'No value type
    REG_SZ = (1)                           'Unicode nul terminated string
    REG_EXPAND_SZ = (2)                    'Unicode nul terminated string w/enviornment var
    REG_BINARY = (3)                       'Free form binary
    reg_dword = (4)                        '32-bit number
    reg_dword_little_endian = (4)          '32-bit number (same as REG_DWORD)
    reg_dword_big_endian = (5)             '32-bit number
    REG_LINK = (6)                         'Symbolic Link (unicode)
    REG_MULTI_SZ = (7)                     'Multiple Unicode strings
    REG_RESOURCE_LIST = (8)                'Resource list in the resource map
    REG_FULL_RESOURCE_DESCRIPTOR = (9)     'Resource list in the hardware description
    REG_RESOURCE_REQUIREMENTS_LIST = (10)
End Enum

'========================================================================='
' Reg Key Security Options
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK))

'Open/Create Options
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1

'Key creation/open disposition
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_OPENED_EXISTING_KEY = &H2

'masks for the predefined standard access types
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF

'Define severity codes
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_ACCESS_DENIED = 5
Private Const ERROR_INVALID_DATA = 13&
Private Const ERROR_MORE_DATA = 234 '  dderror
Private Const ERROR_NO_MORE_ITEMS = 259


'Structures Needed For Registry Prototypes
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

'Registry Function Prototypes
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, szData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExByte Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, szData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExByte Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, szData As Byte, ByRef lpcbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpData As Long, ByVal lpcbData As Long) As Long
Private Declare Function RegEnumValueLong Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegEnumValueStr Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegEnumValueByte Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
' Other declares:
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
' Attributes:
Private m_hClassKey As Long
Private m_sSectionKey As String
Private m_sValueKey As String
Private m_vValue As Variant
Private m_sSetValue As String
Private m_vDefault As Variant
Private m_eValueType As RegValueType

'========================================================================='

Friend Sub Init(ByVal PreviousInChain As Registry, ByVal Root As RegRoot, ByVal Path As String)
    Set mPreviousInChain = PreviousInChain
    Let mRoot = Root
    Let mPath = Path
End Sub

Public Function OpenRegistry(ByVal Root As RegRoot, ByVal Path As String, Optional ByVal Chained As Boolean = False) As Registry
    Dim result As New Registry
    If Chained Then
        result.Init Me, Root, Path
    Else
        result.Init Nothing, Root, Path
    End If
    Set OpenRegistry = result
End Function

Public Function OpenKey(ByVal Path As String, Optional ByVal Chained As Boolean = False) As Registry
    Set OpenKey = OpenRegistry(mRoot, Path, Chained)
End Function

Public Function OpenSoftware(ByVal applic As App, Optional ByVal CurrentUser As Boolean = True, Optional ByVal Chained As Boolean = False) As Registry
    Dim Root As RegRoot
    If CurrentUser Then
        Root = HKEY_CURRENT_USER
    Else
        Root = HKEY_LOCAL_MACHINE
    End If
    Set OpenSoftware = OpenRegistry(Root, "/SOFTWARE/" & applic.CompanyName & "/" & applic.ProductName, Chained)
End Function

Public Property Get PreviousInChain() As Registry
    Set PreviousInChain = mPreviousInChain
End Property

Public Property Get Root() As RegRoot
    Root = mRoot
End Property

Public Property Get Path() As String
    Path = mPath
End Property

Public Function Values(Optional ByVal includeDefault As Boolean = True) As Collection
End Function

Public Function SubKeys() As Collection
End Function

Public Function GetDefaultValue(Optional ByVal valDefault As Variant = Null) As Variant
End Function

Public Sub SetDefaultValue(ByVal valValue As Variant)
End Sub

Public Function ValueExists(ByVal valName As String) As Boolean
End Function

Public Sub NewValue(ByVal valName As String, ByVal valValue As Variant, Optional ByVal valType As RegValueType = REG_SZ)
End Sub

Public Function GetValue(ByVal valName As String, Optional ByVal valDefault As Variant = Null) As Variant
    Dim datatype As Long
    Dim datasize As Long
    Dim ldata As Long, sdata As String, bdata() As Byte
    If RegOpenKeyEx(mRoot, mPath, 0, KEY_READ, tempKey) = ERROR_SUCCESS Then
        If RegQueryValueExLong(tempKey, valName, 0&, datatype, 0&, datasize) = ERROR_SUCCESS Then
            Select Case datatype
            Case reg_dword, reg_dword_little_endian
                RegQueryValueExLong tempKey, valName, 0&, datatype, ldata, datasize
                GetValue = CLng(ldata)
            Case reg_dword_big_endian
                RegQueryValueExLong tempKey, valName, 0&, datatype, ldata, datasize
                GetValue = mSwapIndian(ldata)
            Case REG_SZ
                sdata = String(datasize - 1, 0)
                RegQueryValueExStr tempKey, valName, 0&, datatype, sdata, datasize
                GetValue = sdata
            Case REG_MULTI_SZ
                sdata = String(datasize - 1, 0)
                RegQueryValueExStr tempKey, valName, 0&, datatype, sdata, datasize
                GetValue = MultiToColl(sdata)
            Case REG_EXPAND_SZ
                sdata = String(datasize - 1, 0)
                RegQueryValueExStr tempKey, valName, 0&, datatype, sdata, datasize
                GetValue = sdata
            Case Else '(REG_BINARY, ...)'
                ReDim bdata(datasize)
                RegQueryValueExByte tempKey, valName, 0&, datatype, bdata(0), datasize
                GetValue = bdata
            End Select
        Else
            RegCloseKey tempKey
            RaiseError
        End If
        RegCloseKey tempKey
    Else
        RaiseError
    End If
End Function

Public Sub SetValue(ByVal valName As String, ByVal valValue As Variant)
End Sub

Public Sub DeleteValue(ByVal valName As String)
End Sub

Public Function Parent() As Registry
    Set Parent = OpenKey(ParentKey(mPath))
End Function

Public Function Exists() As Boolean
    'Whether the current key exists...
    If RegOpenKeyEx(mRoot, mPath, 0, KEY_READ, tempKey) = ERROR_SUCCESS Then
        Exists = True
        RegCloseKey tempKey
    Else
        Exists = False
    End If
End Function

Public Function NewKey(ByVal Key As String) As Registry
    Dim tsa As SECURITY_ATTRIBUTES
    Dim result As Long
    Dim disp As Long
    If RegOpenKeyEx(mRoot, mPath, 0, KEY_READ, tempKey) = ERROR_SUCCESS Then
        If RegCreateKeyEx(tempKey, Key, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, tsa, result, disp) = ERROR_SUCCESS Then
            RegCloseKey result
            Set NewKey = SubKey(Key)
        Else
            RegCloseKey tempKey
            RaiseError , "Registry create failure."
        End If
        RegCloseKey tempKey
    Else
        RaiseError
    End If
End Function

Public Function SubKey(ByVal Key As String) As Registry
    Set SubKey = OpenKey(mPath & "/" & Key)
End Function

Public Function SubKeyExists(ByVal Key As String) As Boolean
    SubKeyExists = SubKey(Key).Exists
End Function

Public Sub DeleteSubKey(ByVal Key As String)
    If RegOpenKeyEx(mRoot, mPath, 0, KEY_READ, tempKey) = ERROR_SUCCESS Then
        If RegDeleteKey(tempKey, Key) <> ERROR_SUCCESS Then
            RegCloseKey tempKey
            RaiseError , "Registry delete failed."
        End If
        RegCloseKey tempKey
    Else
        RaiseError
    End If
    
End Sub

Public Sub Delete()
    'Deletes the current key
    Parent.DeleteSubKey LastKey(mPath)
End Sub

'========================================================================'

Private Function LastKey(ByVal Path As String)
    Dim i As Long
    i = Len(Path) - InStr(StrReverse(Path), "/") + 2
    If i > Len(Path) Then
        LastKey = Path
    Else
        LastKey = Mid(Path, i)
    End If
End Function

Private Function ParentKey(ByVal Path) As String
    Dim i As Long
    i = Len(Path) - InStr(StrReverse(Path), "/") + 2
    If i > Len(Path) Then
        LastKey = ""
    Else
        LastKey = Left(Path, i)
    End If
End Function

Private Sub RaiseError(Optional ByVal errno As Long = 5, Optional ByVal errmsg As String = "Registry access error.")
    Err.Raise errno, "Registry", errmsg
End Sub

Private Function MultiToCol(ByVal ms As String) As Collection
    Dim result As New Collection
    Dim i As Long
    
    i = InStr(1, ms, Chr(0))
    While i > 1
        result.Add Left(ms, i - 1)
        ms = Mid(ls, i + 1)
        i = InStr(1, ms, Chr(0))
    Wend
    Set MultiToCol = result
    
End Function

'========================================================================'

Private Property Get mKeyExists() As Boolean
    'KeyExists = bCheckKeyExists( '                m_hClassKey, '                m_sSectionKey '            )
Dim hKey As Long
    If RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, 1, hKey) = ERROR_SUCCESS Then
        mKeyExists = True
        RegCloseKey hKey
    Else
        mKeyExists = False
    End If
    
End Property

Private Function mCreateKey() As Boolean
Dim tsa As SECURITY_ATTRIBUTES
Dim hKey As Long
Dim lCreate As Long
Dim e As Long

    'Open or Create the key
    e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, tsa, hKey, lCreate)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to create registry Key: '" & m_sSectionKey
    Else
        mCreateKey = (e = ERROR_SUCCESS)
        'Close the key
        RegCloseKey hKey
    End If
End Function

Private Function mDeleteKey() As Boolean
Dim e As Long
    e = RegDeleteKey(m_hClassKey, m_sSectionKey)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey
    Else
        mDeleteKey = (e = ERROR_SUCCESS)
    End If
    
End Function

Private Function mDeleteValue() As Boolean
Dim e As Long
Dim hKey As Long

    e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ALL_ACCESS, hKey)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to open key '" & m_hClassKey & "',Section: '" & m_sSectionKey & "' for delete access"
    Else
        e = RegDeleteValue(hKey, m_sValueKey)
        If e Then
            Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey
        Else
            mDeleteValue = (e = ERROR_SUCCESS)
        End If
    End If

End Function
Private Property Get mValue() As Variant
Dim vValue As Variant
Dim cData As Long, sdata As String, ordType As Long, e As Long
Dim hKey As Long

    e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, hKey)
    'ApiRaiseIf e

    e = RegQueryValueExLong(hKey, m_sValueKey, 0&, ordType, 0&, cData)
    If e And e <> ERROR_MORE_DATA Then
        mValue = m_vDefault
        Exit Property
    End If
    
    m_eValueType = ordType
    Select Case ordType
    Case reg_dword, reg_dword_little_endian
        Dim iData As Long
        e = RegQueryValueExLong(hKey, m_sValueKey, 0&, _
                               ordType, iData, cData)
        vValue = CLng(iData)
        
    Case reg_dword_big_endian  ' Unlikely, but you never know
        Dim dwData As Long
        e = RegQueryValueExLong(hKey, m_sValueKey, 0&, _
                               ordType, dwData, cData)
        vValue = mSwapEndian(dwData)
        
    Case REG_SZ, REG_MULTI_SZ ' Same thing to Visual Basic
        sdata = String$(cData - 1, 0)
        e = RegQueryValueExStr(hKey, m_sValueKey, 0&, _
                               ordType, sdata, cData)
        vValue = sdata
        
    Case REG_EXPAND_SZ
        sdata = String$(cData - 1, 0)
        e = RegQueryValueExStr(hKey, m_sValueKey, 0&, _
                               ordType, sdata, cData)
        vValue = mExpandEnvStr(sdata)
        
    ' Catch REG_BINARY and anything else
    Case Else
        Dim abData() As Byte
        ReDim abData(cData)
        e = RegQueryValueExByte(hKey, m_sValueKey, 0&, _
                                ordType, abData(0), cData)
        vValue = abData
        
    End Select
    mValue = vValue
    
End Property
Private Property Let mValue( _
        ByVal vValue As Variant _
    )
Dim ordType As Long
Dim c As Long
Dim hKey As Long
Dim e As Long
Dim lCreate As Long
Dim tsa As SECURITY_ATTRIBUTES

    'Open or Create the key
    e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, _
                 KEY_ALL_ACCESS, tsa, hKey, lCreate)
    
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
    Else

        Select Case m_eValueType
        Case REG_BINARY
            If (VarType(vValue) = vbArray + vbByte) Then
                Dim ab() As Byte
                ab = vValue
                ordType = REG_BINARY
                c = UBound(ab) - LBound(ab) - 1
                e = RegSetValueExByte(hKey, m_sValueKey, 0&, ordType, ab(0), c)
            Else
                Err.Raise 26001
            End If
        Case reg_dword, reg_dword_big_endian, reg_dword_little_endian
            If (VarType(vValue) = vbInteger) Or (VarType(vValue) = vbLong) Then
                Dim i As Long
                i = vValue
                ordType = reg_dword
                e = RegSetValueExLong(hKey, m_sValueKey, 0&, ordType, i, 4)
            End If
        Case REG_SZ, REG_EXPAND_SZ
            Dim s As String, iPos As Long
            s = vValue
            ordType = REG_SZ
            ' Assume anything with two non-adjacent percents is expanded string
            iPos = InStr(s, "%")
            If iPos Then
                If InStr(iPos + 2, s, "%") Then ordType = REG_EXPAND_SZ
            End If
            c = Len(s) + 1
            e = RegSetValueExStr(hKey, m_sValueKey, 0&, ordType, s, c)
            
        ' User should convert to a compatible type before calling
        Case Else
            e = ERROR_INVALID_DATA
            
        End Select
        
        If Not e Then
            m_vValue = vValue
        Else
            Err.Raise vbObjectError + 1048 + 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
        End If
        
        'Close the key
        RegCloseKey hKey
    
    End If
    
End Property
Private Function mEnumerateValues( _
        ByRef sKeyNames() As String, _
        ByRef iKeyCount As Long _
    ) As Boolean
Dim lResult As Long
Dim hKey As Long
Dim sName As String
Dim lNameSize As Long
Dim sdata As String
Dim lIndex As Long
Dim cJunk As Long
Dim cNameMax As Long
Dim ft As Currency
   
   ' Log "EnterEnumerateValues"

   iKeyCount = 0
   Erase sKeyNames()
    
   lIndex = 0
   lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, hKey)
   If (lResult = ERROR_SUCCESS) Then
      ' Log "OpenedKey:" & m_hClassKey & "," & m_sSectionKey
      lResult = RegQueryInfoKey(hKey, "", cJunk, 0, _
                               cJunk, cJunk, cJunk, cJunk, _
                               cNameMax, cJunk, cJunk, ft)
       Do While lResult = ERROR_SUCCESS
   
           'Set buffer space
           lNameSize = cNameMax + 1
           sName = String$(lNameSize, 0)
           If (lNameSize = 0) Then lNameSize = 1
           
           ' Log "Requesting Next Value"
         
           'Get value name:
           lResult = RegEnumValue(hKey, lIndex, sName, lNameSize, _
                                  0&, 0&, 0&, 0&)
           ' Log "RegEnumValue returned:" & lResult
           If (lResult = ERROR_SUCCESS) Then
       
                ' Although in theory you can also retrieve the actual
                ' value and type here, I found it always (ultimately) resulted in
                ' a GPF, on Win95 and NT.  Why?  Can anyone help?
       
               sName = Left$(sName, lNameSize)
               ' Log "Enumerated value:" & sName
                 
               iKeyCount = iKeyCount + 1
               ReDim Preserve sKeyNames(1 To iKeyCount) As String
               sKeyNames(iKeyCount) = sName
           End If
           lIndex = lIndex + 1
       Loop
   End If
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If

   ' Log "Exit Enumerate Values"
   mEnumerateValues = True
   Exit Function
   
EnumerateValuesError:
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
   Err.Raise vbObjectError + 1048 + 26003, App.EXEName & ".cRegistry", Err.Description
   Exit Function

End Function
Private Function mEnumerateSections( _
        ByRef sSect() As String, _
        ByRef iSectCount As Long _
    ) As Boolean
Dim lResult As Long
Dim hKey As Long
Dim dwReserved As Long
Dim szBuffer As String
Dim lBuffSize As Long
Dim lIndex As Long
Dim lType As Long
Dim sCompKey As String
Dim iPos As Long

On Error GoTo EnumerateSectionsError

   iSectCount = 0
   Erase sSect
'
   lIndex = 0

   lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
   Do While lResult = ERROR_SUCCESS
       'Set buffer space
       szBuffer = String$(255, 0)
       lBuffSize = Len(szBuffer)
      
      'Get next value
       lResult = RegEnumKey(hKey, lIndex, szBuffer, lBuffSize)
                             
       If (lResult = ERROR_SUCCESS) Then
           iSectCount = iSectCount + 1
           ReDim Preserve sSect(1 To iSectCount) As String
           iPos = InStr(szBuffer, Chr$(0))
           If (iPos > 0) Then
              sSect(iSectCount) = Left(szBuffer, iPos - 1)
           Else
              sSect(iSectCount) = Left(szBuffer, lBuffSize)
           End If
       End If
       
       lIndex = lIndex + 1
   Loop
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
   mEnumerateSections = True
   Exit Function

EnumerateSectionsError:
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
   Err.Raise vbObjectError + 1048 + 26002, App.EXEName & ".cRegistry", Err.Description
   Exit Function
End Function

Private Sub pSetClassValue(ByVal sValue As String)
Dim sSection As String
   mClassKey = HKEY_CLASSES_ROOT
   mValue = sValue
   sSection = mSectionKey
   mClassKey = HKEY_LOCAL_MACHINE
   mSectionKey = "SOFTWARE\Classes\" & sSection
   mValue = sValue
   mSectionKey = sSection
End Sub

Public Sub CreateEXEAssociation( _
        ByVal sExePath As String, _
        ByVal sClassName As String, _
        ByVal sClassDescription As String, _
        ByVal sAssociation As String, _
        Optional ByVal sOpenMenuText As String = "&Open", _
        Optional ByVal bSupportPrint As Boolean = False, _
        Optional ByVal sPrintMenuText As String = "&Print", _
        Optional ByVal bSupportNew As Boolean = False, _
        Optional ByVal sNewMenuText As String = "&New", _
        Optional ByVal bSupportInstall As Boolean = False, _
        Optional ByVal sInstallMenuText As String = "", _
        Optional ByVal lDefaultIconIndex As Long = -1 _
    )
   ' Check if path is wrapped in quotes:
   sExePath = Trim$(sExePath)
   If (Left$(sExePath, 1) <> """") Then
      sExePath = """" & sExePath
   End If
   If (Right$(sExePath, 1) <> """") Then
      sExePath = sExePath & """"
   End If
    
    ' Create the .File to Class association:
   mSectionKey = "." & sAssociation
   mValueType = REG_SZ
   mValueKey = ""
   pSetClassValue sClassName
   
   ' Create the Class shell open command:
   mSectionKey = sClassName
   pSetClassValue sClassDescription
   
   mSectionKey = sClassName & "\shell\open"
   If (sOpenMenuText = "") Then sOpenMenuText = "&Open"
   mValueKey = ""
   pSetClassValue sOpenMenuText
   mSectionKey = sClassName & "\shell\open\command"
   mValueKey = ""
   pSetClassValue sExePath & " ""%1"""
   
   If (bSupportPrint) Then
      mSectionKey = sClassName & "\shell\print"
      If (sPrintMenuText = "") Then sPrintMenuText = "&Print"
      mValueKey = ""
      pSetClassValue sPrintMenuText
      mSectionKey = sClassName & "\shell\print\command"
      mValueKey = ""
      pSetClassValue sExePath & " /p ""%1"""
   End If
   
   If (bSupportInstall) Then
      If (sInstallMenuText = "") Then
         sInstallMenuText = "&Install " & sAssociation
      End If
      mSectionKey = sClassName & "\shell\add"
      mValueKey = ""
      pSetClassValue sInstallMenuText
      mSectionKey = sClassName & "\shell\add\command"
      mValueKey = ""
      pSetClassValue sExePath & " /a ""%1"""
   End If
   
   If (bSupportNew) Then
      mSectionKey = sClassName & "\shell\new"
      mValueKey = ""
      If (sNewMenuText = "") Then sNewMenuText = "&New"
      pSetClassValue sNewMenuText
      mSectionKey = sClassName & "\shell\new\command"
      mValueKey = ""
      pSetClassValue sExePath & " /n ""%1"""
   End If
   
   If lDefaultIconIndex > -1 Then
      mSectionKey = sClassName & "\DefaultIcon"
      mValueKey = ""
      pSetClassValue sExePath & "," & CStr(lDefaultIconIndex)
   End If
    
End Sub

Public Sub CreateAdditionalEXEAssociations( _
      ByVal sClassName As String, _
      ParamArray vItems() As Variant _
   )
Dim iItems As Long
Dim iItem As Long
   
   On Error Resume Next
   iItems = UBound(vItems) + 1
   If (iItems Mod 3) <> 0 Or (Err.Number <> 0) Then
      Err.Raise vbObjectError + 1048 + 26004, App.EXEName & ".cRegistry", "Invalid parameter list passed to CreateAdditionalEXEAssociations - expected Name/Text/Command"
   Else
      ' Check if it exists:
      mSectionKey = sClassName
      If Not (mKeyExists) Then
         Err.Raise vbObjectError + 1048 + 26005, App.EXEName & ".cRegistry", "Error - attempt to create additional associations before class defined."
      Else
         For iItem = 0 To iItems - 1 Step 3
            mValueType = REG_SZ
            mSectionKey = sClassName & "\shell\" & vItems(iItem)
            mValueKey = ""
            pSetClassValue vItems(iItem + 1)
            mSectionKey = sClassName & "\shell\" & vItems(iItem) & "\command"
            mValueKey = ""
            pSetClassValue vItems(iItem + 2)
         Next iItem
      End If
   End If
   
End Sub

Private Property Get mValueType() As RegValueType
    mValueType = m_eValueType
End Property

Private Property Let mValueType(ByVal eValueType As RegValueType)
    m_eValueType = eValueType
End Property

Private Property Get mClassKey() As RegRoot
    mClassKey = m_hClassKey
End Property

Private Property Let mClassKey( _
        ByVal eKey As RegRoot _
    )
    m_hClassKey = eKey
End Property
Private Property Get mSectionKey() As String
    mSectionKey = m_sSectionKey
End Property
Private Property Let mSectionKey( _
        ByVal sSectionKey As String _
    )
    m_sSectionKey = sSectionKey
End Property
Private Property Get mValueKey() As String
    mValueKey = m_sValueKey
End Property
Private Property Let mValueKey( _
        ByVal sValueKey As String _
    )
    m_sValueKey = sValueKey
End Property
Private Property Get mDefault() As Variant
    mDefault = m_vDefault
End Property
Private Property Let mDefault( _
        ByVal vDefault As Variant _
    )
    m_vDefault = vDefault
End Property
Private Function mSwapEndian(ByVal dw As Long) As Long
    CopyMemory ByVal VarPtr(mSwapEndian) + 3, dw, 1
    CopyMemory ByVal VarPtr(mSwapEndian) + 2, ByVal VarPtr(dw) + 1, 1
    CopyMemory ByVal VarPtr(mSwapEndian) + 1, ByVal VarPtr(dw) + 2, 1
    CopyMemory mSwapEndian, ByVal VarPtr(dw) + 3, 1
End Function
Private Function mExpandEnvStr(sdata As String) As String
    Dim c As Long, s As String
    ' Get the length
    s = "" ' Needed to get around Windows 95 limitation
    c = ExpandEnvironmentStrings(sdata, s, c)
    ' Expand the string
    s = String$(c - 1, 0)
    c = ExpandEnvironmentStrings(sdata, s, c)
    mExpandEnvStr = s
End Function

