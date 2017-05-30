Attribute VB_Name = "VB4032_Tools"
'TOOLS
DefInt A-Z

Type RECT
       Left        As Integer
       Top         As Integer
       right       As Integer
       botom       As Integer
End Type


Type POINTAPI
       x           As Integer
       y           As Integer
End Type


Type WINDOWPLACEMENT
       Length      As Integer
       FLAGS       As Integer
       showCmd     As Integer
       ptMinPosition As POINTAPI
       ptMaxPosition As POINTAPI
       rcNormalPosition As RECT
End Type

Declare Function sndPlaySound Lib "MMSYSTEM.DLL" (ByVal lpszSoundName$, ByVal wFlags%) As Integer
Declare Sub MessageBeep Lib "user32" (ByVal wType As Integer)

Declare Function GetWindowsDirectory Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

Declare Function GetProfileInt% Lib "kernel32" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%)
Declare Function GetProfileString% Lib "kernel32" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%)
Declare Function WriteProfileString% Lib "kernel32" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$)

Declare Function GetPrivateProfileInt% Lib "kernel32" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFileName$)
Declare Function GetPrivateProfileString% Lib "kernel32" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$)
Declare Function WritePrivateProfileString% Lib "kernel32" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$, ByVal lpFileName$)

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Declare Function GetModuleUsage% Lib "kernel32" (ByVal hModule%)

Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long


Declare Function GetWindowText Lib "user32" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
Declare Function FindWindow Lib "user32" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function GetNextWindow Lib "user32" (ByVal hwnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetWindow Lib "user32" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd%, lpwndpl As WINDOWPLACEMENT) As Integer

Declare Function GetVersion Lib "kernel32" () As Long
Declare Function GetFreeSpace Lib "kernel32" (ByVal wFlags As Integer) As Long
Declare Function GetWinFlags Lib "kernel32" () As Long
Declare Function GetModuleHandle Lib "kernel32" (ByVal lpModuleName As String) As Integer
Declare Function GetFreeSystemResources Lib "user32" (ByVal fuSysResource As Integer) As Integer
Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal wNewWord As Integer) As Integer

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Global Const GFSR_SYSTEMRESOURCES = &H0
Global Const GFSR_GDIRESOURCES = &H1
Global Const GFSR_USERRESOURCES = &H2

Global Const WF_CPU086 = &H40
Global Const WF_CPU186 = &H80
Global Const WF_CPU286 = &H2
Global Const WF_CPU386 = &H4
Global Const WF_CPU486 = &H8
Global Const WF_80x87 = &H400
Global Const WF_STANDARD = &H10
Global Const WF_ENHANCED = &H20

Global Const GW_HWNDFIRST = 0
Global Const GW_HWNDLAST = 1
Global Const GW_HWNDNEXT = 2
Global Const GW_HWNDPREV = 3
Global Const GW_OWNER = 4
Global Const GW_CHILD = 5

Global Const WM_USER = &H400
Global Const EM_LIMITTEXT = WM_USER + 21

Const LB_SETHORIZONTALEXTENT = &H400 + 21

Private Wait_Level As Long
Public Function AnsiToHTML$(a$)
Dim c As String, char As Integer
r$ = ""
For n = 1 To Len(a$)
  char = Asc(Mid$(a$, n, 1))
  c = Mid$(a$, n, 1)
  Select Case char    'Translates ASCII to ANSI.
    Case 34: c = "&quot;"
    Case 38: c = "&amp;"
    Case 60: c = "&lt;"
    Case 62: c = "&gt;"
    Case 224: c = "&agrave;"
    Case 232: c = "&egrave;"
    Case 233: c = "&eacute;"
    Case 235: c = "&euml;"
    Case 169: c = "&copy;"
    Case 174: c = "&reg;"
    Case 231: c = "&ccedil;"
    Case 223: c = "&beta;"
    Case 215: c = "x"
  End Select
  r$ = r$ + c
Next
AnsiToHTML$ = r$
End Function

Public Function AnsiToAscii$(a$)
Dim c As Integer, char As Integer
r$ = ""
For n = 1 To Len(a$)
  char = Asc(Mid$(a$, n, 1))
  Select Case char    'Translates ASCII to ANSI.
    Case 9: c = 9
    Case 13: c = 13
    Case 32 To 126: c = char
    Case 130: c = 44
    Case 131: c = 159
    Case 136: c = 94
    Case 139: c = 60
    Case 145: c = 96
    Case 146: c = 39
    Case 147: c = 96
    Case 148: c = 39
    Case 149: c = 249
    Case 150: c = 45
    Case 151: c = 196
    Case 152: c = 126
    Case 155: c = 62
    Case 159: c = 89
    Case 160: c = 32
    Case 161: c = 173
    Case 162: c = 155
    Case 163: c = 156
    Case 164: c = 15
    Case 165: c = 157
    Case 166: c = 124
    Case 167: c = 21
    Case 170: c = 166
    Case 171: c = 174
    Case 172: c = 170
    Case 173: c = 45
    Case 176: c = 248
    Case 177: c = 241
    Case 178: c = 253
    Case 180: c = 39
    Case 181: c = 230
    Case 182: c = 227
    Case 183: c = 250
    Case 186: c = 167
    Case 187: c = 175
    Case 188: c = 172
    Case 189: c = 171
    Case 191: c = 168
    Case 192: c = 65
    Case 193: c = 65
    Case 194: c = 65
    Case 195: c = 65
    Case 196: c = 142
    Case 197: c = 143
    Case 198: c = 146
    Case 199: c = 128
    Case 200: c = 69
    Case 201: c = 144
    Case 202: c = 69
    Case 203: c = 69
    Case 204: c = 73
    Case 205: c = 73
    Case 206: c = 73
    Case 207: c = 73
    Case 209: c = 165
    Case 210: c = 79
    Case 211: c = 79
    Case 212: c = 79
    Case 213: c = 79
    Case 214: c = 153
    Case 215: c = 42
    Case 216: c = 48
    Case 217: c = 85
    Case 218: c = 85
    Case 219: c = 85
    Case 220: c = 154
    Case 221: c = 89
    Case 223: c = 225
    Case 224: c = 133
    Case 225: c = 160
    Case 226: c = 131
    Case 227: c = 97
    Case 228: c = 132
    Case 229: c = 134
    Case 230: c = 145
    Case 231: c = 135
    Case 232: c = 138
    Case 233: c = 130
    Case 234: c = 136
    Case 235: c = 137
    Case 236: c = 141
    Case 237: c = 161
    Case 238: c = 140
    Case 239: c = 139
    Case 241: c = 164
    Case 242: c = 149
    Case 243: c = 162
    Case 244: c = 147
    Case 245: c = 111
    Case 246: c = 148
    Case 247: c = 246
    Case 248: c = 237
    Case 249: c = 151
    Case 250: c = 163
    Case 251: c = 150
    Case 252: c = 129
    Case 253: c = 121
    Case 255: c = 152
    Case Else: c = -1    'Character can't be translated
  End Select
  If c <> -1 Then r$ = r$ + Chr$(c)
Next
AnsiToAscii$ = r$
End Function

Public Function AsciiToAnsi$(a$)
Dim c As Integer, char As Integer
r$ = ""
For n = 1 To Len(a$)
  char = Asc(Mid$(a$, n, 1))
  Select Case char    'Translates ASCII to ANSI.
    Case 9: c = 9
    Case 13: c = 13
    Case 15: c = 164
    Case 20: c = 182
    Case 21: c = 167
    Case 32 To 126: c = char
    Case 128: c = 199
    Case 129: c = 252
    Case 130: c = 233
    Case 131: c = 226
    Case 132: c = 228
    Case 133: c = 224
    Case 134: c = 229
    Case 135: c = 231
    Case 136: c = 234
    Case 137: c = 235
    Case 138: c = 232
    Case 139: c = 239
    Case 140: c = 238
    Case 141: c = 236
    Case 142: c = 196
    Case 143: c = 197
    Case 144: c = 201
    Case 145: c = 230
    Case 146: c = 198
    Case 147: c = 244
    Case 148: c = 246
    Case 149: c = 242
    Case 150: c = 251
    Case 151: c = 249
    Case 152: c = 255
    Case 153: c = 214
    Case 154: c = 220
    Case 155: c = 162
    Case 156: c = 163
    Case 157: c = 165
    Case 159: c = 131
    Case 160: c = 225
    Case 161: c = 237
    Case 162: c = 243
    Case 163: c = 250
    Case 164: c = 241
    Case 165: c = 209
    Case 166: c = 170
    Case 167: c = 186
    Case 168: c = 191
    Case 170: c = 172
    Case 171: c = 189
    Case 172: c = 188
    Case 173: c = 161
    Case 174: c = 171
    Case 175: c = 187
    Case 176: c = 127
    Case 177: c = 127
    Case 178: c = 127
    Case 179: c = 124
    Case 180: c = 43
    Case 181: c = 43
    Case 182: c = 43
    Case 183: c = 43
    Case 184: c = 43
    Case 185: c = 43
    Case 186: c = 124
    Case 187: c = 43
    Case 188: c = 43
    Case 189: c = 43
    Case 190: c = 43
    Case 191: c = 43
    Case 192: c = 43
    Case 193: c = 43
    Case 194: c = 43
    Case 195: c = 43
    Case 196: c = 151
    Case 197: c = 43
    Case 198: c = 43
    Case 199: c = 43
    Case 200: c = 43
    Case 201: c = 43
    Case 202: c = 43
    Case 203: c = 43
    Case 204: c = 43
    Case 205: c = 61
    Case 206: c = 43
    Case 207: c = 43
    Case 208: c = 43
    Case 209: c = 43
    Case 210: c = 43
    Case 211: c = 43
    Case 212: c = 43
    Case 213: c = 43
    Case 214: c = 43
    Case 215: c = 43
    Case 216: c = 43
    Case 217: c = 43
    Case 218: c = 43
    Case 219: c = 127
    Case 220: c = 127
    Case 221: c = 127
    Case 222: c = 127
    Case 223: c = 127
    Case 225: c = 223
    Case 227: c = 182
    Case 230: c = 181
    Case 237: c = 248
    Case 241: c = 177
    Case 246: c = 247
    Case 248: c = 176
    Case 249: c = 149
    Case 250: c = 183
    Case 253: c = 178
    Case 254: c = 127
    Case 255: c = 32
    Case Else: c = -1    'Character can't be translated
  End Select
  If c <> -1 Then r$ = r$ + Chr$(c)
Next
AsciiToAnsi$ = r$
End Function

Sub AddHorScrollbarToList(l As ListBox, w%)
  If w% > (l.Width \ 15) Then x& = SendMessage&(l.hwnd, LB_SETHORIZONTALEXTENT, w%, (0&))
End Sub

Sub AllowOnlyKeys(k$, KeyAscii As Integer)
  'This subroutine is to be placed in the KeyPress event
  'of Textboxes for instance, and allows only the characters
  'put in k$ to be typed.
  'Special keys such as Delete, Backspace,... are always
  'allowed.

  If KeyAscii >= 32 Then
     If InStr(k$, Chr$(KeyAscii)) = 0 Then
        KeyAscii = 0
     End If
  End If

End Sub


Public Function BuildCollection(Optional item1 As Variant, Optional item2 As Variant, Optional item3 As Variant, Optional item4 As Variant, Optional item5 As Variant, Optional item6 As Variant, Optional item7 As Variant, Optional item8 As Variant, Optional item9 As Variant, Optional item10 As Variant) As Collection
  Dim answer As New Collection
  
  If Not IsMissing(item1) Then answer.Add item1
  If Not IsMissing(item2) Then answer.Add item2
  If Not IsMissing(item3) Then answer.Add item3
  If Not IsMissing(item4) Then answer.Add item4
  If Not IsMissing(item5) Then answer.Add item5
  If Not IsMissing(item6) Then answer.Add item6
  If Not IsMissing(item7) Then answer.Add item7
  If Not IsMissing(item8) Then answer.Add item8
  If Not IsMissing(item9) Then answer.Add item9
  If Not IsMissing(item10) Then answer.Add item10
  Set BuildCollection = answer

End Function
Function Decode(Value As Variant, a As Variant, Optional b As Variant, Optional c As Variant, Optional d As Variant, Optional e As Variant, Optional f As Variant, Optional g As Variant, Optional h As Variant, Optional i As Variant, Optional j As Variant) As Variant
  Dim dict As Collection
  Dim Default As Variant

  If IsMissing(b) Then Default = a Else dict.Add BuildCollection(a, b)
  If IsMissing(d) Then Default = IIf(IsMissing(c), Null, c): GoTo Decode Else dict.Add BuildCollection(c, d): GoTo Decode
  If IsMissing(f) Then Default = IIf(IsMissing(e), Null, e): GoTo Decode Else dict.Add BuildCollection(e, f): GoTo Decode
  If IsMissing(h) Then Default = IIf(IsMissing(g), Null, g): GoTo Decode Else dict.Add BuildCollection(g, h): GoTo Decode
  If IsMissing(j) Then Default = IIf(IsMissing(i), Null, i): GoTo Decode Else dict.Add BuildCollection(i, j): GoTo Decode

Decode:
  Dim v As Collection
  For Each v In dict
    If v(1) = Value Then Decode = v(2)
  Next
  Decode = Default

End Function

Public Function EmptyCollection() As Collection
  Dim answer As New Collection
  Set EmptyCollection = answer
End Function

Public Function IncreaseCollection(aCollection As Collection, Optional item1 As Variant, Optional item2 As Variant, Optional item3 As Variant, Optional item4 As Variant, Optional item5 As Variant, Optional item6 As Variant, Optional item7 As Variant, Optional item8 As Variant, Optional item9 As Variant, Optional item10 As Variant) As Collection
  Dim answer As New Collection
  Dim item As Variant
  
  For Each item In aCollection
    answer.Add item
  Next
  If Not IsMissing(item1) Then answer.Add item1
  If Not IsMissing(item2) Then answer.Add item2
  If Not IsMissing(item3) Then answer.Add item3
  If Not IsMissing(item4) Then answer.Add item4
  If Not IsMissing(item5) Then answer.Add item5
  If Not IsMissing(item6) Then answer.Add item6
  If Not IsMissing(item7) Then answer.Add item7
  If Not IsMissing(item8) Then answer.Add item8
  If Not IsMissing(item9) Then answer.Add item9
  If Not IsMissing(item10) Then answer.Add item10
  Set IncreaseCollection = answer

End Function
Sub CenterForm(Whichform As Form)
  Whichform.Top = (Screen.Height - Whichform.Height) \ 2
  Whichform.Left = (Screen.Width - Whichform.Width) \ 2
End Sub

Sub Dec(x As Variant, Optional n As Variant)
  If IsMissing(n) Then
     x = x - 1
  Else
     x = x - n
  End If
End Sub

Sub Encrypt(secret$, Password$)
  ' secret$ = the string you wish to encrypt or decrypt.
  ' PassWord$ = the password with which to encrypt the string.
  l = Len(Password$)
  For x = 1 To Len(secret$)
    char = Asc(Mid$(Password$, (x Mod l) - l * ((x Mod l) = 0), 1))
    Mid$(secret$, x, 1) = Chr$(Asc(Mid$(secret$, x, 1)) Xor char)
  Next
End Sub

Function Exist(Path$) As Integer
  x% = FreeFile
  On Error Resume Next
  Open Path$ For Input Access Read As x%
  If Err = 0 Then
     Exist = True
  Else
     Exist = False
  End If
  Close x%
End Function

Function GetFreeMemory() As Long
  GetFreeMemory = GetFreeSpace(0)
End Function

Function GetFreeResources() As Integer
  GetFreeResources = GetFreeSystemResources(GFSR_SYSTEMRESOURCES)
End Function

Function GetIntFromMyINI(File As String, Section As String, Keyword As String, Default As Integer) As Integer
  GetIntFromMyINI = GetPrivateProfileInt(Section, Keyword, Default, File)
End Function

Function GetIntFromWININI(Section As String, Keyword As String, Default As Integer) As Integer
  GetIntFromWININI = GetProfileInt(Section, Keyword, Default)
End Function

Function GetStrFromMyINI(File As String, Section As String, Keyword As String, Default As String) As String
  'For compatibility purposes only
  GetStrFromMyINI = GetSetting(File, Section, Keyword, Default)
End Function

Function GetStrFromWININI(Section As String, Keyword As String, Default As String) As String
  Dim t As String * 128
  Dim valid As Integer
  valid = GetProfileString(Section, Keyword, Default, t, Len(t))
  a$ = RTrim$(LTrim$(Left$(t, valid)))
  If Left$(a$ + "*", 1) = Chr$(34) Then
     p% = InStr(2, a$, Chr$(34))
     If p% <> 0 Then
        a$ = Mid$(a$, 2, p% - 2)
     Else
        p% = InStr(a$, ";")
        If p% <> 0 Then
           a$ = RTrim$(Left$(a$, p% - 1))
        End If
     End If
  Else
     p% = InStr(a$, ";")
     If p% <> 0 Then
        a$ = RTrim$(Left$(a$, p% - 1))
     End If
  End If
  GetStrFromWININI = a$
End Function

Public Function GetWaitLevel() As Long
  GetWaitLevel = Wait_Level
End Function

Function GetWindowsDir() As String
  Dim s As String * 145
  i% = GetWindowsDirectory(s, 144)
  GetWindowsDir = Left$(s, i%)
End Function

Function GetWindowsMode() As Integer
  'Returns 0 for standard mode and 1 for enhanced mode :
  f& = GetWinFlags&()
  If f& And WF_ENHANCED Then
     GetWindowsMode = 1
  Else
     GetWindowsMode = 0
  End If
End Function

Function GetWinVersion$()
  Dim min%
  On Error Resume Next
  v& = GetVersion()
  maj% = v& And &HFF
  min% = (v& And &HFF00&) \ 256
  GetWinVersion$ = Format$(maj%, "0") + "." + Format$(min%, "0")
End Function

Sub HideControlMenuItem(Frm As Form, item$)
  Select Case RTrim$(LTrim$(UCase$(item$)))
    Case "RESTORE":  itemid& = &HF120&
    Case "MOVE":     itemid& = &HF010&
    Case "SIZE":     itemid& = &HF000&
    Case "MINIMIZE": itemid& = &HF020&
    Case "MAXIMIZE": itemid& = &HF030&
    Case "CLOSE":    itemid& = &HF060&
    Case Else
      MessageBeep 16
      MsgBox "INVALID PARAMETER IN " + Chr$(13) + "TOOLS1.BAS: SUB HideControlMenuItem ( Form , """ + item$ + """ )", 16
      Stop
  End Select
  hMenu& = GetSystemMenu(Frm.hwnd, 0)
  d& = DeleteMenu(hMenu&, itemid&, &H0&)
End Sub

Function IIf(condition As Boolean, iftrue As Variant, iffalse As Variant) As Variant
  If condition Then IIf = iftrue Else IIf = iffalse
End Function

Sub Inc(x As Variant, Optional n As Variant)
  If IsMissing(n) Then
     x = x + 1
  Else
     x = x + n
  End If
End Sub

Public Function InCollection(item As Variant, col As Collection) As Boolean
  Dim i As Variant
  
  InCollection = False
  
  For Each i In col
    If i = item Then InCollection = True: Exit Function
  Next

End Function

Public Function IsNothing(o As Object)
  Dim dummy As String
  On Error Resume Next
  IsNothing = False
  Err.Number = 0
  dummy = CStr(o)
  If Err.Number = 91 Then IsNothing = True
End Function

Sub KeyPressUpcase(KeyAscii As Integer)
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Sub LimitTextLenInControl(ctrl As Control, Size%)
  l& = SendMessage&(ctrl.hwnd, EM_LIMITTEXT, Size%, (0&))
End Sub

Function MakeTextboxReadonly(txtbox As TextBox, Bool%) As Integer
  'This function makes a textbox control readonly or
  'readwrite depending on the value of Bool% (being
  'true for readonly, else false).
  'The function returns true if the operation succeeded.
  If SendMessage(txtbox.hwnd, &H41F, Bool%, 0&) = 0& Then
     MakeTextboxReadonly% = False
  Else
     MakeTextboxReadonly% = True
  End If
End Function

Public Function Max(a As Variant, b As Variant) As Variant
  If a > b Then Max = a Else Max = b
End Function

Public Function min(a As Variant, b As Variant) As Variant
  If a < b Then min = a Else min = b
End Function

Function NVL(Param As Variant, Default As Variant) As Variant
  If IsNull(Param) Then NVL = Default Else NVL = Param
End Function

Function NVLi(Param As Variant, Default As Long) As Long
  If IsNull(Param) Then
     NVLi = Default
  Else
     NVLi = CLng(Param)
  End If
End Function

Function NVLs$(Param As Variant, Default$)
  If IsNull(Param) Then
     NVLs$ = Default$
  Else
     NVLs$ = Param
  End If
End Function

Function ReadVarLenString$(ByVal FileHandle%)
  'This function reads a string of variable length from a
  'binary file opened as #FileHandle%. The filepointer
  'points to an integer containing the length of the
  'string that follows this integer.
  Get #FileHandle%, , i%
  s$ = Space$(i%)
  Get #FileHandle%, , s$
  ReadVarLenString$ = s$
End Function

Public Sub RestoreWaitLevel(wl As Long)
  Wait_Level = wl + 1
  Wait False
End Sub

Function SearchWindow(Search$) As Integer
  'This function is used to look for an active application.
  'The function scans all 'top windows' and returns the handle
  'of the first window it finds that contains the Search$
  'parameter in it's caption.
  'If no window is found, it returns 0.
  'See 'Visual Basic How-To' page 284 for example.
  Dim capt As String * 256
  dest$ = UCase$(Search$)
  wnd% = FindWindow(0&, 0&)
  wnd% = GetWindow(wnd%, GH_HWNDFIRST)
  While wnd% <> 0
    tchars% = GetWindowText(wnd%, capt, 256)
    If tchars% > 0 Then
       Source$ = UCase$(Left$(capt, tchars%))
       If InStr(Source$, dest$) > 0 Then
          SearchWindow = wnd%
          Exit Function
       End If
    End If
    wnd% = GetNextWindow(wnd%, GW_HWNDNEXT)
  Wend
  SearchWindow = 0
End Function

Sub SelectAllText(c As Control)
  'Selects all text in a text control
  
  If TypeOf c Is TextBox Then
     c.SelStart = 0
     c.SelLength = Len(c.Text)
  ElseIf TypeOf c Is ComboBox And c.Style < 2 Then
     c.SelStart = 0
     c.SelLength = Len(c.Text)
  End If
  
End Sub

Sub SetIntInMyINI(File As String, Section As String, Keyword As String, Value As Integer)
  SetStrInMyINI File, Section, Keyword, Format$(Value, "0")
End Sub

Sub SetIntInWININI(Section As String, Keyword As String, Value As Integer)
  SetStrInWININI Section, Keyword, Format$(Value, "0")
End Sub

Sub SetListboxTabStops(Lstbox As ListBox, NbrOfTabstops%, i1%, i2%, i3%, i4%)
  ReDim Tabs(1 To 10) As Long
  Tabs(1) = i1%
  Tabs(2) = i2%
  Tabs(3) = i3%
  Tabs(4) = i4%
  adr& = lstrcpy(Tabs(1), Tabs(1))
  d& = SendMessage(Lstbox.hwnd, (&H413), (NbrOfTabstops%), (adr&))
End Sub

Sub SetStrInMyINI(File As String, Section As String, Keyword As String, Value As String)
  valid% = WritePrivateProfileString%(Section, Keyword, Value, File)
End Sub

Sub SetStrInWININI(Section As String, Keyword As String, Value As String)
  valid% = WriteProfileString%(Section, Keyword, Value)
End Sub

Sub ToUpper(KeyAscii As Integer)
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Sub Wait(ByVal Status As Boolean)
' This subroutine accepts true or false as parameter
' and sets the mousepointer to hourglass or to default.
' Calls to this routine may be nested

  If Status Then
     Wait_Level = Wait_Level + 1
     Screen.MousePointer = 11
  Else
     If Wait_Level > 0 Then Wait_Level = Wait_Level - 1
     If Wait_Level = 0 Then Screen.MousePointer = 0
  End If

End Sub

Sub WaitForShell(x%)
  'This subroutine waits for a shelled process to end.
  'You can call this routine as follows :
  ' WaitForShell Shell(..., ...)

  While GetModuleUsage(x%) > 0
    DoEvents
  Wend

End Sub

Sub WindowTopMost(hwnd As Long, State As Long)
  'Use the global constants HWND_TOPMOST and HWND_NOTOPMOST
  'to set the window TOPMOST or not.

  d& = SetWindowPos(hwnd, State, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

End Sub

