Attribute VB_Name = "Clipboard"
Option Compare Database
Option Explicit

' Based on:
' https://msdn.microsoft.com/en-us/library/office/ff194373.aspx

#If VBA7 Then
Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
#Else
Private Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "User32" () As Long
Private Declare Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags&, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
#End If

Function GetText()
   Dim hClipMemory As LongPtr
   Dim lpClipMemory As LongPtr
   Dim retval As LongPtr
   Dim MyString As String
 
   If OpenClipboard(0&) = 0 Then
      Err.Raise 5, "Clipboard", "Cannot open Clipboard. It may already be open."
   End If
          
   ' Obtain the handle to the global memory
   ' block that is referencing the text.
   hClipMemory = GetClipboardData(1)
   If IsNull(hClipMemory) Then
      CloseClipboard
      Err.Raise 5, "Clipboard", "Could not allocate memory for clipboard reading."
   End If
 
   ' Lock Clipboard memory so we can reference
   ' the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)
 
   If Not IsNull(lpClipMemory) Then
      MyString = Space$(4096) 'MAX_LENGTH
      retval = lstrcpy(MyString, lpClipMemory)
      retval = GlobalUnlock(hClipMemory)
       
      ' Peel off the null terminating character.
      MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
      GetText = MyString
      CloseClipboard
   Else
      CloseClipboard
      Err.Raise 5, "Clipboard", "Could not lock memory to copy string from."
   End If
 
End Function
