Attribute VB_Name = "DateLocalization"
' ******* COPYRIGHT NOTICE *******
' This module was written by Philipp Stiefel and published at https://codekabinett.com/rdumps.php?Lang=2&targetDoc=format-date-language-country-vba-access
' You are granted permission to use this module in your own application free of charge, as long as you leave this copyright notice in place.
' You may NOT publish this sample on its own without significant functional changes or enhancements without written consent of the original author!

Option Compare Database
Option Explicit

Private Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer ' Assumes vbSunday as FirstDayOfWeek
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type

' this enum is based on the "Time Flags for GetTimeFormat." from WinNls.h
Public Enum TimeFormat
    NoMinutesOrSeconds = &H1    ' do not use minutes or seconds
    NoSeconds = &H2             ' do not use seconds
    NoTimeMarker = &H4          ' do not use time marker
    Force24HourFormat = &H8     ' always use 24 hour format
    Locale_NoUserOverride = &H80000000           ' // do not use user overrides
End Enum

' this enum is based on the "Date Flags for GetDateFormat." from WinNls.h
Public Enum DateFormat
    ShortDate = &H1         ' use short date picture
    LongDate = &H2         ' use long date picture
    Use_Alt_Calendar = &H4              ' use alternate calendar (if any)
    YearMonth = &H8         ' use year month picture
    LtrReading = &H10        ' add marks for left to right reading order layout
    RtlReading = &H20         ' add marks for right to left reading order layout
    AutoLayout = &H40        ' add appropriate marks for left-to-right or right-to-left reading order layout
    Locale_NoUserOverride = &H80000000           ' // do not use user overrides
End Enum


' These locale names are not used in the code, they are included here for informational purposes only.
Private Const LOCALE_NAME_USER_DEFAULT      As String = vbNullString
Private Const LOCALE_NAME_INVARIANT         As String = ""
Private Const LOCALE_NAME_SYSTEM_DEFAULT    As String = "!x-sys-default-locale"


Private Declare PtrSafe Function GetDateFormatEx Lib "Kernel32" ( _
    ByVal lpLocaleName As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef lpDate As SYSTEMTIME, _
    ByVal lpFormat As LongPtr, _
    ByVal lpDateStr As LongPtr, _
    ByVal cchDate As Long, _
    ByVal lpCalendar As LongPtr _
) As Long

Private Declare PtrSafe Function GetTimeFormatEx Lib "Kernel32" ( _
    ByVal lpLocaleName As LongPtr, _
    ByVal dwFlags As Long, _
    ByRef lpTime As SYSTEMTIME, _
    ByVal lpFormat As LongPtr, _
    ByVal lpTimeStr As LongPtr, _
    ByVal cchTime As Long _
) As Long

' For a list of valid customFormatPicture options see the following resources
' Day, Month, Year, and Era Format Pictures: https://docs.microsoft.com/en-us/windows/win32/intl/day--month--year--and-era-format-pictures
' Hour, Minute, and Second Format Pictures: https://docs.microsoft.com/en-us/windows/win32/intl/hour--minute--and-second-format-pictures


Public Function FormatDateForLocale(ByVal theDate As Date, ByVal LocaleName As String, _
                                    Optional ByVal format As DateFormat = 0, Optional ByVal customFormatPicture As String = vbNullString _
                                    ) As String
    Dim retVal As String
    Dim formattedDateBuffer As String
    Dim sysTime As SYSTEMTIME
    Dim apiRetVal As Long
    
    Const BUFFER_CHARCOUNT As Long = 50
    
    sysTime = DateToSystemTime(theDate)
    
    formattedDateBuffer = String(BUFFER_CHARCOUNT, vbNullChar)
    apiRetVal = GetDateFormatEx(StrPtr(LocaleName), format, sysTime, StrPtr(customFormatPicture), StrPtr(formattedDateBuffer), BUFFER_CHARCOUNT, 0)
   
    If apiRetVal > 0 Then
        retVal = Left(formattedDateBuffer, apiRetVal - 1)
    End If
    
    FormatDateForLocale = retVal

End Function

Public Function FormatTimeForLocale(ByVal theDate As Date, ByVal LocaleName As String, _
                                    Optional ByVal format As TimeFormat = 0, Optional ByVal customFormatPicture As String = vbNullString _
                                    ) As String

    Dim retVal As String
    Dim formattedTimeBuffer As String
    Dim sysTime As SYSTEMTIME
    Dim apiRetVal As Long
    
    Const BUFFER_CHARCOUNT As Long = 30
    
    sysTime = DateToSystemTime(theDate)
        
    formattedTimeBuffer = String(BUFFER_CHARCOUNT, vbNullChar)
    apiRetVal = GetTimeFormatEx(StrPtr(LocaleName), format, sysTime, StrPtr(customFormatPicture), StrPtr(formattedTimeBuffer), BUFFER_CHARCOUNT)
 
    If apiRetVal > 0 Then
        retVal = Left(formattedTimeBuffer, apiRetVal - 1)
    End If
 
    FormatTimeForLocale = retVal

End Function

Private Function DateToSystemTime(ByVal theDate As Date) As SYSTEMTIME
    
    Dim retVal As SYSTEMTIME
    
    With retVal
        .wDay = Day(theDate)
        .wMonth = Month(theDate)
        .wYear = Year(theDate)
        .wHour = Hour(theDate)
        .wMinute = Minute(theDate)
        .wSecond = Second(theDate)
        .wDayOfWeek = Weekday(theDate, vbSunday)
    End With

    DateToSystemTime = retVal

End Function


' ******* COPYRIGHT NOTICE *******
' This module was written by Philipp Stiefel and published at https://codekabinett.com/rdumps.php?Lang=2&targetDoc=format-date-language-country-vba-access
' You are granted permission to use this module in your own application free of charge, as long as you leave this copyright notice in place.
' You may NOT publish this sample on its own without significant functional changes or enhancements without written consent of the original author!



