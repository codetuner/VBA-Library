Attribute VB_Name = "PersonalSharedMacros"
Option Explicit

Sub FixDecimalPoints()

    Dim c As range
    
    For Each c In Selection
        If c.Value = "NULL" Then
            c.Value = Null
        ElseIf InStr(c.Formula, "=") = 0 And InStr(c.Value, ".") > 0 Then
            c.Value = Val(c.Value)
        End If
    Next
    
End Sub

Sub CampareRanges()

    Dim range1Str As String
    Dim range1 As range
    Dim range2Str As String
    Dim range2 As range

    range1Str = InputBox("Range 1 to compare : ", "Compare Ranges", "Sheet1!A1")
    If range1Str = "" Then Exit Sub
    Set range1 = ActiveWorkbook.Sheets(1).range(range1Str)
    Set range1 = range1.CurrentRegion
    
    range2Str = InputBox("Range 2 to compare : ", "Compare Ranges", "Sheet2!A1")
    If range2Str = "" Then Exit Sub
    Set range2 = ActiveWorkbook.Sheets(1).range(range2Str)
    Set range2 = range2.CurrentRegion
    
    If MsgBox("Compare " & range1.Address & " with " & range2.Address & " ?") <> vbOK Then Exit Sub

End Sub

Sub CampareGivenRanges(ByVal range1 As range, ByVal range2 As range)

    ' CampareGivenRanges Application.Workbooks(2).Sheets("Source").Range("A1"), Application.Workbooks(3).Sheets("Source").Range("A1")

    Dim r As Integer, c As Integer
    
    'Set range1 = range1.CurrentRegion
    'Set range2 = range2.CurrentRegion
           
    For r = 0 To range1.CurrentRegion.Rows.Count - 1
        For c = 0 To range1.CurrentRegion.Columns.Count - 1
            If ("" & range1.Offset(r, c).Value) <> ("" & range2.Offset(r, c).Value) Then
                range2.Offset(r, c).Interior.ColorIndex = 3 'Red
            Else
                range2.Offset(r, c).Interior.ColorIndex = 2 'White
            End If
        Next
    Next

End Sub

Sub IgnoreAllErrorsOnSelection()

    Dim rangeToValidate As range
    
    ' Beperk range tot de usedrange van de active sheet:
    Set rangeToValidate = Intersect(Selection, Selection.Worksheet.UsedRange)
    
    ' Indien selectie buiten used range viel:
    If rangeToValidate Is Nothing Then
        MsgBox "Selecteer eerst cellen waar inhoud in staat.", vbInformation, "Ignore All Errors on Selection"
        Exit Sub
    End If
    
    ' Indien nog te groot: geef melding en breek af:
    If rangeToValidate.Rows.Count > 10000 Or rangeToValidate.Columns.Count > 1000 Then
        MsgBox "Een te groot vlak is geselecteerd." & vbNewLine & "Maak een kleinere selectie om schoon te maken.", vbExclamation, "Ignore All Errors on Selection"
        Exit Sub
    End If

    ' Overloop elke cel, en verwijder foutmeldingen:
    Dim cl As range
    Dim errnr As Integer
    For Each cl In rangeToValidate.Cells
        For errnr = 1 To 9
            If cl.Errors(errnr).Value = True Then
                cl.Errors(errnr).Ignore = True
            End If
        Next
    Next
    
End Sub

Sub CleanSelection()

    Dim rangeToClean As range
    
	' If a single cell is selected, select whole current region:
	If Selection.Cells.Count = 1 Then
	    Selection.CurrentRegion.Select
	End If
	
    ' Limit range to the used range of the active sheet:
    Set rangeToClean = Intersect(Selection, Selection.Worksheet.UsedRange)
    
    ' If selection was not in used range:
    If rangeToClean Is Nothing Then
        MsgBox "Select cells with content first.", vbInformation, "Clean Selection Macro"
        Exit Sub
    End If
    
    ' If selection too big, abort:
    If rangeToClean.Rows.Count > 10000 Or rangeToClean.Columns.Count > 1000 Then
        MsgBox "You have selected a too large region." & vbNewLine & "Try cleaning smaller regions at a time.", vbExclamation, "Clean Selection Macro"
        Exit Sub
    End If

    ' Loop over all cells, and if their value is a string, clean it:
    Dim cl As range
    For Each cl In rangeToClean.Cells
        If TypeName(cl.Value) = "String" Then
            cl.Value = CleanString(cl.Value)
        End If
    Next

End Sub

Function CleanString(ByVal text As String) As String

    Dim regel As Integer
    Dim regels() As String
    Dim lengte As Integer

    ' Replace tabs by spaces:
    text = Replace(text, vbTab, " ")
    
    ' Replace newline/#13/#10 by #10:
    text = Replace(text, vbNewLine, Chr(10))
    text = Replace(text, Chr(13), Chr(10))
    'text = Replace(text, Chr(10), ", ")
    
    ' Split text on each #10 and trim every line:
    regels = Split(text, Chr(10))
    For regel = 0 To UBound(regels)
        regels(regel) = Trim(regels(regel))
    Next
    text = Join(regels, ", ")
    
    ' Remove double spaces until no more:
    Do
        lengte = Len(text)
        text = Replace(text, "  ", " ")
    Loop Until lengte = Len(text)

    ' Return result:
    CleanString = text

End Function

Public Sub ExportSelectionToUnicodeCsv()

    Dim r As range
    
    ' Limit the selected range to the UsedRange of the selection worksheet:
    Set r = Intersect(Selection, Selection.Worksheet.UsedRange)
    
    ' If intersect is empty:
    If r Is Nothing Then
        MsgBox "Please select cells with values first.", vbInformation, "Export Selection to Unicode CSV macro"
        Exit Sub
    End If
    
    Dim filename As String
    filename = Replace(ActiveWorkbook.FullName, ".xlsx", "") & "(" & Selection.Worksheet.Name & ").ucsv"
    filename = InputBox("Save " & Replace(r.Address, "$", "") & " as: ", "Export to Unicode CSV", filename)
    If filename = "" Then Exit Sub
    
    ExportToUnicodeCsv filename, r, ";", True
    
    MsgBox "File written:" & vbNewLine & filename

End Sub

Public Sub ExportToUnicodeCsv(ByVal filename As String, ByVal table As range, ByVal colDelimiter As String, ByVal includeCultureHeader)

    Dim rng As range
    Dim c As Integer, r As Integer
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim v As Variant
    
    Set ts = fso.OpenTextFile(filename, ForWriting, True, TristateTrue) 'TristateTrue=Unicode
    
    If includeCultureHeader Then ts.WriteLine "#Culture: en-US"
    
    Set rng = table.Cells(1, 1)
    For r = 0 To table.Rows.Count - 1
        v = rng.Offset(r, 0).Value
        GoSub WriteValue
        For c = 1 To table.Columns.Count - 1
            ts.Write colDelimiter
            v = rng.Offset(r, c).Value
            GoSub WriteValue
        Next
        ts.WriteLine
    Next
    
Finally:
    ts.Close

Exit Sub

WriteValue:
    If IsEmpty(v) Then
        'Do nothing
    ElseIf TypeName(v) = "String" Then
        v = CleanString(v)
        If InStr(v, colDelimiter) > 0 Then
            ts.Write """" & Replace(v, """", """""") & """"
        Else
            ts.Write v
        End If
    ElseIf TypeName(v) = "Double" Then
        ts.Write Trim(Str(v))
    ElseIf TypeName(v) = "Date" Then
        ts.Write Format(v, "yyyy/MM/dd HH:mm:ss")
    Else
        Debug.Print TypeName(v)
        ts.Write Str(v)
    End If
Return

ErrorHandler:
    MsgBox Err.Description & vbNewLine & "Cell: " & Replace(table.Offset(r, c).Address, "$", ""), vbCritical, "Export to Unicode CSV"
    Resume Finally
End Sub

