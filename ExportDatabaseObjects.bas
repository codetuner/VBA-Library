Attribute VB_Name = "ExportDatabaseObjects"
Option Compare Database
Option Explicit

''' From:
''' http://www.access-programmers.co.uk/forums/showthread.php?t=99179

Public Function ExportDatabaseObjectsFx()
    Call ExportDatabaseObjects
End Function

Public Sub ExportDatabaseObjects()

    Dim db As Database
    'Dim db As DAO.Database
    Dim td As TableDef
    Dim d As Document
    Dim c As Container
    Dim v As Variant
    Dim i As Integer
    Dim sExportLocation As String
    
    Set db = CurrentDb()
    
    'sExportLocation = CurrentDb.Name + ".export." & Format(DateTime.Now, "yyyyMMddTHHmmss")
    sExportLocation = CurrentDb.Name + " Parts"
    On Error Resume Next
    MkDir sExportLocation
    CreateObject("Scripting.FileSystemObject").DeleteFolder sExportLocation
    Pause 2, True
    MkDir sExportLocation
        
    On Error GoTo 0
'    On Error GoTo Err_ExportDatabaseObjects
    
    'For Each td In db.TableDefs 'Tables
    '    If Left(td.Name, 4) <> "MSys" Then
    '        DoCmd.TransferText AcTextTransferType.acExportDelim, "ExportDelimited", td.Name, sExportLocation & "\Table_" & td.Name & ".txt", True
    '    End If
    'Next td
    
    For Each v In Access.Application.CurrentData.AllTables
        If Left(v.Name, 4) <> "MSys" Then
            Access.Application.ExportXML acExportTable, v.Name, _
                sExportLocation & "\TableData_" & v.Name & ".xml", _
                sExportLocation & "\TableDef_" & v.Name & ".xsd", , , acUTF8
        End If
    Next
    
    Set c = db.Containers("Forms")
    For Each d In c.Documents
        Application.SaveAsText acForm, d.Name, sExportLocation & "\Form_" & d.Name & ".txt"
    Next d
    
    Set c = db.Containers("Reports")
    For Each d In c.Documents
        Application.SaveAsText acReport, d.Name, sExportLocation & "\Report_" & d.Name & ".txt"
    Next d
    
    Set c = db.Containers("Scripts")
    For Each d In c.Documents
        Application.SaveAsText acMacro, d.Name, sExportLocation & "\Macro_" & d.Name & ".txt"
    Next d
    
    Set c = db.Containers("Modules")
    For Each d In c.Documents
        Application.SaveAsText acModule, d.Name, sExportLocation & "\Module_" & d.Name & ".txt"
    Next d
    
    For i = 0 To db.QueryDefs.Count - 1
        Application.SaveAsText acQuery, db.QueryDefs(i).Name, sExportLocation & "\Query_" & db.QueryDefs(i).Name & ".txt"
    Next i
    
    Set db = Nothing
    Set c = Nothing
    
    MsgBox "All database objects have been exported as a text file to " & sExportLocation, vbInformation
    
    Shell "explorer.exe /select,""" & sExportLocation & """", vbNormalFocus
    
Exit_ExportDatabaseObjects:
    Exit Sub
    
Err_ExportDatabaseObjects:
    Debug.Print Err.Number & " - " & Err.Description
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_ExportDatabaseObjects
    
End Sub

Public Sub Pause(ByVal durationInSeconds As Double, ByVal withDoEvents As Boolean)

    Dim endTime As Double
    endTime = Timer() + durationInSeconds
    While Timer < endTime
        If withDoEvents Then DoEvents
    Wend

End Sub
