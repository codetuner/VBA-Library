Attribute VB_Name = "OledbImport"
Option Compare Database
Option Explicit

''' <summary>
''' Returns the given string, or the matching registry value if the given string is found as key in the registry under
'''   HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ConnectionStrings\OleDB
''' To store a connectionstring, use RegEdit or the following code:
'''   SaveSetting "ConnectionStrings", "OleDB", "AdventureWorks", "Provider=sqloledb;Server=(local);Database=AdventureWorks;Trusted_Connection=yes;"
''' </summary>
Public Function GetOledbConnectionString(ByVal oledbConnectionStringOrRegkey As String)
    GetOledbConnectionString = GetSetting("ConnectionStrings", "OleDB", oledbConnectionStringOrRegkey, oledbConnectionStringOrRegkey)
End Function

''' <summary>
''' Creates a table with the given targetTableName and fills it with data obtained from the given query.
''' To create the table with no rows, use a SELECT TOP 0 ... query.
''' Returns number of rows imported.
''' </summary>
Public Function ImportToTable(ByVal targetTableName As String, ByVal oledbConnectionStringOrRegkey As String, ByVal sql As String) As Long
    
    ' Verify table does not yet exist:
    Dim td As TableDef
    For Each td In Application.CurrentDb.TableDefs
        If td.Name = targetTableName Then
            'For Each fld In td.Fields
            '    Debug.Print fld.Name, fld.Type, fld.Size
            'Next
            Err.Raise 5, "AdodbImport", "Table with name '" & targetTableName & "' already exists."
        End If
    Next
    
    ' Lookup connectionstring in registry:
    oledbConnectionStringOrRegkey = GetOledbConnectionString(oledbConnectionStringOrRegkey)
    
    Dim adoconn As New ADODB.Connection
    Dim adocmd As New ADODB.Command
    Dim adors As ADODB.Recordset
    
    ' Execute query:
    adoconn.Open oledbConnectionStringOrRegkey
    Set adocmd.ActiveConnection = adoconn
    adocmd.CommandText = sql
    Set adors = adocmd.Execute
    
    ' Create table with initial #LocalID field:
    Dim newtd As TableDef
    Dim newfld As field
    Set newtd = CurrentDb.CreateTableDef(targetTableName)
    Set newfld = newtd.CreateField("#LocalID", DataTypeEnum.dbLong)
    newfld.Attributes = dbAutoIncrField
    newtd.Fields.Append newfld
    
    ' Add additional fields from query:
    Dim adofld As Object
    For Each adofld In adors.Fields
        Set newfld = newtd.CreateField
        newfld.Name = adofld.Name
        Select Case adofld.Type
            Case ADODB.DataTypeEnum.adBigInt:
                newfld.Type = DataTypeEnum.dbLong
            Case ADODB.DataTypeEnum.adBoolean:
                newfld.Type = DataTypeEnum.dbBoolean
            Case ADODB.DataTypeEnum.adCurrency:
                newfld.Type = DataTypeEnum.dbCurrency
            Case ADODB.DataTypeEnum.adDate:
                newfld.Type = DataTypeEnum.dbDate
            Case ADODB.DataTypeEnum.adDBDate:
                newfld.Type = DataTypeEnum.dbDate
            Case ADODB.DataTypeEnum.adDBTimeStamp:
                newfld.Type = DataTypeEnum.dbDate
            Case ADODB.DataTypeEnum.adDecimal:
                newfld.Type = DataTypeEnum.dbCurrency
            Case ADODB.DataTypeEnum.adNumeric:
                newfld.Type = DataTypeEnum.dbCurrency
            Case ADODB.DataTypeEnum.adDouble:
                newfld.Type = DataTypeEnum.dbDouble
            'Case ADODB.DataTypeEnum.adGUID:
            '    newfld.Type = DataTypeEnum.dbText
            '    newfld.Size = 64
            Case ADODB.DataTypeEnum.adInteger:
                newfld.Type = DataTypeEnum.dbLong
            Case ADODB.DataTypeEnum.adSingle:
                newfld.Type = DataTypeEnum.dbSingle
            Case ADODB.DataTypeEnum.adSmallInt:
                newfld.Type = DataTypeEnum.dbInteger
            Case ADODB.DataTypeEnum.adTinyInt:
                newfld.Type = DataTypeEnum.dbByte
            Case ADODB.DataTypeEnum.adVarChar:
                newfld.Type = DataTypeEnum.dbText
                newfld.Size = adofld.DefinedSize
                newfld.AllowZeroLength = True
            Case ADODB.DataTypeEnum.adVarWChar:
                If adofld.DefinedSize <= 256 Then
                    newfld.Type = DataTypeEnum.dbText
                    newfld.Size = adofld.DefinedSize
                    newfld.AllowZeroLength = True
                Else
                    newfld.Type = DataTypeEnum.dbMemo
                End If
            Case ADODB.DataTypeEnum.adWChar:
                If adofld.DefinedSize <= 256 Then
                    newfld.Type = DataTypeEnum.dbText
                    newfld.Size = adofld.DefinedSize
                    newfld.AllowZeroLength = True
                Else
                    newfld.Type = DataTypeEnum.dbMemo
                End If
            Case ADODB.DataTypeEnum.adLongVarChar:
                newfld.Type = DataTypeEnum.dbMemo
            Case ADODB.DataTypeEnum.adLongVarWChar:
                newfld.Type = DataTypeEnum.dbMemo
            Case ADODB.DataTypeEnum.adBSTR:
                If adofld.DefinedSize <= 256 Then
                    newfld.Type = DataTypeEnum.dbText
                    newfld.Size = adofld.DefinedSize
                    newfld.AllowZeroLength = True
                Else
                    newfld.Type = DataTypeEnum.dbMemo
                End If
            Case ADODB.DataTypeEnum.adChar:
                If adofld.DefinedSize <= 256 Then
                    newfld.Type = DataTypeEnum.dbText
                    newfld.Size = adofld.DefinedSize
                    newfld.AllowZeroLength = True
                Else
                    newfld.Type = DataTypeEnum.dbMemo
                End If
            Case ADODB.DataTypeEnum.adGUID:
                newfld.Type = DataTypeEnum.dbText
                newfld.Size = 38
                newfld.AllowZeroLength = False
            Case Else:
                Err.Raise 5, "AdodbImport", "Type of field '" & adofld.Name & "' not supported."
        End Select
        newtd.Fields.Append newfld
    Next
    
    ' Create a primary key index based in #LocalID:
    Dim idx As Index
    Set idx = newtd.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Unique = True
    idx.Fields.Append idx.CreateField("#LocalID")
    newtd.Indexes.Append idx
    
    ' Create table:
    CurrentDb.TableDefs.Append newtd

    ' Open recordset on table:
    Dim tdrs As Recordset
    Set tdrs = CurrentDb.OpenRecordset(targetTableName, dbOpenTable)

    ' Copy rows:
    Dim rowcount As Long
    While Not adors.EOF
        
        tdrs.AddNew
        For Each adofld In adors.Fields
            tdrs(adofld.Name).Value = adofld.Value
        Next
        tdrs.Update
        rowcount = rowcount + 1
        
        adors.MoveNext
    Wend

    ' Free resources:
    tdrs.Close
    Set tdrs = Nothing
    adors.Close
    adoconn.Close
    Set adors = Nothing
    Set adocmd = Nothing
    Set adoconn = Nothing
    
    ' Return rowcount:
    ImportToTable = rowcount

End Function

''' <summary>
''' Append rows returned from the given query to the target table.
''' Returns number of rows appended.
''' </summary>
Public Function AppendToTable(ByVal targetTableName As String, ByVal oledbConnectionStringOrRegkey As String, ByVal sql As String) As Long

    ' Lookup connectionstring in registry:
    oledbConnectionStringOrRegkey = GetOledbConnectionString(oledbConnectionStringOrRegkey)
    
    Dim adoconn As New ADODB.Connection
    Dim adocmd As New ADODB.Command
    Dim adors As ADODB.Recordset
    
    ' Execute query:
    adoconn.Open oledbConnectionStringOrRegkey
    Set adocmd.ActiveConnection = adoconn
    adocmd.CommandText = sql
    Set adors = adocmd.Execute

    ' Open recordset on table:
    Dim tdrs As Recordset
    Set tdrs = CurrentDb.OpenRecordset(targetTableName, dbOpenTable)

    ' Build a fieldmap so fields which are missing in the target table can easily be skipped without error:
    Dim fieldmap As New Collection
    Dim fld As field
    For Each fld In tdrs.Fields
        fieldmap.Add True, fld.Name
    Next
    On Error Resume Next
    Dim adofld As Object
    For Each adofld In adors.Fields
        fieldmap.Add False, adofld.Name
    Next
    On Error GoTo 0

    ' Copy rows:
    Dim rowcount As Long
    While Not adors.EOF
        
        tdrs.AddNew
        For Each adofld In adors.Fields
            If fieldmap(adofld.Name) Then
                ' Copy field value:
                tdrs(adofld.Name).Value = adofld.Value
                '' Fix if string field is empty string, make it null:
                'If adofld.Type = 129 Or adofld.Type = 201 Or adofld.Type = 203 Or adofld.Type = 200 Or adofld.Type = 202 Or adofld.Type = 130 Then
                '    If adofld.value = "" Then
                '        adofld.value = Null
                '    End If
                'End If
            End If
        Next
        tdrs.Update
        rowcount = rowcount + 1
        
        adors.MoveNext
    Wend

    ' Free resources:
    tdrs.Close
    Set tdrs = Nothing
    adors.Close
    adoconn.Close
    Set adors = Nothing
    Set adocmd = Nothing
    Set adoconn = Nothing

    ' Return rowcount:
    AppendToTable = rowcount

End Function
