Attribute VB_Name = "modComprehensiveImport"
Option Compare Database
Option Explicit

' ============================================================================
' Comprehensive Data Import Module
' Imports data from CSV templates with full field support
' ============================================================================

Private Const IMPORT_PATH As String = "C:\Import\"  ' Change to your path

' Dictionary to map OrderNumber -> SOID after import
Private dictOrderMap As Object

Public Sub ImportAllData()
    ' Main entry point - imports all three tables in correct order

    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    ' Initialize mapping dictionary
    Set dictOrderMap = CreateObject("Scripting.Dictionary")

    ' Confirm with user
    If MsgBox("This will import data from CSV files in:" & vbCrLf & _
              IMPORT_PATH & vbCrLf & vbCrLf & _
              "Files expected:" & vbCrLf & _
              "- ImportTemplate_SalesOrders.csv" & vbCrLf & _
              "- ImportTemplate_SalesOrderEntry.csv" & vbCrLf & _
              "- ImportTemplate_OrderSeq.csv" & vbCrLf & vbCrLf & _
              "Continue?", vbYesNo + vbQuestion, "Confirm Import") = vbNo Then
        Exit Sub
    End If

    ' Start transaction
    DBEngine.BeginTrans

    ' Step 1: Import SalesOrders
    Debug.Print "Step 1: Importing SalesOrders..."
    ImportSalesOrders

    ' Step 2: Build OrderNumber -> SOID mapping
    Debug.Print "Step 2: Building SOID mapping..."
    BuildOrderMapping

    ' Step 3: Import SalesOrderEntry with SOID mapping
    Debug.Print "Step 3: Importing SalesOrderEntry..."
    ImportSalesOrderEntry

    ' Step 4: Import OrderSeq
    Debug.Print "Step 4: Importing OrderSeq..."
    ImportOrderSeq

    ' Commit transaction
    DBEngine.CommitTrans

    ' Report results
    Dim msg As String
    msg = "Import completed successfully!" & vbCrLf & vbCrLf & _
          "Records imported:" & vbCrLf & _
          "- SalesOrders: " & DCount("*", "SalesOrders") & vbCrLf & _
          "- SalesOrderEntry: " & DCount("*", "SalesOrderEntry") & vbCrLf & _
          "- OrderSeq: " & DCount("*", "OrderSeq")

    MsgBox msg, vbInformation, "Import Complete"

    Exit Sub

ErrorHandler:
    DBEngine.Rollback
    MsgBox "Import failed: " & Err.Description & vbCrLf & _
           "All changes have been rolled back.", vbCritical, "Import Error"
End Sub

Private Sub ImportSalesOrders()
    ' Import SalesOrders from CSV using staging table approach

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsStage As DAO.Recordset
    Dim sql As String

    Set db = CurrentDb

    ' Create staging table
    On Error Resume Next
    db.Execute "DROP TABLE tblImportSalesOrders", dbFailOnError
    On Error GoTo 0

    ' Import CSV to staging table
    DoCmd.TransferText acImportDelim, , "tblImportSalesOrders", _
        IMPORT_PATH & "ImportTemplate_SalesOrders.csv", True

    ' Insert into SalesOrders from staging
    Set rsStage = db.OpenRecordset("SELECT * FROM tblImportSalesOrders", dbOpenSnapshot)
    Set rs = db.OpenRecordset("SalesOrders", dbOpenDynaset)

    Do Until rsStage.EOF
        rs.AddNew

        ' Required fields
        rs!OrderNumber = rsStage!OrderNumber
        rs!OrderType = Nz(rsStage!OrderType, "SALES")
        rs!BaseToken = rsStage!BaseToken
        rs!SystemLetter = Nz(rsStage!SystemLetter, "P")
        rs!BackorderNo = Nz(rsStage!BackorderNo, 0)
        rs!CustomerName = rsStage!CustomerName

        ' Optional fields
        If Not IsNull(rsStage!CustomerCode) Then rs!CustomerCode = rsStage!CustomerCode
        If Not IsNull(rsStage!PONumber) Then rs!PONumber = rsStage!PONumber
        If Not IsNull(rsStage!DateReceived) Then rs!DateReceived = rsStage!DateReceived
        If Not IsNull(rsStage!DateBilled) Then rs!DateBilled = rsStage!DateBilled

        ' Boolean with default
        rs!ActiveFlag = IIf(LCase(Nz(rsStage!ActiveFlag, "true")) = "true", True, False)

        ' New feature fields
        If Not IsNull(rsStage!BatchID) And rsStage!BatchID <> "" Then
            rs!BatchID = rsStage!BatchID
        End If

        If Not IsNull(rsStage!DateCreated) Then
            rs!DateCreated = rsStage!DateCreated
        Else
            rs!DateCreated = Now()
        End If

        If Not IsNull(rsStage!DateBackorderCreated) Then
            rs!DateBackorderCreated = rsStage!DateBackorderCreated
        End If

        ' Cancellation fields
        If Not IsNull(rsStage!DateCanceled) Then rs!DateCanceled = rsStage!DateCanceled
        If Not IsNull(rsStage!CancelReason) Then rs!CancelReason = rsStage!CancelReason
        If Not IsNull(rsStage!CanceledBy) Then rs!CanceledBy = rsStage!CanceledBy

        rs.Update
        rsStage.MoveNext
    Loop

    rsStage.Close
    rs.Close

    ' Cleanup staging table
    db.Execute "DROP TABLE tblImportSalesOrders", dbFailOnError

    Debug.Print "  Imported " & DCount("*", "SalesOrders") & " SalesOrders records"
End Sub

Private Sub BuildOrderMapping()
    ' Build dictionary mapping OrderNumber -> SOID

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT SOID, OrderNumber FROM SalesOrders", dbOpenSnapshot)

    Do Until rs.EOF
        If Not dictOrderMap.Exists(rs!OrderNumber) Then
            dictOrderMap.Add rs!OrderNumber, rs!SOID
        End If
        rs.MoveNext
    Loop

    rs.Close

    Debug.Print "  Built mapping for " & dictOrderMap.Count & " orders"
End Sub

Private Sub ImportSalesOrderEntry()
    ' Import SalesOrderEntry from CSV, mapping OrderNumberDisplay to SOID

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsStage As DAO.Recordset
    Dim mappedSOID As Long
    Dim orderNum As String
    Dim skippedCount As Long

    Set db = CurrentDb

    ' Create staging table
    On Error Resume Next
    db.Execute "DROP TABLE tblImportSalesOrderEntry", dbFailOnError
    On Error GoTo 0

    ' Import CSV to staging table
    DoCmd.TransferText acImportDelim, , "tblImportSalesOrderEntry", _
        IMPORT_PATH & "ImportTemplate_SalesOrderEntry.csv", True

    ' Insert into SalesOrderEntry from staging
    Set rsStage = db.OpenRecordset("SELECT * FROM tblImportSalesOrderEntry", dbOpenSnapshot)
    Set rs = db.OpenRecordset("SalesOrderEntry", dbOpenDynaset)

    skippedCount = 0

    Do Until rsStage.EOF
        orderNum = Nz(rsStage!OrderNumberDisplay, "")

        ' Look up SOID from OrderNumber
        If dictOrderMap.Exists(orderNum) Then
            mappedSOID = dictOrderMap(orderNum)

            rs.AddNew
            rs!SOID = mappedSOID
            rs!QualifierCode = rsStage!QualifierCode
            rs!SequenceNo = Nz(rsStage!SequenceNo, 0)
            rs!OrderNumberDisplay = orderNum
            rs!IsDeleted = IIf(LCase(Nz(rsStage!IsDeleted, "false")) = "true", True, False)

            If Not IsNull(rsStage!CreatedOn) Then
                rs!CreatedOn = rsStage!CreatedOn
            Else
                rs!CreatedOn = Now()
            End If

            rs.Update
        Else
            ' Try using SOID directly from CSV if OrderNumber lookup fails
            If Not IsNull(rsStage!SOID) Then
                rs.AddNew
                rs!SOID = rsStage!SOID
                rs!QualifierCode = rsStage!QualifierCode
                rs!SequenceNo = Nz(rsStage!SequenceNo, 0)
                rs!OrderNumberDisplay = orderNum
                rs!IsDeleted = IIf(LCase(Nz(rsStage!IsDeleted, "false")) = "true", True, False)

                If Not IsNull(rsStage!CreatedOn) Then
                    rs!CreatedOn = rsStage!CreatedOn
                Else
                    rs!CreatedOn = Now()
                End If

                rs.Update
            Else
                Debug.Print "  WARNING: No SOID found for OrderNumber: " & orderNum
                skippedCount = skippedCount + 1
            End If
        End If

        rsStage.MoveNext
    Loop

    rsStage.Close
    rs.Close

    ' Cleanup staging table
    db.Execute "DROP TABLE tblImportSalesOrderEntry", dbFailOnError

    Debug.Print "  Imported " & DCount("*", "SalesOrderEntry") & " SalesOrderEntry records"
    If skippedCount > 0 Then
        Debug.Print "  WARNING: Skipped " & skippedCount & " records with no matching SOID"
    End If
End Sub

Private Sub ImportOrderSeq()
    ' Import OrderSeq from CSV

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsStage As DAO.Recordset

    Set db = CurrentDb

    ' Create staging table
    On Error Resume Next
    db.Execute "DROP TABLE tblImportOrderSeq", dbFailOnError
    On Error GoTo 0

    ' Import CSV to staging table
    DoCmd.TransferText acImportDelim, , "tblImportOrderSeq", _
        IMPORT_PATH & "ImportTemplate_OrderSeq.csv", True

    ' Insert into OrderSeq from staging
    Set rsStage = db.OpenRecordset("SELECT * FROM tblImportOrderSeq", dbOpenSnapshot)
    Set rs = db.OpenRecordset("OrderSeq", dbOpenDynaset)

    Do Until rsStage.EOF
        rs.AddNew
        rs!Scope = rsStage!Scope
        rs!BaseToken = rsStage!BaseToken
        rs!QualifierCode = rsStage!QualifierCode
        rs!SystemLetter = rsStage!SystemLetter
        rs!NextSeq = rsStage!NextSeq

        If Not IsNull(rsStage!LastUpdated) Then
            rs!LastUpdated = rsStage!LastUpdated
        Else
            rs!LastUpdated = Now()
        End If

        rs.Update
        rsStage.MoveNext
    Loop

    rsStage.Close
    rs.Close

    ' Cleanup staging table
    db.Execute "DROP TABLE tblImportOrderSeq", dbFailOnError

    Debug.Print "  Imported " & DCount("*", "OrderSeq") & " OrderSeq records"
End Sub

' ============================================================================
' Utility Functions
' ============================================================================

Public Sub GenerateOrderSeqFromExisting()
    ' Generates OrderSeq records based on existing SalesOrders/SalesOrderEntry data
    ' Run this after importing if you didn't have OrderSeq data

    Dim db As DAO.Database
    Dim sql As String

    Set db = CurrentDb

    sql = "INSERT INTO OrderSeq (Scope, BaseToken, QualifierCode, SystemLetter, NextSeq, LastUpdated) " & _
          "SELECT s.OrderType, s.BaseToken, e.QualifierCode, s.SystemLetter, " & _
          "       MAX(e.SequenceNo) + 1, Now() " & _
          "FROM SalesOrders s " & _
          "INNER JOIN SalesOrderEntry e ON s.SOID = e.SOID " & _
          "GROUP BY s.OrderType, s.BaseToken, e.QualifierCode, s.SystemLetter"

    On Error Resume Next
    db.Execute sql, dbFailOnError

    If Err.Number = 0 Then
        MsgBox "Generated " & DCount("*", "OrderSeq") & " OrderSeq records from existing data.", _
               vbInformation, "OrderSeq Generated"
    Else
        MsgBox "Error generating OrderSeq: " & Err.Description, vbExclamation, "Error"
    End If
End Sub

Public Sub ValidateImportedData()
    ' Validates data integrity after import

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim issues As String
    Dim issueCount As Long

    Set db = CurrentDb
    issues = ""
    issueCount = 0

    ' Check for orphaned SalesOrderEntry records
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM SalesOrderEntry " & _
        "WHERE SOID NOT IN (SELECT SOID FROM SalesOrders)", dbOpenSnapshot)
    If rs!Cnt > 0 Then
        issues = issues & "- " & rs!Cnt & " orphaned SalesOrderEntry records" & vbCrLf
        issueCount = issueCount + 1
    End If
    rs.Close

    ' Check for invalid SystemLetter references
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM SalesOrders " & _
        "WHERE SystemLetter NOT IN (SELECT SystemLetter FROM SystemLetter)", dbOpenSnapshot)
    If rs!Cnt > 0 Then
        issues = issues & "- " & rs!Cnt & " orders with invalid SystemLetter" & vbCrLf
        issueCount = issueCount + 1
    End If
    rs.Close

    ' Check for invalid QualifierCode references
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM SalesOrderEntry " & _
        "WHERE QualifierCode NOT IN (SELECT QualifierCode FROM QualifierType)", dbOpenSnapshot)
    If rs!Cnt > 0 Then
        issues = issues & "- " & rs!Cnt & " entries with invalid QualifierCode" & vbCrLf
        issueCount = issueCount + 1
    End If
    rs.Close

    ' Check for duplicate OrderNumbers
    Set rs = db.OpenRecordset( _
        "SELECT OrderNumber, COUNT(*) AS Cnt FROM SalesOrders " & _
        "GROUP BY OrderNumber HAVING COUNT(*) > 1", dbOpenSnapshot)
    If Not rs.EOF Then
        issues = issues & "- Duplicate OrderNumbers found" & vbCrLf
        issueCount = issueCount + 1
    End If
    rs.Close

    ' Check for missing OrderSeq entries
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM " & _
        "(SELECT DISTINCT s.OrderType, s.BaseToken, e.QualifierCode, s.SystemLetter " & _
        " FROM SalesOrders s INNER JOIN SalesOrderEntry e ON s.SOID = e.SOID) AS Combos " & _
        "WHERE NOT EXISTS (SELECT 1 FROM OrderSeq q " & _
        "  WHERE q.Scope = Combos.OrderType " & _
        "  AND q.BaseToken = Combos.BaseToken " & _
        "  AND q.QualifierCode = Combos.QualifierCode " & _
        "  AND q.SystemLetter = Combos.SystemLetter)", dbOpenSnapshot)
    If rs!Cnt > 0 Then
        issues = issues & "- " & rs!Cnt & " missing OrderSeq entries (run GenerateOrderSeqFromExisting)" & vbCrLf
        issueCount = issueCount + 1
    End If
    rs.Close

    ' Report results
    If issueCount = 0 Then
        MsgBox "Data validation passed - no issues found!", vbInformation, "Validation Complete"
    Else
        MsgBox "Found " & issueCount & " issue(s):" & vbCrLf & vbCrLf & issues, _
               vbExclamation, "Validation Issues"
    End If
End Sub

Public Sub ExportCurrentData()
    ' Exports current table data to CSV files for backup/modification

    Dim exportPath As String
    exportPath = IMPORT_PATH & "Export_" & Format(Now(), "yyyymmdd_hhnnss") & "\"

    ' Create export folder
    MkDir exportPath

    ' Export tables
    DoCmd.TransferText acExportDelim, , "SalesOrders", exportPath & "SalesOrders.csv", True
    DoCmd.TransferText acExportDelim, , "SalesOrderEntry", exportPath & "SalesOrderEntry.csv", True
    DoCmd.TransferText acExportDelim, , "OrderSeq", exportPath & "OrderSeq.csv", True

    MsgBox "Data exported to:" & vbCrLf & exportPath, vbInformation, "Export Complete"
End Sub
