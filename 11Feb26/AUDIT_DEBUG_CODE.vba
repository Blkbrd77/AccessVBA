'================================================================================
' AUDIT DEBUG CODE
' Purpose: Enhanced versions with debug logging to diagnose issues
' Date: 2026-02-11
'
' Instructions:
' 1. Use this code to replace the non-working versions
' 2. Open Immediate Window (Ctrl+G) to see debug output
' 3. Run the stamp operation
' 4. Review debug messages to identify the problem
'================================================================================

'================================================================================
' ENHANCED LogOrderAction - WITH DEBUG LOGGING
' Module: basAuditLogging
' Replace the existing LogOrderAction with this version
'================================================================================

Public Sub LogOrderAction( _
    ByVal SOID As Long, _
    ByVal OrderNumber As String, _
    ByVal action As String, _
    Optional ByVal oldStatus As String = "", _
    Optional ByVal newStatus As String = "", _
    Optional ByVal reason As String = "" _
)
    ' Order-specific audit logging with debug output
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Debug.Print "========================================="
    Debug.Print "LogOrderAction Called"
    Debug.Print "  SOID: " & SOID
    Debug.Print "  OrderNumber: " & OrderNumber
    Debug.Print "  Action: " & action
    Debug.Print "  OldStatus: " & oldStatus
    Debug.Print "  NewStatus: " & newStatus
    Debug.Print "  Reason: " & reason
    Debug.Print "  ActionBy: " & Environ("USERNAME")
    Debug.Print "  ComputerName: " & Environ("COMPUTERNAME")
    Debug.Print "  Timestamp: " & Now

    On Error GoTo EH

    Set db = CurrentDb
    Debug.Print "  Database opened"

    Set rs = db.OpenRecordset("tblOrderAudit", dbOpenDynaset)
    Debug.Print "  Recordset opened"

    rs.AddNew
    Debug.Print "  AddNew called"

    rs!SOID = SOID
    rs!OrderNumber = OrderNumber
    rs!action = action
    rs!ActionTimestamp = Now
    rs!ActionBy = Environ("USERNAME")
    rs!ComputerName = Environ("COMPUTERNAME")
    rs!oldStatus = oldStatus
    rs!newStatus = newStatus
    rs!reason = reason

    Debug.Print "  All fields set"

    rs.Update
    Debug.Print "  Update called - RECORD SAVED!"

    rs.Close
    Debug.Print "LogOrderAction SUCCESS"
    Debug.Print "========================================="

    Exit Sub

EH:
    Debug.Print "========================================="
    Debug.Print "LogOrderAction ERROR!"
    Debug.Print "  Error Number: " & Err.Number
    Debug.Print "  Error Description: " & Err.Description
    Debug.Print "  Error Source: " & Err.Source
    Debug.Print "========================================="

    ' Try to show more specific error info
    Select Case Err.Number
        Case 3078
            Debug.Print "  >> Table 'tblOrderAudit' not found!"
            Debug.Print "  >> Run CreateAuditTables() to create it"
        Case 3265
            Debug.Print "  >> Field name mismatch!"
            Debug.Print "  >> Check field names in tblOrderAudit"
        Case 3131
            Debug.Print "  >> Invalid field value!"
            Debug.Print "  >> Check data types match table design"
        Case 3027
            Debug.Print "  >> Database is read-only!"
            Debug.Print "  >> Check file permissions"
    End Select

    ' Silently ignore errors (don't crash the app)
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Sub


'================================================================================
' ENHANCED cmdStampBilled_Click - WITH DEBUG LOGGING
' Form: frmOrderList
' Replace the existing cmdStampBilled_Click with this version
'================================================================================

Private Sub cmdStampBilled_Click()
    On Error GoTo EH

    Dim oldDateBilled As Variant
    Dim newDateBilled As Variant
    Dim lngSOID As Long
    Dim sOrderNumber As String

    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "STAMP BILLED: START"
    Debug.Print "========================================="

    ' Must be on an existing record
    If Me.NewRecord Then
        Debug.Print "ERROR: NewRecord - exiting"
        MsgBox "Please select or save a record before stamping.", vbExclamation
        Exit Sub
    End If

    ' Capture current state BEFORE making changes
    lngSOID = Nz(Me!SOID, 0)
    sOrderNumber = Nz(Me!OrderNumber, "")
    oldDateBilled = Me!DateBilled  ' May be Null

    Debug.Print "Current Record:"
    Debug.Print "  SOID: " & lngSOID
    Debug.Print "  OrderNumber: " & sOrderNumber
    Debug.Print "  Current DateBilled: " & IIf(IsNull(oldDateBilled), "(null)", oldDateBilled)

    If lngSOID = 0 Then
        Debug.Print "ERROR: SOID is 0 - invalid record"
        MsgBox "Invalid record selected.", vbExclamation
        Exit Sub
    End If

    ' Clear any old values
    Debug.Print "Clearing TempVars..."
    On Error Resume Next
    TempVars.Remove "StampBilledDate"
    TempVars.Remove "StampBilledResult"
    On Error GoTo EH

    ' Open the dialog modally
    Debug.Print "Opening dialog..."
    DoCmd.OpenForm "dlgStampBilledDate", WindowMode:=acDialog
    Debug.Print "Dialog closed"

    ' Check result
    Dim result As String
    result = Nz(TempVars("StampBilledResult"), "Cancel")
    Debug.Print "Dialog result: " & result

    If result <> "OK" Then
        Debug.Print "User canceled - exiting"
        Exit Sub
    End If

    ' Get new value from dialog
    newDateBilled = TempVars("StampBilledDate")
    Debug.Print "New DateBilled from dialog: " & newDateBilled

    If IsNull(newDateBilled) Then
        Debug.Print "ERROR: newDateBilled is Null!"
        MsgBox "No date was selected.", vbExclamation
        Exit Sub
    End If

    ' Write the chosen date to your bound field
    Debug.Print "Updating Me!DateBilled field..."
    Me!DateBilled = newDateBilled
    Debug.Print "Field updated"

    ' Persist immediately
    If Me.Dirty Then
        Debug.Print "Record is dirty - saving..."
        Me.Dirty = False
        Debug.Print "Record saved"
    Else
        Debug.Print "Record was not dirty"
    End If

    ' Format values for audit
    Dim oldValue As String
    Dim newValue As String
    oldValue = IIf(IsNull(oldDateBilled), "", Format(oldDateBilled, "yyyy-mm-dd"))
    newValue = Format(newDateBilled, "yyyy-mm-dd")

    Debug.Print ""
    Debug.Print "Calling LogOrderAction..."
    Debug.Print "  Parameters:"
    Debug.Print "    SOID: " & lngSOID
    Debug.Print "    OrderNumber: " & sOrderNumber
    Debug.Print "    Action: STAMP_BILLED"
    Debug.Print "    OldStatus: " & oldValue
    Debug.Print "    NewStatus: " & newValue

    ' ---- AUDIT: Stamp success ----
    Call LogOrderAction(lngSOID, sOrderNumber, "STAMP_BILLED", _
                   oldValue, newValue, "Billed date stamped by user")

    Debug.Print "LogOrderAction returned"

    ' Refresh your formatted display textbox
    Debug.Print "Refreshing txtStampDate..."
    On Error Resume Next
    Me!txtStampDate.Requery
    On Error GoTo EH
    Debug.Print "Refresh complete"

    Debug.Print "========================================="
    Debug.Print "STAMP BILLED: SUCCESS"
    Debug.Print "========================================="
    Debug.Print ""

    Exit Sub

EH:
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "STAMP BILLED: ERROR"
    Debug.Print "  Error Number: " & Err.Number
    Debug.Print "  Error Description: " & Err.Description
    Debug.Print "  Error Source: " & Err.Source
    Debug.Print "========================================="

    ' ---- AUDIT: Stamp failure ----
    Debug.Print "Logging failure to audit..."
    On Error Resume Next  ' Don't let audit error prevent error message
    Call LogOrderAction(Nz(Me!SOID, 0), Nz(Me!OrderNumber, ""), "STAMP_BILLED_FAILED", "", "", _
                   "Err " & Err.Number & ": " & Err.Description)
    Debug.Print "Failure logged"
    On Error GoTo 0

    MsgBox "Stamp Billed failed: " & Err.Description, vbExclamation
    Debug.Print ""
End Sub


'================================================================================
' DIAGNOSTIC PROCEDURES
' Run these in the Immediate Window to diagnose issues
'================================================================================

'--------------------------------------------------------------------------------
' Test 1: Check if tables exist
'--------------------------------------------------------------------------------
Public Sub Test1_CheckTablesExist()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST 1: Checking if audit tables exist"
    Debug.Print "========================================="

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim foundOrderAudit As Boolean
    Dim foundAuditLog As Boolean

    Set db = CurrentDb

    For Each tdf In db.TableDefs
        If tdf.Name = "tblOrderAudit" Then foundOrderAudit = True
        If tdf.Name = "tblAuditLog" Then foundAuditLog = True
    Next

    Debug.Print "tblOrderAudit exists: " & foundOrderAudit
    Debug.Print "tblAuditLog exists: " & foundAuditLog

    If Not foundOrderAudit Then
        Debug.Print ""
        Debug.Print "ERROR: tblOrderAudit is MISSING!"
        Debug.Print "Run CreateAuditTables() to create it"
    End If

    If Not foundAuditLog Then
        Debug.Print ""
        Debug.Print "ERROR: tblAuditLog is MISSING!"
        Debug.Print "You may need to create this table too"
    End If

    Debug.Print "========================================="
    Debug.Print ""
End Sub


'--------------------------------------------------------------------------------
' Test 2: Check table structure
'--------------------------------------------------------------------------------
Public Sub Test2_CheckTableStructure()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST 2: Checking tblOrderAudit structure"
    Debug.Print "========================================="

    On Error GoTo EH

    Dim db As DAO.Database
    Dim fld As DAO.Field

    Set db = CurrentDb

    Debug.Print "Fields in tblOrderAudit:"
    For Each fld In db.TableDefs("tblOrderAudit").Fields
        Debug.Print "  " & fld.Name & " (Type: " & fld.Type & ")"
    Next

    Debug.Print "========================================="
    Debug.Print ""
    Exit Sub

EH:
    Debug.Print ""
    Debug.Print "ERROR: " & Err.Number & " - " & Err.Description
    If Err.Number = 3078 Then
        Debug.Print "  >> Table 'tblOrderAudit' not found!"
        Debug.Print "  >> Run CreateAuditTables() to create it"
    End If
    Debug.Print "========================================="
    Debug.Print ""
End Sub


'--------------------------------------------------------------------------------
' Test 3: Test LogOrderAction directly
'--------------------------------------------------------------------------------
Public Sub Test3_TestLogOrderAction()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST 3: Testing LogOrderAction directly"
    Debug.Print "========================================="

    On Error GoTo EH

    ' Call the function directly
    Debug.Print "Calling LogOrderAction with test data..."
    Call LogOrderAction(9999, "TEST-001", "TEST_ACTION", "OldVal", "NewVal", "Test from immediate window")

    Debug.Print ""
    Debug.Print "Checking if record was created..."

    ' Check if record was created
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblOrderAudit WHERE SOID=9999 ORDER BY ActionTimestamp DESC")

    If Not rs.EOF Then
        Debug.Print "SUCCESS: Test record found!"
        Debug.Print "  SOID: " & rs!SOID
        Debug.Print "  Action: " & rs!action
        Debug.Print "  ActionBy: " & rs!ActionBy
        Debug.Print "  OldStatus: " & rs!oldStatus
        Debug.Print "  NewStatus: " & rs!newStatus

        ' Clean up
        rs.Close
        db.Execute "DELETE FROM tblOrderAudit WHERE SOID=9999", dbFailOnError
        Debug.Print ""
        Debug.Print "Test record deleted"
    Else
        Debug.Print "ERROR: Test record NOT found in database!"
        Debug.Print "LogOrderAction may have failed silently"
        rs.Close
    End If

    Debug.Print "========================================="
    Debug.Print ""
    Exit Sub

EH:
    Debug.Print ""
    Debug.Print "ERROR: " & Err.Number & " - " & Err.Description
    Debug.Print "========================================="
    Debug.Print ""
End Sub


'--------------------------------------------------------------------------------
' Test 4: Manual insert test
'--------------------------------------------------------------------------------
Public Sub Test4_ManualInsertTest()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST 4: Manual insert test"
    Debug.Print "========================================="

    On Error GoTo EH

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Debug.Print "Opening recordset..."
    Set rs = db.OpenRecordset("tblOrderAudit", dbOpenDynaset)

    Debug.Print "Adding new record..."
    rs.AddNew
    rs!SOID = 8888
    rs!OrderNumber = "MANUAL-TEST"
    rs!action = "MANUAL_INSERT"
    rs!ActionTimestamp = Now
    rs!ActionBy = "TEST USER"
    rs!ComputerName = "TEST PC"
    rs!oldStatus = "old"
    rs!newStatus = "new"
    rs!reason = "Manual test insert"

    Debug.Print "Saving record..."
    rs.Update
    Debug.Print "SUCCESS: Record saved!"

    rs.Close

    Debug.Print ""
    Debug.Print "Verifying record..."
    Set rs = db.OpenRecordset("SELECT * FROM tblOrderAudit WHERE SOID=8888")
    If Not rs.EOF Then
        Debug.Print "VERIFIED: Record exists in table"
        Debug.Print "  SOID: " & rs!SOID
        Debug.Print "  Action: " & rs!action
    Else
        Debug.Print "ERROR: Record not found!"
    End If
    rs.Close

    ' Clean up
    db.Execute "DELETE FROM tblOrderAudit WHERE SOID=8888", dbFailOnError
    Debug.Print ""
    Debug.Print "Test record deleted"

    Debug.Print "========================================="
    Debug.Print ""
    Exit Sub

EH:
    Debug.Print ""
    Debug.Print "ERROR: " & Err.Number & " - " & Err.Description

    Select Case Err.Number
        Case 3078
            Debug.Print "  >> Table 'tblOrderAudit' not found!"
        Case 3265
            Debug.Print "  >> Field name error - check field names"
        Case 3027
            Debug.Print "  >> Database is read-only!"
        Case Else
            Debug.Print "  >> Unknown error"
    End Select

    Debug.Print "========================================="
    Debug.Print ""
End Sub


'--------------------------------------------------------------------------------
' Test 5: Check database permissions
'--------------------------------------------------------------------------------
Public Sub Test5_CheckPermissions()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST 5: Checking database permissions"
    Debug.Print "========================================="

    Dim db As DAO.Database
    Set db = CurrentDb

    Debug.Print "Database Name: " & db.Name
    Debug.Print "Database Updatable: " & db.Updatable

    ' Try to create a temp table
    On Error Resume Next
    db.Execute "CREATE TABLE TempPermTest (ID INT)", dbFailOnError

    If Err.Number = 0 Then
        Debug.Print "Write Permission: YES (can create tables)"
        db.Execute "DROP TABLE TempPermTest", dbFailOnError
    Else
        Debug.Print "Write Permission: NO or LIMITED"
        Debug.Print "  Error: " & Err.Description
    End If

    Debug.Print "========================================="
    Debug.Print ""
End Sub


'--------------------------------------------------------------------------------
' CREATE TABLE: Run if tblOrderAudit doesn't exist
'--------------------------------------------------------------------------------
Public Sub CreateAuditTables()
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "Creating tblOrderAudit..."
    Debug.Print "========================================="

    On Error GoTo EH

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index

    Set db = CurrentDb

    ' Create table
    Set tdf = db.CreateTableDef("tblOrderAudit")

    ' Add fields
    Set fld = tdf.CreateField("AuditID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    tdf.Fields.Append tdf.CreateField("SOID", dbLong)
    tdf.Fields.Append tdf.CreateField("OrderNumber", dbText, 50)
    tdf.Fields.Append tdf.CreateField("action", dbText, 50)
    tdf.Fields.Append tdf.CreateField("ActionTimestamp", dbDate)
    tdf.Fields.Append tdf.CreateField("ActionBy", dbText, 100)
    tdf.Fields.Append tdf.CreateField("ComputerName", dbText, 100)
    tdf.Fields.Append tdf.CreateField("oldStatus", dbText, 255)
    tdf.Fields.Append tdf.CreateField("newStatus", dbText, 255)
    tdf.Fields.Append tdf.CreateField("reason", dbMemo)

    Debug.Print "Fields added"

    ' Create primary key
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("AuditID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "Primary key created"

    ' Append table
    db.TableDefs.Append tdf
    Debug.Print "Table created!"

    ' Create indexes
    Set tdf = db.TableDefs("tblOrderAudit")

    Set idx = tdf.CreateIndex("idx_SOID")
    Set fld = idx.CreateField("SOID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("idx_Action")
    Set fld = idx.CreateField("action")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("idx_Timestamp")
    Set fld = idx.CreateField("ActionTimestamp")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "Indexes created"

    Debug.Print "========================================="
    Debug.Print "SUCCESS: tblOrderAudit created!"
    Debug.Print "========================================="
    Debug.Print ""

    MsgBox "tblOrderAudit created successfully!", vbInformation
    Exit Sub

EH:
    Debug.Print ""
    Debug.Print "ERROR: " & Err.Number & " - " & Err.Description

    If Err.Number = 3010 Then
        Debug.Print "Table already exists - no action needed"
        MsgBox "Table 'tblOrderAudit' already exists.", vbInformation
    Else
        MsgBox "Error creating table: " & Err.Description, vbCritical
    End If

    Debug.Print "========================================="
    Debug.Print ""
End Sub


'================================================================================
' USAGE INSTRUCTIONS
'================================================================================
'
' 1. Press Ctrl+G to open Immediate Window
'
' 2. Run diagnostic tests in order:
'    Test1_CheckTablesExist
'    Test2_CheckTableStructure
'    Test3_TestLogOrderAction
'    Test4_ManualInsertTest
'    Test5_CheckPermissions
'
' 3. If Test1 shows table missing:
'    CreateAuditTables
'
' 4. Replace LogOrderAction in basAuditLogging with debug version
'
' 5. Replace cmdStampBilled_Click in Form_frmOrderList with debug version
'
' 6. Try stamping a billed date and watch debug output
'
' 7. Review debug messages to identify the problem
'
'================================================================================
