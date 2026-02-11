# Audit Logging Troubleshooting Guide

**Issue**: STAMP actions not appearing in tblAuditLog or tblOrderAudit
**Date**: 2026-02-11

---

## Quick Diagnostic Steps

Follow these steps in order to diagnose the issue:

### Step 1: Verify Tables Exist

Run this in the Access Immediate Window (Ctrl+G):

```vba
' Check if tables exist
Debug.Print "=== TABLE CHECK ==="
Dim db As DAO.Database
Set db = CurrentDb
Dim tdf As DAO.TableDef
Dim foundOrderAudit As Boolean
Dim foundAuditLog As Boolean

For Each tdf In db.TableDefs
    If tdf.Name = "tblOrderAudit" Then foundOrderAudit = True
    If tdf.Name = "tblAuditLog" Then foundAuditLog = True
Next

Debug.Print "tblOrderAudit exists: " & foundOrderAudit
Debug.Print "tblAuditLog exists: " & foundAuditLog

If Not foundOrderAudit Then
    Debug.Print "ERROR: tblOrderAudit table is MISSING!"
End If
If Not foundAuditLog Then
    Debug.Print "ERROR: tblAuditLog table is MISSING!"
End If
```

**Expected Output**:
```
=== TABLE CHECK ===
tblOrderAudit exists: True
tblAuditLog exists: True
```

**If tables are missing**, jump to [Create Missing Tables](#create-missing-tables) section below.

---

### Step 2: Verify Table Structure

Run this to check field structure:

```vba
' Check tblOrderAudit structure
Debug.Print "=== FIELD CHECK: tblOrderAudit ==="
Set db = CurrentDb
Dim fld As DAO.Field
On Error Resume Next
For Each fld In db.TableDefs("tblOrderAudit").Fields
    Debug.Print fld.Name & " (" & TypeName(fld.Type) & ")"
Next
If Err.Number <> 0 Then
    Debug.Print "ERROR: " & Err.Description
End If
On Error GoTo 0
```

**Expected Fields**:
- SOID (Long)
- OrderNumber (Text)
- action (Text)
- ActionTimestamp (Date/Time)
- ActionBy (Text)
- ComputerName (Text)
- oldStatus (Text)
- newStatus (Text)
- reason (Text/Memo)

---

### Step 3: Test LogOrderAction Function Directly

Run this test in the Immediate Window:

```vba
' Direct test of LogOrderAction
Debug.Print "=== TESTING LogOrderAction ==="
On Error GoTo EH

Call LogOrderAction(999, "TEST-001", "TEST_ACTION", "OldVal", "NewVal", "Test from immediate window")
Debug.Print "SUCCESS: LogOrderAction executed without error"

' Check if record was created
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 * FROM tblOrderAudit WHERE SOID=999 ORDER BY ActionTimestamp DESC")
If Not rs.EOF Then
    Debug.Print "VERIFIED: Test record found in tblOrderAudit"
    Debug.Print "  SOID: " & rs!SOID
    Debug.Print "  Action: " & rs!action
    Debug.Print "  ActionBy: " & rs!ActionBy
Else
    Debug.Print "ERROR: Test record NOT found in tblOrderAudit"
End If
rs.Close

' Clean up test record
CurrentDb.Execute "DELETE FROM tblOrderAudit WHERE SOID=999", dbFailOnError
Debug.Print "Test record deleted"

Exit Sub
EH:
Debug.Print "ERROR: " & Err.Number & " - " & Err.Description
```

**Expected Output**:
```
=== TESTING LogOrderAction ===
SUCCESS: LogOrderAction executed without error
VERIFIED: Test record found in tblOrderAudit
  SOID: 999
  Action: TEST_ACTION
  ActionBy: [YourUsername]
Test record deleted
```

**If this fails**, the problem is with the LogOrderAction function itself.

---

### Step 4: Add Debug Logging to cmdStampBilled_Click

Modify your cmdStampBilled_Click to add debug output:

```vba
Private Sub cmdStampBilled_Click()
    On Error GoTo EH

    Dim oldDateBilled As Variant
    Dim newDateBilled As Variant
    Dim lngSOID As Long
    Dim sOrderNumber As String

    Debug.Print "=== STAMP BILLED START ==="

    ' Must be on an existing record
    If Me.NewRecord Then
        MsgBox "Please select or save a record before stamping.", vbExclamation
        Exit Sub
    End If

    ' Capture current state BEFORE making changes
    lngSOID = Nz(Me!SOID, 0)
    sOrderNumber = Nz(Me!OrderNumber, "")
    oldDateBilled = Me!DateBilled  ' May be Null

    Debug.Print "SOID: " & lngSOID
    Debug.Print "OrderNumber: " & sOrderNumber
    Debug.Print "Old DateBilled: " & IIf(IsNull(oldDateBilled), "(null)", oldDateBilled)

    ' Clear any old values
    On Error Resume Next
    TempVars.Remove "StampBilledDate"
    TempVars.Remove "StampBilledResult"
    On Error GoTo EH

    ' Open the dialog modally
    DoCmd.OpenForm "dlgStampBilledDate", WindowMode:=acDialog

    ' Check result
    If Nz(TempVars("StampBilledResult"), "Cancel") <> "OK" Then
        Debug.Print "User canceled stamp dialog"
        Exit Sub
    End If

    ' Get new value from dialog
    newDateBilled = TempVars("StampBilledDate")
    Debug.Print "New DateBilled: " & newDateBilled

    ' Write the chosen date to your bound field
    Me!DateBilled = newDateBilled

    ' Persist immediately
    If Me.Dirty Then Me.Dirty = False

    Debug.Print "About to call LogOrderAction..."

    ' ---- AUDIT: Stamp success ----
    Call LogOrderAction(lngSOID, sOrderNumber, "STAMP_BILLED", _
                   IIf(IsNull(oldDateBilled), "", Format(oldDateBilled, "yyyy-mm-dd")), _
                   Format(newDateBilled, "yyyy-mm-dd"), _
                   "Billed date stamped by user")

    Debug.Print "LogOrderAction called successfully"

    ' Refresh your formatted display textbox
    Me!txtStampDate.Requery

    Debug.Print "=== STAMP BILLED COMPLETE ==="
    Exit Sub

EH:
    Debug.Print "=== ERROR IN STAMP BILLED ==="
    Debug.Print "Error: " & Err.Number & " - " & Err.Description

    ' ---- AUDIT: Stamp failure ----
    On Error Resume Next  ' Don't let audit error prevent error message
    Call LogOrderAction(Nz(Me!SOID, 0), Nz(Me!OrderNumber, ""), "STAMP_BILLED_FAILED", "", "", _
                   "Err " & Err.Number & ": " & Err.Description)
    On Error GoTo 0

    MsgBox "Stamp Billed failed: " & Err.Description, vbExclamation
End Sub
```

**After adding debug logging**:
1. Open Immediate Window (Ctrl+G)
2. Clear the window
3. Run the stamp operation
4. Review debug output

**Expected Debug Output**:
```
=== STAMP BILLED START ===
SOID: 123
OrderNumber: 2025-001
Old DateBilled: (null)
New DateBilled: 2/11/2026
About to call LogOrderAction...
LogOrderAction called successfully
=== STAMP BILLED COMPLETE ===
```

---

### Step 5: Check LogOrderAction Function Code

Verify the LogOrderAction function in basAuditLogging module looks like this:

```vba
Public Sub LogOrderAction( _
    ByVal SOID As Long, _
    ByVal OrderNumber As String, _
    ByVal action As String, _
    Optional ByVal oldStatus As String = "", _
    Optional ByVal newStatus As String = "", _
    Optional ByVal reason As String = "" _
)
    ' Order-specific audit logging
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Debug.Print "=== LogOrderAction START ==="
    Debug.Print "  SOID: " & SOID
    Debug.Print "  OrderNumber: " & OrderNumber
    Debug.Print "  Action: " & action

    On Error GoTo EH

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblOrderAudit", dbOpenDynaset)

    Debug.Print "  Recordset opened successfully"

    rs.AddNew
    rs!SOID = SOID
    rs!OrderNumber = OrderNumber
    rs!action = action
    rs!ActionTimestamp = Now
    rs!ActionBy = Environ("USERNAME")
    rs!ComputerName = Environ("COMPUTERNAME")
    rs!oldStatus = oldStatus
    rs!newStatus = newStatus
    rs!reason = reason
    rs.Update

    Debug.Print "  Record saved successfully"
    Debug.Print "=== LogOrderAction COMPLETE ==="

    rs.Close
    Exit Sub

EH:
    Debug.Print "=== LogOrderAction ERROR ==="
    Debug.Print "  Error: " & Err.Number & " - " & Err.Description
    ' Silently ignore errors (don't crash the app)
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Sub
```

**Note**: The original code has `On Error Resume Next` at the top which **silently swallows all errors**. This is why you might not see failures. The version above adds debug logging.

---

## Common Issues and Solutions

### Issue 1: Tables Don't Exist

**Symptom**: Step 1 shows tables missing

**Solution**: Create the tables (see below)

---

### Issue 2: Wrong Field Names

**Symptom**: Error "Field not found" in debug output

**Solution**: Field names in the table must match exactly (case-sensitive in some contexts):
- `SOID` (not `soid` or `SOId`)
- `ActionTimestamp` (not `ActionTimeStamp` or `Timestamp`)
- `action` (not `Action`)

**Fix**: Rename fields in table design to match exactly.

---

### Issue 3: On Error Resume Next Hiding Errors

**Symptom**: No errors shown, but no records created

**Solution**: The original LogOrderAction has `On Error Resume Next` which silently ignores all errors. Replace it with the version above that has error handling with debug output.

---

### Issue 4: Wrong Table Name

**Symptom**: Error "Table 'tblOrderAudit' not found"

**Solution**:
1. Verify table name is exactly `tblOrderAudit` (case-sensitive)
2. Check for typos in LogOrderAction function
3. Verify you're connected to the correct database

---

### Issue 5: Permissions Issue

**Symptom**: Error "You do not have permission to insert into table"

**Solution**:
1. Verify you have write permissions to the backend database
2. Check if database is read-only
3. Try compacting and repairing the database

---

### Issue 6: No Debug Output

**Symptom**: Debug.Print statements don't show anything

**Solution**:
1. Make sure Immediate Window is open (Ctrl+G)
2. Verify code is actually running (add MsgBox statements)
3. Check if code compilation succeeded

---

## Create Missing Tables

If tables don't exist, run this code:

```vba
Public Sub CreateAuditTables()
    On Error GoTo EH

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index

    Set db = CurrentDb

    Debug.Print "Creating tblOrderAudit..."

    ' Create tblOrderAudit
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

    ' Create primary key
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("AuditID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    ' Append table
    db.TableDefs.Append tdf

    Debug.Print "tblOrderAudit created successfully!"

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

    Debug.Print "Indexes created successfully!"

    MsgBox "tblOrderAudit created successfully!", vbInformation
    Exit Sub

EH:
    If Err.Number = 3010 Then  ' Table already exists
        Debug.Print "Table already exists"
        MsgBox "Table 'tblOrderAudit' already exists.", vbInformation
    Else
        Debug.Print "Error: " & Err.Number & " - " & Err.Description
        MsgBox "Error creating table: " & Err.Description, vbCritical
    End If
End Sub
```

**To run**:
1. Copy code to a new module
2. Press F5 to run CreateAuditTables
3. Verify table was created
4. Try audit logging again

---

## Verification Query

After fixing the issue, run this to verify audit entries:

```sql
SELECT TOP 10
    ActionTimestamp,
    action,
    SOID,
    OrderNumber,
    oldStatus,
    newStatus,
    ActionBy,
    reason
FROM tblOrderAudit
ORDER BY ActionTimestamp DESC;
```

---

## Manual Insert Test

To verify table is writable, try a manual insert:

```vba
Sub TestManualInsert()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblOrderAudit", dbOpenDynaset)

    rs.AddNew
    rs!SOID = 9999
    rs!OrderNumber = "MANUAL-TEST"
    rs!action = "MANUAL_INSERT"
    rs!ActionTimestamp = Now
    rs!ActionBy = "TEST USER"
    rs!ComputerName = "TEST PC"
    rs!oldStatus = "old"
    rs!newStatus = "new"
    rs!reason = "Manual test insert"
    rs.Update

    rs.Close

    MsgBox "Manual insert successful! Check tblOrderAudit for SOID=9999"
End Sub
```

If this works, the table is fine and the problem is in the LogOrderAction call.

---

## Next Steps

1. Run Step 1 to verify tables exist
2. If tables missing, run CreateAuditTables
3. Run Step 3 to test LogOrderAction directly
4. If Step 3 works, add debug logging to cmdStampBilled_Click
5. Review debug output to find where it's failing
6. Apply appropriate fix from Common Issues section

---

*End of Troubleshooting Guide*
