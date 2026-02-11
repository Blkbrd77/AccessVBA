'================================================================================
' AUDIT IMPLEMENTATION CODE
' Purpose: Enhanced audit logging for STAMP operations
' Date: 2026-02-11
'
' Instructions:
' 1. Open Form_frmOrderList in Design View
' 2. Open VBA Editor (Alt+F11)
' 3. Find the cmdStampBilled_Click() subroutine
' 4. Replace the existing code with the ENHANCED version below
'================================================================================

'================================================================================
' ENHANCED cmdStampBilled_Click() - WITH AUDIT LOGGING
' Location: Form_frmOrderList
'================================================================================

Private Sub cmdStampBilled_Click()
    On Error GoTo EH

    Dim oldDateBilled As Variant
    Dim newDateBilled As Variant
    Dim lngSOID As Long
    Dim sOrderNumber As String

    ' Must be on an existing record
    If Me.NewRecord Then
        MsgBox "Please select or save a record before stamping.", vbExclamation
        Exit Sub
    End If

    ' Capture current state BEFORE making changes
    lngSOID = Nz(Me!SOID, 0)
    sOrderNumber = Nz(Me!OrderNumber, "")
    oldDateBilled = Me!DateBilled  ' May be Null

    ' Clear any old values
    On Error Resume Next
    TempVars.Remove "StampBilledDate"
    TempVars.Remove "StampBilledResult"
    On Error GoTo EH

    ' Open the dialog modally
    DoCmd.OpenForm "dlgStampBilledDate", WindowMode:=acDialog

    ' Check result
    If Nz(TempVars("StampBilledResult"), "Cancel") <> "OK" Then
        Exit Sub
    End If

    ' Get new value from dialog
    newDateBilled = TempVars("StampBilledDate")

    ' Write the chosen date to your bound field
    Me!DateBilled = newDateBilled

    ' Persist immediately (recommended to ensure audit is accurate)
    If Me.Dirty Then Me.Dirty = False

    ' ---- AUDIT: Stamp success ----
    LogOrderAction lngSOID, sOrderNumber, "STAMP_BILLED", _
                   IIf(IsNull(oldDateBilled), "", Format(oldDateBilled, "yyyy-mm-dd")), _
                   Format(newDateBilled, "yyyy-mm-dd"), _
                   "Billed date stamped by user"

    ' Refresh your formatted display textbox (if it is an expression based on DateBilled)
    Me!txtStampDate.Requery

    Exit Sub

EH:
    ' ---- AUDIT: Stamp failure ----
    LogOrderAction Nz(Me!SOID, 0), Nz(Me!OrderNumber, ""), "STAMP_BILLED_FAILED", "", "", _
                   "Err " & Err.Number & ": " & Err.Description

    MsgBox "Stamp Billed failed: " & Err.Description, vbExclamation
End Sub


'================================================================================
' ORIGINAL cmdStampBilled_Click() - FOR REFERENCE
' (This is the version being REPLACED)
'================================================================================
'
'Private Sub cmdStampBilled_Click()
'    On Error GoTo EH
'
'    ' Must be on an existing record
'    If Me.NewRecord Then
'        MsgBox "Please select or save a record before stamping.", vbExclamation
'        Exit Sub
'    End If
'
'    ' Clear any old values
'    On Error Resume Next
'    TempVars.Remove "StampBilledDate"
'    TempVars.Remove "StampBilledResult"
'    On Error GoTo 0
'
'    ' Open the dialog modally
'    DoCmd.OpenForm "dlgStampBilledDate", WindowMode:=acDialog
'
'    ' Check result
'    If Nz(TempVars("StampBilledResult"), "Cancel") <> "OK" Then
'        Exit Sub
'    End If
'
'    ' Write the chosen date to your bound field
'    Me!DateBilled = TempVars("StampBilledDate")
'
'    ' Persist immediately (optional but recommended)
'    If Me.Dirty Then Me.Dirty = False
'
'    ' Refresh your formatted display textbox (if it is an expression based on DateBilled)
'    Me!txtStampDate.Requery
'
'    Exit Sub
'
'EH:
'    MsgBox "Stamp Billed failed: " & Err.Description, vbExclamation
'End Sub


'================================================================================
' VERIFICATION QUERIES
' Run these in Access to verify audit logging is working
'================================================================================

'--------------------------------------------------------------------------------
' Query 1: View Recent Audit Activity
'--------------------------------------------------------------------------------
' SELECT ActionTimestamp, action, SOID, OrderNumber,
'        oldStatus, newStatus, ActionBy, reason
' FROM tblOrderAudit
' ORDER BY ActionTimestamp DESC;

'--------------------------------------------------------------------------------
' Query 2: View All STAMP Operations
'--------------------------------------------------------------------------------
' SELECT ActionTimestamp, SOID, OrderNumber,
'        oldStatus AS OldDate, newStatus AS NewDate,
'        ActionBy, reason
' FROM tblOrderAudit
' WHERE action = 'STAMP_BILLED'
' ORDER BY ActionTimestamp DESC;

'--------------------------------------------------------------------------------
' Query 3: View Order History (All Actions on One Order)
'--------------------------------------------------------------------------------
' SELECT ActionTimestamp, action, oldStatus, newStatus, ActionBy, reason
' FROM tblOrderAudit
' WHERE SOID = 123  -- Replace with actual SOID
' ORDER BY ActionTimestamp;

'--------------------------------------------------------------------------------
' Query 4: Operations Summary by User
'--------------------------------------------------------------------------------
' SELECT ActionBy, action, COUNT(*) AS OperationCount
' FROM tblOrderAudit
' GROUP BY ActionBy, action
' ORDER BY ActionBy, action;

'--------------------------------------------------------------------------------
' Query 5: Failed Operations
'--------------------------------------------------------------------------------
' SELECT ActionTimestamp, action, SOID, OrderNumber, reason
' FROM tblOrderAudit
' WHERE action LIKE '%_FAILED'
' ORDER BY ActionTimestamp DESC;


'================================================================================
' TEST PLAN
'================================================================================

' Test Case 1: Successful STAMP Operation
' ----------------------------------------
' 1. Open Form_frmOrderList
' 2. Select an order that has NO DateBilled value
' 3. Click "Stamp Billed" button
' 4. Enter a date (e.g., 2/11/2026)
' 5. Click OK
' 6. Run Query 2 above
' 7. VERIFY: New audit entry with action='STAMP_BILLED', oldStatus='', newStatus='2026-02-11'

' Test Case 2: Re-STAMP Operation (Change Existing Date)
' -------------------------------------------------------
' 1. Open Form_frmOrderList
' 2. Select an order that ALREADY HAS DateBilled value
' 3. Note the current DateBilled value
' 4. Click "Stamp Billed" button
' 5. Enter a different date
' 6. Click OK
' 7. Run Query 2 above
' 8. VERIFY: New audit entry with oldStatus=original date, newStatus=new date

' Test Case 3: Cancel STAMP Dialog
' ---------------------------------
' 1. Open Form_frmOrderList
' 2. Select an order
' 3. Click "Stamp Billed" button
' 4. Click Cancel in the dialog
' 5. Run Query 2 above
' 6. VERIFY: NO new audit entry created (operation was canceled)

' Test Case 4: Error Scenario
' ----------------------------
' 1. Temporarily rename tblOrderAudit to simulate error
' 2. Try to stamp billed date
' 3. Operation should still succeed (audit uses On Error Resume Next)
' 4. Run Query 5 above
' 5. VERIFY: Error is handled gracefully, operation completes
' 6. Rename tblOrderAudit back to original name

' Test Case 5: Verify All Operations
' -----------------------------------
' 1. Perform one of each operation:
'    - Cancel an order
'    - Create a backorder
'    - Stamp billed date
'    - Batch commit new orders
' 2. Run Query 1 above
' 3. VERIFY: All four operations appear in audit log


'================================================================================
' AUDIT TABLE STRUCTURE (For Reference)
' If tables don't exist, create them with these structures
'================================================================================

'--------------------------------------------------------------------------------
' tblOrderAudit - Order-specific business events
'--------------------------------------------------------------------------------
' CREATE TABLE tblOrderAudit (
'     AuditID          AUTOINCREMENT PRIMARY KEY,
'     SOID             LONG,
'     OrderNumber      TEXT(50),
'     action           TEXT(50),      -- CANCEL, BACKORDER_CREATE, STAMP_BILLED, BATCH_COMMIT
'     ActionTimestamp  DATETIME,
'     ActionBy         TEXT(100),     -- Environ("USERNAME")
'     ComputerName     TEXT(100),     -- Environ("COMPUTERNAME")
'     oldStatus        TEXT(255),     -- Previous value
'     newStatus        TEXT(255),     -- New value
'     reason           MEMO           -- Details, error messages, or notes
' );
'
' CREATE INDEX idx_OrderAudit_SOID ON tblOrderAudit(SOID);
' CREATE INDEX idx_OrderAudit_Action ON tblOrderAudit(action);
' CREATE INDEX idx_OrderAudit_Timestamp ON tblOrderAudit(ActionTimestamp);

'--------------------------------------------------------------------------------
' tblAuditLog - General purpose audit trail
'--------------------------------------------------------------------------------
' CREATE TABLE tblAuditLog (
'     AuditID          AUTOINCREMENT PRIMARY KEY,
'     AuditTimestamp   DATETIME,
'     tableName        TEXT(100),
'     recordID         LONG,
'     action           TEXT(50),
'     fieldName        TEXT(100),
'     oldValue         TEXT(255),
'     newValue         TEXT(255),
'     UserName         TEXT(100),     -- Environ("USERNAME")
'     ComputerName     TEXT(100),     -- Environ("COMPUTERNAME")
'     notes            MEMO
' );
'
' CREATE INDEX idx_AuditLog_Table ON tblAuditLog(tableName);
' CREATE INDEX idx_AuditLog_RecordID ON tblAuditLog(recordID);
' CREATE INDEX idx_AuditLog_Timestamp ON tblAuditLog(AuditTimestamp);


'================================================================================
' TROUBLESHOOTING
'================================================================================

' Problem: "Object doesn't support this property or method" error
' Solution: Verify basAuditLogging module exists and LogOrderAction function is defined

' Problem: "Invalid use of Null" error
' Solution: Check that SOID and OrderNumber are not Null before calling LogOrderAction

' Problem: No audit entries appearing
' Solution:
'   1. Verify tblOrderAudit table exists
'   2. Check field names match exactly (case-sensitive in SQL)
'   3. Verify you have write permissions to the table
'   4. Check that LogOrderAction uses "On Error Resume Next" (errors are silent)

' Problem: Audit timestamp is wrong
' Solution: Ensure computer clock is set correctly, ActionTimestamp = Now

' Problem: ActionBy shows wrong user
' Solution: Environ("USERNAME") returns Windows username, verify with Debug.Print Environ("USERNAME")


'================================================================================
' COMPLIANCE & REPORTING
'================================================================================

' For compliance and audit trail reporting, create these saved queries:

' Query: rpt_AuditTrail_Daily
' SELECT Format(ActionTimestamp, "Short Date") AS AuditDate,
'        action AS Operation,
'        COUNT(*) AS Count
' FROM tblOrderAudit
' WHERE ActionTimestamp >= Date()-30
' GROUP BY Format(ActionTimestamp, "Short Date"), action
' ORDER BY Format(ActionTimestamp, "Short Date") DESC, action;

' Query: rpt_AuditTrail_UserActivity
' SELECT ActionBy AS UserName,
'        action AS Operation,
'        COUNT(*) AS OperationCount,
'        MIN(ActionTimestamp) AS FirstOperation,
'        MAX(ActionTimestamp) AS LastOperation
' FROM tblOrderAudit
' GROUP BY ActionBy, action
' ORDER BY ActionBy, action;

' Query: rpt_OrderLifecycle
' SELECT SOID, OrderNumber,
'        ActionTimestamp,
'        action,
'        oldStatus AS PreviousState,
'        newStatus AS NewState,
'        ActionBy
' FROM tblOrderAudit
' WHERE SOID IN (SELECT SOID FROM tblOrderAudit GROUP BY SOID HAVING COUNT(*) > 1)
' ORDER BY SOID, ActionTimestamp;


'================================================================================
' End of Audit Implementation Code
'================================================================================
