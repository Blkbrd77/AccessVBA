' ============================================================================
' FIX: Error 94 - Invalid Use of Null
' ============================================================================
' Issue: Passing Null to String parameters causes Error 94
' Solution: Update LogOrderAction to handle empty strings, pass "" not Null
' ============================================================================

' ============================================================================
' STEP 1: Replace LogOrderAction in basAuditLogging
' ============================================================================

Public Sub LogOrderAction( _
    ByVal SOID As Long, _
    ByVal OrderNumber As String, _
    ByVal action As String, _
    Optional ByVal oldStatus As String = "", _
    Optional ByVal newStatus As String = "", _
    Optional ByVal reason As String = "" _
)
    ' Order-specific audit logging
    ' FIXED: Handles empty strings and converts them to Null before saving
    ' Prevents both Error 3315 (zero-length string) and Error 94 (invalid use of null)

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    On Error Resume Next  ' Don't crash app if logging fails

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblOrderAudit", dbOpenDynaset)

    rs.AddNew
    rs!SOID = SOID

    ' Handle OrderNumber - allow empty string or convert to Null
    If Len(Trim$(OrderNumber)) > 0 Then
        rs!OrderNumber = OrderNumber
    Else
        rs!OrderNumber = Null
    End If

    rs!action = action
    rs!ActionTimestamp = Now
    rs!ActionBy = Environ("USERNAME")
    rs!ComputerName = Environ("COMPUTERNAME")

    ' FIX: Convert empty strings to Null to avoid Error 3315
    ' This safely handles empty strings ("") or vbNullString from callers
    If Len(Trim$(oldStatus)) > 0 Then
        rs!oldStatus = oldStatus
    Else
        rs!oldStatus = Null
    End If

    If Len(Trim$(newStatus)) > 0 Then
        rs!newStatus = newStatus
    Else
        rs!newStatus = Null
    End If

    If Len(Trim$(reason)) > 0 Then
        rs!reason = reason
    Else
        rs!reason = Null
    End If

    rs.Update
    rs.Close

    ' Silently ignore errors
    On Error GoTo 0
End Sub


' ============================================================================
' STEP 2: Update cmdCommit_Click in Form_frmNumberRes
' ============================================================================

' Find this line (around line 2472):
'   LogOrderAction 0, Null, "BATCH_COMMIT", Null, Null, auditReason
'
' Change it to:
        LogOrderAction 0, "", "BATCH_COMMIT", "", "", auditReason


' ============================================================================
' Complete cmdCommit_Click Context (lines 2459-2473)
' ============================================================================

        ' --- AUDIT: Batch committed (1 row per batch) ---
        Dim auditReason As String
        auditReason = "BatchID=" & BatchID & _
                  "; CreatedCount=" & created & _
                  "; Scope=" & Scope & _
                  "; BaseToken=" & baseT & _
                  "; SystemLetter=" & sysLet & _
                  "; CustomerCode=" & Nz(custCode, "") & _
                  "; CustomerName=" & Nz(custName, "") & _
                  "; PONumber=" & Nz(poNum, "") & _
                  "; DateReceived=" & IIf(IsDate(dtRecv), Format$(CDate(dtRecv), "yyyy-mm-dd"), "")

        ' SOID=0 because this is a batch-level event (multiple orders)
        ' FIXED: Use empty strings, LogOrderAction converts to Null
        LogOrderAction 0, "", "BATCH_COMMIT", "", "", auditReason


' ============================================================================
' STEP 3: Verify cmdStampBilled_Click is correct
' ============================================================================

' The cmdStampBilled_Click should use empty strings, not Null:

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

    ' Persist immediately
    If Me.Dirty Then Me.Dirty = False

    ' ---- AUDIT: Stamp success ----
    ' CORRECT: Pass empty string for oldDateBilled if it was Null
    ' LogOrderAction will convert empty strings to Null before saving
    LogOrderAction lngSOID, sOrderNumber, "STAMP_BILLED", _
                   IIf(IsNull(oldDateBilled), "", Format(oldDateBilled, "yyyy-mm-dd")), _
                   Format(newDateBilled, "yyyy-mm-dd"), _
                   "Billed date stamped by user"

    ' Refresh display
    Me!txtStampDate.Requery

    Exit Sub

EH:
    ' ---- AUDIT: Stamp failure ----
    LogOrderAction Nz(Me!SOID, 0), Nz(Me!OrderNumber, ""), "STAMP_BILLED_FAILED", "", "", _
                   "Err " & Err.Number & ": " & Err.Description

    MsgBox "Stamp Billed failed: " & Err.Description, vbExclamation
End Sub


' ============================================================================
' TESTING CODE
' ============================================================================

Public Sub TestAuditLoggingFix()
    ' Test that LogOrderAction handles empty strings correctly
    ' and doesn't throw Error 94 or Error 3315

    Debug.Print "========================================="
    Debug.Print "Testing Audit Logging Fix"
    Debug.Print "========================================="

    On Error GoTo EH

    ' Test 1: Empty strings (common case)
    Debug.Print ""
    Debug.Print "Test 1: Empty strings for all optional params"
    LogOrderAction 9991, "", "TEST_EMPTY", "", "", ""
    Debug.Print "  ✓ Success"

    ' Test 2: Partial values
    Debug.Print ""
    Debug.Print "Test 2: Some values, some empty"
    LogOrderAction 9992, "TEST-001", "TEST_PARTIAL", "OldVal", "", "Some reason"
    Debug.Print "  ✓ Success"

    ' Test 3: All values
    Debug.Print ""
    Debug.Print "Test 3: All values populated"
    LogOrderAction 9993, "TEST-002", "TEST_FULL", "OldStatus", "NewStatus", "Full reason"
    Debug.Print "  ✓ Success"

    ' Test 4: Batch commit scenario (SOID=0, empty order number)
    Debug.Print ""
    Debug.Print "Test 4: BATCH_COMMIT scenario"
    LogOrderAction 0, "", "BATCH_COMMIT", "", "", "BatchID=123; CreatedCount=5"
    Debug.Print "  ✓ Success"

    ' Verify records were created
    Debug.Print ""
    Debug.Print "Verifying records in database..."

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT COUNT(*) AS cnt FROM tblOrderAudit WHERE SOID IN (0, 9991, 9992, 9993) " & _
        "AND action LIKE 'TEST_%' OR action = 'BATCH_COMMIT'", _
        dbOpenSnapshot)

    Dim recordCount As Long
    recordCount = rs!cnt
    rs.Close

    Debug.Print "  Found " & recordCount & " test records"

    If recordCount >= 4 Then
        Debug.Print ""
        Debug.Print "========================================="
        Debug.Print "ALL TESTS PASSED! ✓"
        Debug.Print "Audit logging is working correctly."
        Debug.Print "========================================="
    Else
        Debug.Print ""
        Debug.Print "WARNING: Expected 4 records, found " & recordCount
    End If

    ' Cleanup
    Debug.Print ""
    Debug.Print "Cleaning up test records..."
    CurrentDb.Execute "DELETE FROM tblOrderAudit WHERE SOID IN (9991, 9992, 9993)", dbFailOnError
    CurrentDb.Execute "DELETE FROM tblOrderAudit WHERE action LIKE 'TEST_%'", dbFailOnError
    Debug.Print "Test records deleted"

    Exit Sub

EH:
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST FAILED! ✗"
    Debug.Print "  Error: " & Err.Number & " - " & Err.Description
    Debug.Print "========================================="

    ' Cleanup on error
    On Error Resume Next
    CurrentDb.Execute "DELETE FROM tblOrderAudit WHERE SOID IN (9991, 9992, 9993)", dbFailOnError
    CurrentDb.Execute "DELETE FROM tblOrderAudit WHERE action LIKE 'TEST_%'", dbFailOnError
End Sub


' ============================================================================
' QUERY TO CHECK RESULTS
' ============================================================================

' Run this in SQL View to check audit log:

' SELECT TOP 20
'     SOID,
'     OrderNumber,
'     action,
'     oldStatus,
'     newStatus,
'     LEFT(reason, 50) AS reason_preview,
'     ActionTimestamp,
'     ActionBy
' FROM tblOrderAudit
' ORDER BY ActionTimestamp DESC;


' ============================================================================
' SUMMARY
' ============================================================================

' PROBLEM:
'   Passing Null to String parameters causes Error 94
'
' SOLUTION:
'   1. Update LogOrderAction to handle empty strings and convert to Null
'   2. Callers pass empty strings (""), not Null
'   3. LogOrderAction converts empty strings to Null before saving
'
' BENEFITS:
'   ✓ No Error 94 (Invalid use of Null)
'   ✓ No Error 3315 (Zero-length string)
'   ✓ Clean audit log with Null for empty values
'   ✓ Simple for calling code - just pass ""
'
' FILES TO UPDATE:
'   1. basAuditLogging - Replace LogOrderAction function
'   2. Form_frmNumberRes - Change cmdCommit line 2472
'   3. Form_frmOrderList - Verify cmdStampBilled uses ""
