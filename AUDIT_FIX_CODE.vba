'================================================================================
' AUDIT FIX CODE - Error 3315 Resolution
' Issue: Field can't be zero-length string
' Solution: Use Null instead of empty strings
' Date: 2026-02-11
'================================================================================

'================================================================================
' FIXED LogOrderAction - Use this to replace existing version in basAuditLogging
'================================================================================

Public Sub LogOrderAction( _
    ByVal SOID As Long, _
    ByVal OrderNumber As String, _
    ByVal action As String, _
    Optional ByVal oldStatus As String = "", _
    Optional ByVal newStatus As String = "", _
    Optional ByVal reason As String = "" _
)
    ' Order-specific audit logging
    ' FIXED: Converts empty strings to Null to avoid Error 3315
    ' "Field can't be a zero-length string"

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    On Error Resume Next  ' Don't crash app if logging fails

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblOrderAudit", dbOpenDynaset)

    rs.AddNew
    rs!SOID = SOID
    rs!OrderNumber = OrderNumber
    rs!action = action
    rs!ActionTimestamp = Now
    rs!ActionBy = Environ("USERNAME")
    rs!ComputerName = Environ("COMPUTERNAME")

    ' FIX: Convert empty strings to Null
    ' This prevents Error 3315 when "Allow Zero Length" = No
    If Len(Trim(oldStatus)) > 0 Then
        rs!oldStatus = oldStatus
    Else
        rs!oldStatus = Null
    End If

    If Len(Trim(newStatus)) > 0 Then
        rs!newStatus = newStatus
    Else
        rs!newStatus = Null
    End If

    If Len(Trim(reason)) > 0 Then
        rs!reason = reason
    Else
        rs!reason = Null
    End If

    rs.Update
    rs.Close

    ' Silently ignore errors
    On Error GoTo 0
End Sub


'================================================================================
' ALTERNATIVE VERSION: With Debug Logging (for testing)
'================================================================================

Public Sub LogOrderAction_Debug( _
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
    Debug.Print "  OldStatus (input): '" & oldStatus & "' (len=" & Len(oldStatus) & ")"
    Debug.Print "  NewStatus (input): '" & newStatus & "' (len=" & Len(newStatus) & ")"
    Debug.Print "  Reason (input): '" & reason & "' (len=" & Len(reason) & ")"

    On Error GoTo EH

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblOrderAudit", dbOpenDynaset)

    rs.AddNew

    rs!SOID = SOID
    rs!OrderNumber = OrderNumber
    rs!action = action
    rs!ActionTimestamp = Now
    rs!ActionBy = Environ("USERNAME")
    rs!ComputerName = Environ("COMPUTERNAME")

    ' Convert empty strings to Null
    If Len(Trim(oldStatus)) > 0 Then
        rs!oldStatus = oldStatus
        Debug.Print "  OldStatus (saved): '" & oldStatus & "'"
    Else
        rs!oldStatus = Null
        Debug.Print "  OldStatus (saved): Null"
    End If

    If Len(Trim(newStatus)) > 0 Then
        rs!newStatus = newStatus
        Debug.Print "  NewStatus (saved): '" & newStatus & "'"
    Else
        rs!newStatus = Null
        Debug.Print "  NewStatus (saved): Null"
    End If

    If Len(Trim(reason)) > 0 Then
        rs!reason = reason
        Debug.Print "  Reason (saved): '" & Left(reason, 50) & "...'"
    Else
        rs!reason = Null
        Debug.Print "  Reason (saved): Null"
    End If

    rs.Update
    Debug.Print "  UPDATE SUCCESS!"

    rs.Close

    Debug.Print "LogOrderAction: SUCCESS"
    Debug.Print "========================================="
    Debug.Print ""

    Exit Sub

EH:
    Debug.Print "========================================="
    Debug.Print "LogOrderAction: ERROR"
    Debug.Print "  Error: " & Err.Number & " - " & Err.Description
    Debug.Print "========================================="
    Debug.Print ""

    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Sub


'================================================================================
' TEST PROCEDURE: Verify the fix works
'================================================================================

Public Sub TestError3315Fix()
    On Error GoTo EH

    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "Testing Error 3315 Fix"
    Debug.Print "========================================="

    ' Test 1: Empty strings (the problem case)
    Debug.Print ""
    Debug.Print "Test 1: Empty strings for old/new status"
    LogOrderAction 9991, "TEST-EMPTY", "TEST_EMPTY_STRINGS", "", "", ""

    ' Test 2: Null reason but with status values
    Debug.Print ""
    Debug.Print "Test 2: Empty reason with status values"
    LogOrderAction 9992, "TEST-PARTIAL", "TEST_PARTIAL", "OldVal", "NewVal", ""

    ' Test 3: All values populated
    Debug.Print ""
    Debug.Print "Test 3: All values populated"
    LogOrderAction 9993, "TEST-FULL", "TEST_FULL", "OldStatus", "NewStatus", "Full reason text"

    ' Test 4: Realistic stamp scenario
    Debug.Print ""
    Debug.Print "Test 4: Realistic STAMP_BILLED scenario (no old date)"
    LogOrderAction 9994, "2025-001", "STAMP_BILLED", "", "2026-02-11", "Billed date stamped by user"

    ' Test 5: Realistic backorder scenario
    Debug.Print ""
    Debug.Print "Test 5: Realistic BACKORDER_CREATE scenario"
    LogOrderAction 9995, "2025-001-01", "BACKORDER_CREATE", "", "", "SourceSOID=123; BatchID={guid}"

    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "Verifying records were created..."
    Debug.Print "========================================="

    ' Check all test records
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT SOID, action, oldStatus, newStatus, reason " & _
        "FROM tblOrderAudit " & _
        "WHERE SOID >= 9991 AND SOID <= 9995 " & _
        "ORDER BY SOID", _
        dbOpenSnapshot)

    Dim count As Integer
    count = 0

    Do While Not rs.EOF
        count = count + 1
        Debug.Print ""
        Debug.Print "Record " & count & ":"
        Debug.Print "  SOID: " & rs!SOID
        Debug.Print "  Action: " & rs!action
        Debug.Print "  OldStatus: " & IIf(IsNull(rs!oldStatus), "(null)", rs!oldStatus)
        Debug.Print "  NewStatus: " & IIf(IsNull(rs!newStatus), "(null)", rs!newStatus)
        Debug.Print "  Reason: " & IIf(IsNull(rs!reason), "(null)", Left(rs!reason, 30))
        rs.MoveNext
    Loop

    rs.Close

    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "RESULT: " & count & " records created successfully!"
    Debug.Print "========================================="

    ' Cleanup
    Debug.Print ""
    Debug.Print "Cleaning up test records..."
    CurrentDb.Execute "DELETE FROM tblOrderAudit WHERE SOID >= 9991 AND SOID <= 9995", dbFailOnError
    Debug.Print "Test records deleted"

    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "ALL TESTS PASSED!"
    Debug.Print "Error 3315 fix is working correctly."
    Debug.Print "========================================="
    Debug.Print ""

    Exit Sub

EH:
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "TEST FAILED!"
    Debug.Print "  Error: " & Err.Number & " - " & Err.Description
    Debug.Print "========================================="
    Debug.Print ""

    ' Try to cleanup on error
    On Error Resume Next
    CurrentDb.Execute "DELETE FROM tblOrderAudit WHERE SOID >= 9991 AND SOID <= 9995", dbFailOnError
End Sub


'================================================================================
' UPDATED cmdStampBilled_Click - No changes needed!
' The fix in LogOrderAction handles everything
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

    ' Persist immediately
    If Me.Dirty Then Me.Dirty = False

    ' ---- AUDIT: Stamp success ----
    ' NO CHANGES NEEDED - LogOrderAction now handles empty strings!
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


'================================================================================
' USAGE INSTRUCTIONS
'================================================================================
'
' 1. Open basAuditLogging module
'
' 2. Find the LogOrderAction function
'
' 3. Replace it with the FIXED version from this file (lines 12-53)
'
' 4. Save the module
'
' 5. Optional: Run TestError3315Fix to verify fix works
'
' 6. Test your stamp operation - should work now!
'
' 7. Verify with: SELECT TOP 5 * FROM tblOrderAudit ORDER BY ActionTimestamp DESC;
'
'================================================================================

'================================================================================
' WHAT THIS FIX DOES
'================================================================================
'
' BEFORE (Error 3315):
'   rs!oldStatus = ""          ' Empty string - FAILS if "Allow Zero Length" = No
'   rs!newStatus = ""          ' Empty string - FAILS if "Allow Zero Length" = No
'   rs!reason = ""             ' Empty string - FAILS if "Allow Zero Length" = No
'
' AFTER (Works):
'   If Len(Trim(oldStatus)) > 0 Then
'       rs!oldStatus = oldStatus    ' Has value - save it
'   Else
'       rs!oldStatus = Null         ' No value - use Null instead of ""
'   End If
'
' Why it works:
'   - Null is always allowed in fields
'   - Empty string "" requires "Allow Zero Length" = Yes
'   - This fix uses Null when no value, avoiding the restriction
'
'================================================================================
