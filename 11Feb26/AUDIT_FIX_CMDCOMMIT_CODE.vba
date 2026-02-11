' ============================================================================
' FIX: cmdCommit Audit Logging - Error 3315
' ============================================================================
' Form: Form_frmNumberRes
' Event: cmdCommit_Click()
' Issue: Empty strings ("") cause Error 3315
' Fix: Replace empty strings with Null
' ============================================================================

' --- OPTION 1: Just the One Line Fix ---
' Find this line (around line 2472):
'   LogOrderAction 0, "", "BATCH_COMMIT", "", "", auditReason
'
' Replace with this:
    LogOrderAction 0, Null, "BATCH_COMMIT", Null, Null, auditReason


' --- OPTION 2: Complete Context (Lines 2459-2473) ---
' Replace the entire audit logging section:

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

        ' *** FIXED: Use Null instead of empty strings to avoid Error 3315 ***
        ' SOID=0 because this is a batch-level event (multiple orders)
        LogOrderAction 0, Null, "BATCH_COMMIT", Null, Null, auditReason


' --- OPTION 3: Complete cmdCommit_Click() Function ---
' If you prefer to replace the entire function:

Private Sub cmdCommit_Click()
    On Error GoTo EH

    Trace "Dialog.Commit: start"
    Err.Clear

    Dim Scope As String, baseT As String, sysLet As String
    If Not ValidateHeader(Scope, baseT, sysLet) Then Exit Sub
    If Not HasAnyQty() Then
        MsgBox "Enter at least one quantity > 0.", vbExclamation
        Exit Sub
    End If

    Dim created As Long, BatchID As String
    Dim custCode As Variant, custName As Variant, poNum As Variant, dtRecv As Variant
    custCode = Nz(Me.cboCustomerCode, Null)
    custName = Nz(Me.txtCustomerName, Null)
    poNum = Nz(Me.txtPONumber, Null)
    dtRecv = Nz(Me.txtDateReceived, Null)

    ' Use your existing CommitBatch signature exactly as-is
    If CommitBatch(Scope, baseT, sysLet, custCode, custName, poNum, dtRecv, created, BatchID) Then
        ' Report success via TempVars (caller will decide what to do)
        On Error Resume Next
        TempVars("BatchResult") = "Committed"
        TempVars("BatchID") = BatchID
        TempVars("CreatedCount") = CStr(created)
        TempVars.Remove "BatchErr"
        On Error GoTo 0

        Trace "Dialog.Commit: success"
        Err.Clear                 ' <-- prevent stale Err from bubbling to the launcher

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

        ' *** FIXED: Use Null instead of empty strings to avoid Error 3315 ***
        ' SOID=0 because this is a batch-level event (multiple orders)
        LogOrderAction 0, Null, "BATCH_COMMIT", Null, Null, auditReason

        DoCmd.Close acForm, Me.name, acSaveNo
        Exit Sub
    Else
        ' False without Err.Number—treat as error and surface a concise message
        On Error Resume Next
        Dim errMsg As String
        errMsg = Nz(TempVars("BatchErr"), "")
        TempVars.Remove "BatchErr"
        On Error GoTo 0

        If Len(errMsg) > 0 Then
            MsgBox "Could not commit batch: " & errMsg, vbCritical
        Else
            MsgBox "Could not commit batch (unknown reason).", vbCritical
        End If

        Trace "Dialog.Commit: failed"
        Exit Sub
    End If

EH:
    MsgBox "Error committing batch: " & Err.Description, vbCritical
    Trace "Dialog.Commit: error #" & Err.Number & " - " & Err.Description
End Sub


' ============================================================================
' TESTING THE FIX
' ============================================================================

' After applying the fix, test with this SQL query:

Sub TestCmdCommitAudit()
    ' Run this query to verify BATCH_COMMIT events are being logged:

    ' SQL:
    ' SELECT TOP 10
    '     EntryID,
    '     ActionType,
    '     SOID,
    '     PONum,
    '     OldVal,
    '     NewVal,
    '     Reason,
    '     ActionAt,
    '     ActionBy
    ' FROM tblOrderActionLog
    ' WHERE ActionType = 'BATCH_COMMIT'
    ' ORDER BY ActionAt DESC;

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT TOP 10 * FROM tblOrderActionLog " & _
        "WHERE ActionType = 'BATCH_COMMIT' " & _
        "ORDER BY ActionAt DESC", dbOpenSnapshot)

    Debug.Print "=== Recent BATCH_COMMIT Audit Entries ==="
    Debug.Print ""

    If rs.EOF Then
        Debug.Print "NO BATCH_COMMIT entries found!"
        Debug.Print "This means the audit logging is not working."
    Else
        Do While Not rs.EOF
            Debug.Print "EntryID: " & rs!EntryID
            Debug.Print "  SOID: " & Nz(rs!SOID, "(null)")
            Debug.Print "  PONum: " & Nz(rs!PONum, "(null)")
            Debug.Print "  OldVal: " & Nz(rs!OldVal, "(null)")
            Debug.Print "  NewVal: " & Nz(rs!NewVal, "(null)")
            Debug.Print "  Reason: " & Left(Nz(rs!Reason, ""), 100)
            Debug.Print "  ActionAt: " & rs!ActionAt
            Debug.Print "  ActionBy: " & rs!ActionBy
            Debug.Print ""
            rs.MoveNext
        Loop
    End If

    rs.Close
    Set rs = Nothing

    Debug.Print "=== End of Report ==="
End Sub


' ============================================================================
' SUMMARY OF THE FIX
' ============================================================================

' BEFORE (BROKEN):
'   LogOrderAction 0, "", "BATCH_COMMIT", "", "", auditReason
'                      ^^                ^^  ^^
'                      These cause Error 3315!
'
' AFTER (FIXED):
'   LogOrderAction 0, Null, "BATCH_COMMIT", Null, Null, auditReason
'                      ^^^^                ^^^^  ^^^^
'                      These work properly!
'
' WHY IT MATTERS:
' - Empty strings ("") violate database constraints
' - Null properly indicates "no value"
' - Same fix as cmdStampBilled (see AUDIT_FIX_ERROR_3315.md)
' - Ensures complete audit trail for batch operations
