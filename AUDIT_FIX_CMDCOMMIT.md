# Fix cmdCommit Audit Logging - Error 3315

## Problem

The `cmdCommit_Click()` button in `Form_frmNumberRes` has audit logging implemented, but it's failing silently due to Error 3315 (Field can't be a zero-length string).

**Current Code (Line 2472):**
```vba
LogOrderAction 0, "", "BATCH_COMMIT", "", "", auditReason
                   ^^                ^^  ^^
                   These cause Error 3315!
```

## Root Cause

The LogOrderAction call is passing empty strings (`""`) for:
- `PONum` parameter (2nd parameter)
- `OldVal` parameter (4th parameter)
- `NewVal` parameter (5th parameter)

These empty strings violate the database constraint that these fields cannot be zero-length strings.

## Solution

Replace all empty strings (`""`) with `Null` to properly indicate "no value".

## Fixed Code

### Location
- **Form**: `Form_frmNumberRes`
- **Event**: `cmdCommit_Click()`
- **Line**: ~2472

### Before (BROKEN)
```vba
' SOID=0 because this is a batch-level event (multiple orders)
LogOrderAction 0, "", "BATCH_COMMIT", "", "", auditReason
```

### After (FIXED)
```vba
' SOID=0 because this is a batch-level event (multiple orders)
LogOrderAction 0, Null, "BATCH_COMMIT", Null, Null, auditReason
```

## Complete Fixed cmdCommit_Click() Function

```vba
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

        ' *** FIXED: Use Null instead of empty strings ***
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
```

## What Changed

**Only one line changed:**

```diff
- LogOrderAction 0, "", "BATCH_COMMIT", "", "", auditReason
+ LogOrderAction 0, Null, "BATCH_COMMIT", Null, Null, auditReason
```

## Implementation Steps

1. **Open** `Form_frmNumberRes` in Design View
2. **Find** the `cmdCommit` button
3. **Click** "View Code" or right-click → "Build Event"
4. **Locate** the line around 2472 that reads:
   ```vba
   LogOrderAction 0, "", "BATCH_COMMIT", "", "", auditReason
   ```
5. **Replace** with:
   ```vba
   LogOrderAction 0, Null, "BATCH_COMMIT", Null, Null, auditReason
   ```
6. **Save** the form
7. **Test** by clicking the Commit button

## Testing

### Test the Fix

1. **Open** the Number Reservation form
2. **Fill in** batch details
3. **Click** the Commit button
4. **Check** the audit log:

```sql
SELECT TOP 5
    EntryID,
    ActionType,
    SOID,
    PONum,
    OldVal,
    NewVal,
    Reason,
    ActionAt,
    ActionBy
FROM tblOrderActionLog
WHERE ActionType = 'BATCH_COMMIT'
ORDER BY ActionAt DESC;
```

### Expected Result

You should see a record like:
```
ActionType    : BATCH_COMMIT
SOID          : 0
PONum         : NULL
OldVal        : NULL
NewVal        : NULL
Reason        : BatchID=...; CreatedCount=...; Scope=...; etc.
ActionAt      : [timestamp]
ActionBy      : [your username]
```

## Why This Matters

This fix ensures that:
- ✅ Batch commit operations are properly audited
- ✅ You have a complete audit trail of all business events
- ✅ No silent failures in audit logging
- ✅ Consistent use of NULL for empty values across all audit calls

## Related Files

- **Original Fix**: `AUDIT_FIX_ERROR_3315.md`
- **Fixed LogOrderAction**: `AUDIT_FIX_CODE.vba`
- **Implementation Guide**: `AUDIT_IMPLEMENTATION_PLAN.md`
