# Fix Error 94: Invalid Use of Null - Complete Solution

## The Problem

After changing cmdCommit to pass `Null` instead of empty strings:
```vba
LogOrderAction 0, Null, "BATCH_COMMIT", Null, Null, auditReason
```

You're getting **Runtime Error 94: Invalid use of Null** because the `LogOrderAction` function in your database isn't handling Null parameters properly.

## Root Cause

The current LogOrderAction in your database (basAuditLogging module) looks like this:

```vba
' CURRENT VERSION (NOT HANDLING NULL)
Public Sub LogOrderAction( _
    ByVal SOID As Long, _
    ByVal OrderNumber As String, _
    ByVal action As String, _
    Optional ByVal oldStatus As String = "", _
    Optional ByVal newStatus As String = "", _
    Optional ByVal reason As String = "" _
)
    ' ... code ...
    rs!oldStatus = oldStatus    ' <-- ERROR 94 if oldStatus is Null
    rs!newStatus = newStatus    ' <-- ERROR 94 if newStatus is Null
    rs!reason = reason          ' <-- ERROR 94 if reason is Null
    ' ... code ...
End Sub
```

When you pass `Null` values, VBA can't handle them in String parameters, causing Error 94.

## Two-Part Solution

### Part 1: Fix LogOrderAction to Handle Nulls AND Empty Strings

### Part 2: Update cmdCommit to Pass Safe Values

## SOLUTION: Use vbNullString Instead of Null

The **best solution** is to:
1. Keep LogOrderAction parameters as Strings (not Variant)
2. Pass `vbNullString` instead of `Null`
3. Update LogOrderAction to convert empty strings to Null before saving

### Why This Works
- `vbNullString` is a zero-length string constant that's safe to pass to String parameters
- It won't cause Error 94
- LogOrderAction can check if it's empty and convert to Null before saving to database

## COMPLETE FIX

### Step 1: Update LogOrderAction in basAuditLogging

Replace the entire LogOrderAction function with this version:

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
    ' FIXED: Handles both empty strings and converts them to Null
    ' Prevents both Error 3315 and Error 94

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
        rs!OrderNumber = Null  ' Or you might want to save empty string for this field
    End If

    rs!action = action
    rs!ActionTimestamp = Now
    rs!ActionBy = Environ("USERNAME")
    rs!ComputerName = Environ("COMPUTERNAME")

    ' FIX: Convert empty strings to Null to avoid Error 3315
    ' This also safely handles vbNullString passed from callers
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
```

### Step 2: Update cmdCommit_Click in Form_frmNumberRes

Change line ~2472 from:
```vba
LogOrderAction 0, Null, "BATCH_COMMIT", Null, Null, auditReason
```

To:
```vba
LogOrderAction 0, vbNullString, "BATCH_COMMIT", vbNullString, vbNullString, auditReason
```

**OR** better yet, just pass empty strings (LogOrderAction will handle the conversion):
```vba
LogOrderAction 0, "", "BATCH_COMMIT", "", "", auditReason
```

### Step 3: Update cmdStampBilled_Click in Form_frmOrderList

If you changed this to use Null, change it back to empty strings:

```vba
' Keep it simple - pass empty strings, LogOrderAction converts to Null
LogOrderAction lngSOID, sOrderNumber, "STAMP_BILLED", _
               IIf(IsNull(oldDateBilled), "", Format(oldDateBilled, "yyyy-mm-dd")), _
               Format(newDateBilled, "yyyy-mm-dd"), _
               "Billed date stamped by user"
```

## Implementation Steps

1. **Update LogOrderAction First** (this is critical!)
   - Open basAuditLogging module in VBA Editor
   - Find LogOrderAction function
   - Replace entire function with fixed version above
   - Save the module

2. **Revert cmdCommit Changes**
   - Open Form_frmNumberRes in Design View
   - View Code for cmdCommit button
   - Change line 2472 back to:
     ```vba
     LogOrderAction 0, "", "BATCH_COMMIT", "", "", auditReason
     ```
   - Save the form

3. **Test the Fix**
   - Try to commit a batch
   - Should work without Error 94 or Error 3315
   - Check audit log to verify entries are created

## Testing

After applying the fix:

```sql
-- Check recent audit entries
SELECT TOP 10
    EntryID,
    SOID,
    OrderNumber,
    action,
    oldStatus,
    newStatus,
    LEFT(reason, 50) AS reason_preview,
    ActionTimestamp,
    ActionBy
FROM tblOrderAudit
ORDER BY ActionTimestamp DESC;
```

You should see:
- ✅ BATCH_COMMIT entries being created
- ✅ oldStatus, newStatus, reason = NULL (not empty strings)
- ✅ No Error 94
- ✅ No Error 3315

## Summary

### The Issue
Passing `Null` to String parameters causes Error 94

### The Fix
1. Update LogOrderAction to convert empty strings to Null internally
2. Callers can safely pass empty strings (`""`)
3. LogOrderAction handles the conversion before saving to database

### Why This Approach
- ✅ Avoids Error 94 (Invalid use of Null)
- ✅ Avoids Error 3315 (Zero-length string)
- ✅ Simple for callers - just pass empty strings
- ✅ LogOrderAction does all the work
- ✅ Consistent behavior across all audit calls

## Key Takeaway

**Never pass `Null` to String parameters in VBA!**
- Use `""` or `vbNullString` for String parameters
- Let the function convert to Null internally if needed
- This prevents Error 94 while still storing Null in the database
