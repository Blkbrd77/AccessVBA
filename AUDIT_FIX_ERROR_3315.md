# Fix for Error 3315: Field Can't Be Zero-Length String

**Issue**: Audit logging fails with Error 3315
**Root Cause**: oldStatus/newStatus fields don't allow zero-length strings
**Solution**: Pass Null instead of empty strings when no value exists

---

## The Problem

The `tblOrderAudit` table has fields with "Allow Zero Length" = No:
- `oldStatus`
- `newStatus`
- `reason`

When the code passes empty strings `""`, Access rejects them with Error 3315.

This happens for:
- **New orders**: No old status exists
- **Stamp operations**: DateBilled was Null (no old value)
- **Backorder creation**: No status change, just creation

---

## The Fix

Change all empty string parameters to `Null` values.

### Helper Function (Add to basAuditLogging)

```vba
'================================================================================
' Helper: Convert empty strings to Null for database fields
'================================================================================
Private Function NullIfEmpty(ByVal value As String) As Variant
    If Len(Trim(value)) = 0 Then
        NullIfEmpty = Null
    Else
        NullIfEmpty = value
    End If
End Function
```

### Updated LogOrderAction (basAuditLogging)

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

    On Error Resume Next

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblOrderAudit", dbOpenDynaset)

    rs.AddNew
    rs!SOID = SOID
    rs!OrderNumber = OrderNumber
    rs!action = action
    rs!ActionTimestamp = Now
    rs!ActionBy = Environ("USERNAME")
    rs!ComputerName = Environ("COMPUTERNAME")

    ' FIX: Use Null instead of empty strings
    rs!oldStatus = NullIfEmpty(oldStatus)
    rs!newStatus = NullIfEmpty(newStatus)
    rs!reason = NullIfEmpty(reason)

    rs.Update

    rs.Close

    On Error GoTo 0
End Sub
```

### Alternative: Inline Fix (No Helper Function)

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

    On Error Resume Next

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblOrderAudit", dbOpenDynaset)

    rs.AddNew
    rs!SOID = SOID
    rs!OrderNumber = OrderNumber
    rs!action = action
    rs!ActionTimestamp = Now
    rs!ActionBy = Environ("USERNAME")
    rs!ComputerName = Environ("COMPUTERNAME")

    ' FIX: Use IIf to convert empty strings to Null
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

    On Error GoTo 0
End Sub
```

---

## Updated Calling Code

No changes needed to calling code! The fix in LogOrderAction handles everything.

Current calls like this will work:
```vba
' Stamp billed - no old value
LogOrderAction lngSOID, sOrderNumber, "STAMP_BILLED", _
               "", _  ' Empty string will become Null
               Format(newDateBilled, "yyyy-mm-dd"), _
               "Billed date stamped by user"

' Backorder - no status change
LogOrderAction newSOID, newOrderNum, "BACKORDER_CREATE", _
               "", "", _  ' Both empty, both become Null
               "SourceSOID=" & SourceSOID & "; BatchID=" & newBatchID

' Batch commit - no individual status
LogOrderAction 0, "", "BATCH_COMMIT", _
               "", "", _  ' Both empty, both become Null
               auditReason
```

---

## Alternative Solution: Change Table Design

Instead of changing code, you could change the table:

1. Open `tblOrderAudit` in Design View
2. Select `oldStatus` field
3. Set "Allow Zero Length" = **Yes**
4. Repeat for `newStatus` and `reason`
5. Save table

**Pros**: No code changes needed
**Cons**: Allows empty strings in database (may not be desired)

**Recommendation**: Fix the code (preferred) - cleaner data with Nulls instead of empty strings

---

## Test After Fix

```vba
' Test with empty strings - should work now
Sub TestFixFor3315()
    Debug.Print "Testing fix for Error 3315..."

    ' This should now work without error
    LogOrderAction 9999, "TEST-FIX", "TEST_EMPTY_STRINGS", "", "", ""

    Debug.Print "Success! Checking record..."

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT * FROM tblOrderAudit WHERE SOID=9999", _
        dbOpenSnapshot)

    If Not rs.EOF Then
        Debug.Print "Record created successfully"
        Debug.Print "  oldStatus is Null: " & IsNull(rs!oldStatus)
        Debug.Print "  newStatus is Null: " & IsNull(rs!newStatus)
        Debug.Print "  reason is Null: " & IsNull(rs!reason)
    End If

    rs.Close

    ' Cleanup
    CurrentDb.Execute "DELETE FROM tblOrderAudit WHERE SOID=9999", dbFailOnError
    Debug.Print "Test complete - record deleted"
End Sub
```

**Expected output**:
```
Testing fix for Error 3315...
Success! Checking record...
Record created successfully
  oldStatus is Null: True
  newStatus is Null: True
  reason is Null: True
Test complete - record deleted
```

---

## Implementation Steps

1. **Open basAuditLogging module**
2. **Replace LogOrderAction** with the fixed version above
3. **Save the module**
4. **Test stamp operation** - should work now!
5. **Verify audit entry created** with this query:
   ```sql
   SELECT TOP 5 * FROM tblOrderAudit
   ORDER BY ActionTimestamp DESC;
   ```

---

## Why This Happens

Access table fields have a property called "Allow Zero Length":
- **No** (default for new fields): Field must be Null or have actual content
- **Yes**: Field can be Null, empty string, or have content

When "Allow Zero Length" = No, this fails:
```vba
rs!oldStatus = ""  ' Error 3315!
```

This works:
```vba
rs!oldStatus = Null  ' OK
```

---

## Complete Fixed Code

Here's the complete, production-ready LogOrderAction:

```vba
'================================================================================
' basAuditLogging - LogOrderAction (FIXED for Error 3315)
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
    ' Fixed: Converts empty strings to Null to avoid Error 3315

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

    ' Convert empty strings to Null
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
```

---

*End of Fix Documentation*
