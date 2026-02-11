# Business Event Auditing - Quick Reference

**Last Updated**: 2026-02-11

---

## Current Audit Status

| Operation | Status | Action Type | Location |
|-----------|--------|-------------|----------|
| **Cancel** | ✅ Implemented | `CANCEL` / `CANCEL_FAILED` | `Form_frmOrderList.cmdCancelOrder_Click()` |
| **Backorder** | ✅ Implemented | `BACKORDER_CREATE` / `BATCH_COMMIT_FAILED` | `Form_frmOrderList.cmdCreateBackorder_Click()` |
| **Batch Commit** | ✅ Implemented | `BATCH_COMMIT` / `BATCH_COMMIT_FAILED` | `Form_dlgBatchGenerateOrders.cmdCommit_Click()` |
| **Stamp Billed** | ❌ **NOT IMPLEMENTED** | `STAMP_BILLED` / `STAMP_BILLED_FAILED` | `Form_frmOrderList.cmdStampBilled_Click()` |

---

## Quick Implementation Guide

### Step 1: Verify Audit Tables Exist

```sql
-- Check if audit tables exist
SELECT MSysObjects.Name
FROM MSysObjects
WHERE Name IN ('tblOrderAudit', 'tblAuditLog');
```

If tables don't exist, create them (see AUDIT_IMPLEMENTATION_CODE.vba for schema).

### Step 2: Update cmdStampBilled_Click()

1. Open `Form_frmOrderList` in Design View
2. Press `Alt+F11` to open VBA Editor
3. Find `cmdStampBilled_Click()` subroutine
4. Replace with enhanced version from `AUDIT_IMPLEMENTATION_CODE.vba`
5. Save and close

### Step 3: Test

```sql
-- Test query - should show your new STAMP audit entry
SELECT TOP 10 * FROM tblOrderAudit
ORDER BY ActionTimestamp DESC;
```

---

## Audit Function Reference

### LogOrderAction (Order-Specific Events)

```vba
LogOrderAction SOID, OrderNumber, action, oldStatus, newStatus, reason
```

**Example Usage**:
```vba
' Cancel operation
LogOrderAction 123, "2025-001", "CANCEL", "Active", "Canceled", "Customer request"

' Stamp billed date
LogOrderAction 123, "2025-001", "STAMP_BILLED", "", "2026-02-11", "User stamped billed date"

' Backorder creation
LogOrderAction 124, "2025-001-01", "BACKORDER_CREATE", "", "", "SourceSOID=123; BatchID={guid}"

' Batch commit
LogOrderAction 0, "", "BATCH_COMMIT", "", "", "Count=5; OrderType=SALES; BatchID={guid}"
```

### LogAudit (General Purpose)

```vba
LogAudit tableName, recordID, action, fieldName, oldValue, newValue, notes
```

**Example Usage**:
```vba
' General field change
LogAudit "SalesOrders", 123, "UPDATE", "CustomerName", "Old Corp", "New Corp", "Customer renamed"
```

---

## Action Types Reference

| Action | When Used | SOID | OrderNumber | oldStatus | newStatus | reason |
|--------|-----------|------|-------------|-----------|-----------|--------|
| `CANCEL` | Order canceled successfully | Order ID | Order # | "Active" | "Canceled" | Cancel reason |
| `CANCEL_FAILED` | Cancel operation failed | Order ID | Order # | "" | "" | Error message |
| `BACKORDER_CREATE` | Backorder created | New SOID | New Order # | "" | "" | Source details + BatchID |
| `BATCH_COMMIT` | Batch created successfully | 0 | "" | "" | "" | Batch details + count |
| `BATCH_COMMIT_FAILED` | Batch commit failed | 0 | "" | "" | "" | Error message |
| `STAMP_BILLED` | Billed date stamped | Order ID | Order # | Old date | New date | "User stamped..." |
| `STAMP_BILLED_FAILED` | Stamp operation failed | Order ID | Order # | "" | "" | Error message |

---

## Useful Queries

### Recent Activity
```sql
SELECT TOP 50
    ActionTimestamp,
    action,
    SOID,
    OrderNumber,
    ActionBy,
    reason
FROM tblOrderAudit
ORDER BY ActionTimestamp DESC;
```

### Stamp History
```sql
SELECT
    ActionTimestamp,
    OrderNumber,
    oldStatus AS PreviousBilledDate,
    newStatus AS NewBilledDate,
    ActionBy
FROM tblOrderAudit
WHERE action = 'STAMP_BILLED'
ORDER BY ActionTimestamp DESC;
```

### Order Lifecycle
```sql
SELECT
    ActionTimestamp,
    action,
    oldStatus,
    newStatus,
    ActionBy
FROM tblOrderAudit
WHERE SOID = 123  -- Replace with actual SOID
ORDER BY ActionTimestamp;
```

### Failed Operations
```sql
SELECT
    ActionTimestamp,
    action,
    SOID,
    OrderNumber,
    reason
FROM tblOrderAudit
WHERE action LIKE '%_FAILED'
ORDER BY ActionTimestamp DESC;
```

### User Activity Summary
```sql
SELECT
    ActionBy,
    action,
    COUNT(*) AS Count
FROM tblOrderAudit
GROUP BY ActionBy, action
ORDER BY ActionBy, action;
```

---

## Files Reference

| File | Purpose |
|------|---------|
| `AUDIT_IMPLEMENTATION_PLAN.md` | Complete analysis and implementation guide |
| `AUDIT_IMPLEMENTATION_CODE.vba` | VBA code with enhanced cmdStampBilled_Click() |
| `AUDIT_QUICK_REFERENCE.md` | This file - quick reference guide |
| `AllModulesDump` | Full VBA code export (read-only reference) |
| `Index` | Module index |

---

## Code Location Map

### Forms with Audit Logging

**Form_frmOrderList**:
- `cmdCancelOrder_Click()` - Lines ~652-745 - ✅ Has audit logging
- `cmdCreateBackorder_Click()` - Lines ~325-407 - ✅ Has audit logging
- `cmdStampBilled_Click()` - Lines ~614-650 - ❌ **NEEDS audit logging**

**Form_dlgBatchGenerateOrders**:
- `cmdCommit_Click()` - Lines ~2389-2495 - ✅ Has audit logging

### Modules

**basAuditLogging** (Lines 4871-4948):
- `LogAudit()` - General purpose logging
- `LogOrderAction()` - Order-specific logging

**basOrderFunctions** (Lines 3596-3750):
- `CreateBackorder()` - Creates backorder order (called by cmdCreateBackorder)

**basBatchWizard**:
- `CommitBatch()` - Lines ~2827-2880 - Creates batch orders

---

## What Changes Are Needed?

### ✅ Already Working
- Cancel operation audit logging
- Backorder creation audit logging
- Batch commit audit logging
- Audit infrastructure (tables + functions)

### ❌ Needs Implementation
- **Stamp billed date audit logging** (10 lines of code in `cmdStampBilled_Click()`)

### Estimated Work
- **Time**: 30 minutes coding + 1 hour testing
- **Risk**: LOW - isolated change, non-breaking
- **Files**: 1 form (Form_frmOrderList)
- **Lines Changed**: ~10 lines added

---

## Testing Checklist

Before deploying to production:

- [ ] Verify tblOrderAudit table exists
- [ ] Test Cancel - creates CANCEL audit entry
- [ ] Test Cancel with error - creates CANCEL_FAILED entry
- [ ] Test Backorder - creates BACKORDER_CREATE entry
- [ ] Test Backorder with error - creates failure entry
- [ ] Test Batch Commit - creates BATCH_COMMIT entry
- [ ] Test Batch Commit with error - creates BATCH_COMMIT_FAILED entry
- [ ] Test Stamp (after implementation) - creates STAMP_BILLED entry
- [ ] Test Stamp with error - creates STAMP_BILLED_FAILED entry
- [ ] Verify timestamps are accurate
- [ ] Verify usernames are captured correctly
- [ ] Run summary queries to review audit trail
- [ ] Test that failed audit logging doesn't break operations

---

## Support & Troubleshooting

### Issue: No audit entries appearing
**Solution**:
1. Check tblOrderAudit table exists
2. Verify field names match (case-sensitive)
3. Check basAuditLogging module is present
4. Audit logging uses `On Error Resume Next` so errors are silent

### Issue: Wrong timestamp
**Solution**: Check computer clock settings

### Issue: Wrong username
**Solution**: `Environ("USERNAME")` gets Windows username - verify with `Debug.Print Environ("USERNAME")`

### Issue: "Object doesn't support this property"
**Solution**: Verify LogOrderAction function exists in basAuditLogging module

---

## Compliance Notes

This audit system captures:
- **Who**: `ActionBy` (Windows username)
- **What**: `action` (operation type)
- **When**: `ActionTimestamp` (date/time)
- **Where**: `ComputerName` (machine name)
- **Why**: `reason` (business reason or error)
- **Change Details**: `oldStatus` → `newStatus`

Audit entries are **immutable** (no updates/deletes in code).

For compliance reporting, see queries in AUDIT_IMPLEMENTATION_CODE.vba.

---

*End of Quick Reference*
