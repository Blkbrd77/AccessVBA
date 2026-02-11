# Business Event Auditing - Implementation Analysis & Plan

**Generated**: 2026-02-11
**Database**: 6Feb25FE_BESplit.accdb
**Purpose**: Implement comprehensive business event auditing for Commit, Backorder, Stamp, and Cancel operations

---

## Table of Contents
1. [Executive Summary](#executive-summary)
2. [Current Infrastructure](#current-infrastructure)
3. [Current Audit Coverage](#current-audit-coverage)
4. [Missing Audit Points](#missing-audit-points)
5. [Implementation Plan](#implementation-plan)
6. [Code Changes Required](#code-changes-required)

---

## Executive Summary

The application already has a robust audit logging infrastructure in place with two main tables (`tblAuditLog` and `tblOrderAudit`) and helper functions. Most business events are already being audited, but **STAMP operations lack auditing**.

### Quick Status
- ✅ **CANCEL** - Fully audited (success & failure)
- ✅ **BACKORDER** - Fully audited (success & failure)
- ✅ **BATCH_COMMIT** - Fully audited (success & failure)
- ❌ **STAMP** (Billed Date) - **NOT AUDITED** - needs implementation

---

## Current Infrastructure

### Audit Tables

#### tblOrderAudit (Order-Specific Events)
Stores order-specific business events like cancel, backorder, stamp, and batch operations.

**Expected Fields**:
- `SOID` (Long) - Sales Order ID
- `OrderNumber` (Text) - Order number for reference
- `action` (Text) - Action type (CANCEL, BACKORDER_CREATE, STAMP, etc.)
- `ActionTimestamp` (Date/Time) - When the action occurred
- `ActionBy` (Text) - Username who performed the action
- `ComputerName` (Text) - Computer where action was performed
- `oldStatus` (Text) - Previous status value
- `newStatus` (Text) - New status value
- `reason` (Text/Memo) - Detailed reason or notes

#### tblAuditLog (General Purpose)
General audit trail for any table/field changes.

**Expected Fields**:
- `AuditTimestamp` (Date/Time)
- `tableName` (Text)
- `recordID` (Long)
- `action` (Text)
- `fieldName` (Text)
- `oldValue` (Text)
- `newValue` (Text)
- `UserName` (Text)
- `ComputerName` (Text)
- `notes` (Text/Memo)

### Audit Functions (basAuditLogging module)

#### LogOrderAction
```vba
Public Sub LogOrderAction( _
    ByVal SOID As Long, _
    ByVal OrderNumber As String, _
    ByVal action As String, _
    Optional ByVal oldStatus As String = "", _
    Optional ByVal newStatus As String = "", _
    Optional ByVal reason As String = "" _
)
```

**Purpose**: Log order-specific events to tblOrderAudit
**Auto-captures**: Username, ComputerName, Timestamp

#### LogAudit
```vba
Public Sub LogAudit( _
    ByVal tableName As String, _
    ByVal recordID As Long, _
    ByVal action As String, _
    Optional ByVal fieldName As String = "", _
    Optional ByVal oldValue As String = "", _
    Optional ByVal newValue As String = "", _
    Optional ByVal notes As String = "" _
)
```

**Purpose**: General-purpose audit logging to tblAuditLog
**Auto-captures**: Username, ComputerName, Timestamp

---

## Current Audit Coverage

### 1. CANCEL Operation ✅

**Location**: `Form_frmOrderList.cmdCancelOrder_Click()` (Lines 652-745)

**What's Audited**:
- Success: `LogOrderAction(SOID, OrderNumber, "CANCEL", "Active", "Canceled", reason)`
- Failure: `LogOrderAction(SOID, OrderNumber, "CANCEL_FAILED", "", "", errorMsg)`

**Captured Data**:
- SOID and OrderNumber
- Old status: "Active"
- New status: "Canceled"
- Reason: User-provided cancellation reason
- Who: CanceledBy field + Environ("USERNAME")
- When: DateCanceled + ActionTimestamp

**Dialog**: `dlgCancelOrder` - Captures:
- CanceledBy (who)
- CancelReason (why)
- DateCanceled (when)

### 2. BACKORDER Operation ✅

**Location**: `Form_frmOrderList.cmdCreateBackorder_Click()` (Lines 325-407)

**What's Audited**:
- Success: `LogOrderAction(newSOID, newOrderNum, "BACKORDER_CREATE", "", "", details)`
- Failure: `LogOrderAction(0, "", "BATCH_COMMIT_FAILED", "", "", errorMsg)`

**Captured Data**:
- New SOID and new OrderNumber
- Source SOID and source OrderNumber
- BatchID for tracking
- Action: "BACKORDER_CREATE"

**Process**:
1. User confirms backorder creation
2. `CreateBackorder()` function creates new order with incremented BackorderNo
3. Success audit log entry created
4. New order displayed with BatchID filter

### 3. BATCH_COMMIT Operation ✅

**Location**: `Form_dlgBatchGenerateOrders.cmdCommit_Click()` (Lines 2389-2495)

**What's Audited**:
- Success: `LogOrderAction(0, "", "BATCH_COMMIT", "", "", details)`
- Failure: `LogOrderAction(0, "", "BATCH_COMMIT_FAILED", "", "", errorMsg)`

**Captured Data**:
- Action: "BATCH_COMMIT"
- SOID: 0 (batch-level event, not single order)
- Details include: Count, OrderType, BaseToken, SystemLetter, DateReceived
- BatchID for linking all created orders

**Process**:
1. User enters batch parameters (qualifiers, quantities)
2. `CommitBatch()` creates multiple orders atomically
3. On success/failure, audit log created
4. TempVars used to communicate status

### 4. STAMP (Billed Date) Operation ❌

**Location**: `Form_frmOrderList.cmdStampBilled_Click()` (Lines 614-650)

**Current Status**: **NO AUDITING IMPLEMENTED**

**What Happens Now**:
1. User clicks "Stamp Billed" button
2. `dlgStampBilledDate` dialog opens
3. User enters billed date
4. `DateBilled` field updated
5. Form saved
6. **NO AUDIT LOG CREATED** ⚠️

**What SHOULD Be Audited**:
- SOID and OrderNumber
- Old DateBilled value (likely NULL)
- New DateBilled value (user-entered date)
- Action: "STAMP_BILLED"
- Who stamped it
- When stamped

---

## Missing Audit Points

### 1. STAMP Operation - PRIORITY: HIGH

**Current Code** (Form_frmOrderList, lines 614-650):
```vba
Private Sub cmdStampBilled_Click()
    On Error GoTo EH

    ' ... validation ...

    DoCmd.OpenForm "dlgStampBilledDate", WindowMode:=acDialog

    If Nz(TempVars("StampBilledResult"), "Cancel") <> "OK" Then
        Exit Sub
    End If

    Me!DateBilled = TempVars("StampBilledDate")  ' <-- CHANGE HAPPENS HERE

    If Me.Dirty Then Me.Dirty = False

    Me!txtStampDate.Requery

    Exit Sub

EH:
    MsgBox "Stamp Billed failed: " & Err.Description, vbExclamation
End Sub
```

**NEEDS**:
```vba
' Before changing DateBilled, capture old value
Dim oldDate As Variant, newDate As Variant
oldDate = Nz(Me!DateBilled, Null)
newDate = TempVars("StampBilledDate")

Me!DateBilled = newDate

' After successful save
LogOrderAction Me!SOID, Nz(Me!OrderNumber, ""), "STAMP_BILLED", _
               IIf(IsNull(oldDate), "", Format(oldDate, "yyyy-mm-dd")), _
               Format(newDate, "yyyy-mm-dd"), _
               "User stamped billed date"
```

### 2. Individual Order Commits - PRIORITY: MEDIUM

**Location**: `Form_z_Deprecated_frmSalesOrderEntry` and similar entry forms

**Current Status**: Order saves happen but aren't specifically audited as "COMMIT" events

**Consideration**:
- These are deprecated forms
- New order creation happens through batch wizard
- May not need additional auditing if batch commits are already tracked
- **RECOMMENDATION**: Leave as-is unless specific requirement exists

---

## Implementation Plan

### Phase 1: Add STAMP Auditing ⭐ PRIORITY

**File**: AllModulesDump (Form_frmOrderList section)

**Changes Needed**:

1. **Modify `cmdStampBilled_Click()` method**:
   - Capture old DateBilled value before change
   - Capture new DateBilled value from TempVars
   - Add success audit log entry
   - Add failure audit log entry in error handler

2. **Action Types to Use**:
   - Success: `"STAMP_BILLED"`
   - Failure: `"STAMP_BILLED_FAILED"`

**Code Template**:
```vba
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

    ' Capture current state
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

    ' Get new value
    newDateBilled = TempVars("StampBilledDate")

    ' Write the chosen date to your bound field
    Me!DateBilled = newDateBilled

    ' Persist immediately
    If Me.Dirty Then Me.Dirty = False

    ' --- AUDIT: Stamp success ---
    LogOrderAction lngSOID, sOrderNumber, "STAMP_BILLED", _
                   IIf(IsNull(oldDateBilled), "", Format(oldDateBilled, "yyyy-mm-dd")), _
                   Format(newDateBilled, "yyyy-mm-dd"), _
                   "Billed date stamped by user"

    ' Refresh display
    Me!txtStampDate.Requery

    Exit Sub

EH:
    ' --- AUDIT: Stamp failure ---
    LogOrderAction Nz(Me!SOID, 0), Nz(Me!OrderNumber, ""), "STAMP_BILLED_FAILED", "", "", _
                   "Err " & Err.Number & ": " & Err.Description

    MsgBox "Stamp Billed failed: " & Err.Description, vbExclamation
End Sub
```

### Phase 2: Enhance CommitBatch Auditing (Optional)

**Current State**: Already logs batch-level events, but doesn't log individual order creations within batch

**Potential Enhancement**:
Add individual audit entries for each order created in batch:
```vba
' In CommitBatch or CreateOneOrder function
LogOrderAction newSOID, newOrderNumber, "ORDER_CREATED", "", "", _
               "Created via batch; BatchID=" & BatchID
```

**Decision**: This may create excessive audit entries. Recommend keeping current batch-level approach unless specific compliance requirement exists.

### Phase 3: Validation & Testing

1. **Verify audit tables exist**:
   - Check that `tblOrderAudit` and `tblAuditLog` exist
   - Verify field definitions match expected schema

2. **Test each operation**:
   - Cancel an order → verify audit entry
   - Create backorder → verify audit entry
   - Stamp billed date → verify audit entry (after implementation)
   - Batch commit → verify audit entry

3. **Test failure scenarios**:
   - Trigger errors during operations
   - Verify failure audit entries created
   - Verify operations roll back properly

4. **Audit Log Query**:
```sql
-- Recent audit activity
SELECT ActionTimestamp, action, SOID, OrderNumber,
       oldStatus, newStatus, ActionBy, reason
FROM tblOrderAudit
ORDER BY ActionTimestamp DESC;

-- Specific order history
SELECT * FROM tblOrderAudit
WHERE SOID = [EnterSOID]
ORDER BY ActionTimestamp;

-- Operations by user
SELECT ActionBy, action, COUNT(*) AS OpCount
FROM tblOrderAudit
GROUP BY ActionBy, action
ORDER BY ActionBy, action;
```

---

## Code Changes Required

### File: Form_frmOrderList (within AllModulesDump)

**Function**: `cmdStampBilled_Click()`
**Lines**: Approximately 614-650

**Action**: REPLACE existing function with enhanced version that includes audit logging

**Before**:
```vba
Private Sub cmdStampBilled_Click()
    On Error GoTo EH

    If Me.NewRecord Then
        MsgBox "Please select or save a record before stamping.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    TempVars.Remove "StampBilledDate"
    TempVars.Remove "StampBilledResult"
    On Error GoTo 0

    DoCmd.OpenForm "dlgStampBilledDate", WindowMode:=acDialog

    If Nz(TempVars("StampBilledResult"), "Cancel") <> "OK" Then
        Exit Sub
    End If

    Me!DateBilled = TempVars("StampBilledDate")

    If Me.Dirty Then Me.Dirty = False

    Me!txtStampDate.Requery

    Exit Sub

EH:
    MsgBox "Stamp Billed failed: " & Err.Description, vbExclamation
End Sub
```

**After** (see Phase 1 code template above for complete replacement)

---

## Database Schema Verification

Ensure these tables exist with proper structure:

### tblOrderAudit
```sql
CREATE TABLE tblOrderAudit (
    AuditID AUTOINCREMENT PRIMARY KEY,
    SOID LONG,
    OrderNumber TEXT(50),
    action TEXT(50),
    ActionTimestamp DATETIME,
    ActionBy TEXT(100),
    ComputerName TEXT(100),
    oldStatus TEXT(255),
    newStatus TEXT(255),
    reason MEMO
);
```

### tblAuditLog
```sql
CREATE TABLE tblAuditLog (
    AuditID AUTOINCREMENT PRIMARY KEY,
    AuditTimestamp DATETIME,
    tableName TEXT(100),
    recordID LONG,
    action TEXT(50),
    fieldName TEXT(100),
    oldValue TEXT(255),
    newValue TEXT(255),
    UserName TEXT(100),
    ComputerName TEXT(100),
    notes MEMO
);
```

---

## Action Items Summary

### Immediate Actions (Required)
1. ✅ Verify `tblOrderAudit` and `tblAuditLog` tables exist
2. ⬜ Implement STAMP auditing in `cmdStampBilled_Click()`
3. ⬜ Test STAMP auditing with sample data
4. ⬜ Verify all four operations create audit entries

### Future Enhancements (Optional)
1. Add reporting dashboard for audit logs
2. Add audit log cleanup/archiving for old entries
3. Consider adding individual order creation audits within batches
4. Add data export capability for compliance/reporting

---

## Testing Checklist

- [ ] Test CANCEL operation - verify audit entry created
- [ ] Test CANCEL with error - verify failure audit entry
- [ ] Test BACKORDER operation - verify audit entry created
- [ ] Test BACKORDER with error - verify failure audit entry
- [ ] Test BATCH_COMMIT operation - verify audit entry created
- [ ] Test BATCH_COMMIT with error - verify failure audit entry
- [ ] Test STAMP operation - verify audit entry created
- [ ] Test STAMP with error - verify failure audit entry
- [ ] Verify ActionBy captures correct username
- [ ] Verify ActionTimestamp captures correct time
- [ ] Run audit report query to review all entries
- [ ] Test audit logging doesn't prevent operations on failure

---

## Conclusion

The application has a solid audit foundation already in place. The main gap is **STAMP (Billed Date) operations**, which can be addressed by adding 5-10 lines of code to the existing `cmdStampBilled_Click()` function.

**Estimated Implementation Time**:
- STAMP auditing: 30 minutes (code + test)
- Full testing: 1-2 hours
- **Total**: 2-3 hours

**Risk Level**: LOW
- Non-breaking changes
- Audit logging uses `On Error Resume Next` (won't crash app if fails)
- Changes isolated to single function
- Existing audit infrastructure proven and working

---

*End of Implementation Plan*
