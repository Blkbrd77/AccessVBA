# Microsoft Copilot Implementation Guide

Guide for implementing the Order Management System updates using Microsoft Copilot.

---

## Desired End-State User Flow

```
┌─────────────────┐
│  frmOrderList   │  ◄── User starts here, views existing orders
└────────┬────────┘
         │
         │ [New Batch] button
         ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                      dlgBatchGenerateOrders                                  │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  Step 1: Order Type                                                          │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │  Order Type: (•) SALES  ( ) PROJECT                                    │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
│                          │                                                   │
│           ┌──────────────┴──────────────┐                                   │
│           ▼                              ▼                                   │
│  ┌─────────────────────┐      ┌─────────────────────┐                       │
│  │ If SALES:           │      │ If PROJECT:         │                       │
│  │ Auto-generate next  │      │ Show Project combo  │                       │
│  │ BaseToken (576xxx)  │      │ User selects from   │                       │
│  │                     │      │ Projects table      │                       │
│  └─────────────────────┘      └─────────────────────┘                       │
│                          │                                                   │
│                          ▼                                                   │
│  Step 2: Order Details                                                       │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │  Base Token:     [576001] or [TAGO25]  (auto-filled)                   │ │
│  │  System Letter:  [N ▼]                  (dropdown)                     │ │
│  │  Customer:       [ACME Corp ▼]          (dropdown from Customers)      │ │
│  │  PO #:           [_______________]      (text entry)                   │ │
│  │  Date Received:  [1/23/2026]            (default: today)               │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
│                                                                              │
│  Step 3: Qualifier Selection                                                 │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │  Qualifier │ Description          │ Qty                                │ │
│  │  ──────────┼──────────────────────┼─────                               │ │
│  │  CM        │ Construction Mgmt    │ [5]                                │ │
│  │  CE        │ Civil Engineering    │ [3]                                │ │
│  │  EL        │ Electrical           │ [0]                                │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
│                                                                              │
│  Step 4: Preview                                                             │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │  576001-CMN001-00                                                      │ │
│  │  576001-CMN002-00                                                      │ │
│  │  576001-CEN001-00                                                      │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
│                                                                              │
│              [Preview Numbers]    [Commit Batch]    [Cancel]                 │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
         │
         │ [Commit Batch]
         ▼
┌─────────────────┐
│  frmOrderList   │  ◄── Returns here, filtered to show new batch
│  (filtered)     │
└─────────────────┘
```

---

## How to Use This Guide with Copilot

### Step 1: Start a New Copilot Conversation

Open Microsoft Copilot and paste the following context message:

---

**CONTEXT MESSAGE (Copy and paste this first):**

```
I'm working on a Microsoft Access VBA application for order management. I need help implementing changes to the batch order generation system.

DATABASE STRUCTURE:
- SalesOrders: Main order header table (SOID, OrderType, BaseToken, SystemLetter, BackorderNo, CustomerCode, CustomerName, PONumber, DateReceived, OrderNumber, BatchID)
- SalesOrderEntry: Qualifier details per order (SOEntryID, SOID, QualifierCode, SequenceNo, OrderNumberDisplay)
- OrderSeq: Sequence counters (Scope, BaseToken, QualifierCode, SystemLetter, NextSeq)
- QualifierType: Master qualifier codes (QCID, QualifierCode, QualifierName)
- Customers: Customer lookup (CustomerID, CustomerCode, CustomerName, IsActive)
- Projects: Project base tokens (ProjectID, BaseToken, Program)
- SystemLetter: System letters (SystemLetter, Description, ActiveFlag)

ORDER NUMBER FORMAT (unified for both SALES and PROJECT):
<BaseToken>-<QualifierCode><SystemLetter><Seq>-<BackorderNo>
Examples:
- SALES: 576001-CMN001-00
- PROJECT: TAGO25-CMN001-00

SALES BASETOKEN SCHEME:
- Year 2026 starts at 576000
- Year 2027 starts at 577000
- Formula: 576000 + (year - 2026) * 1000

KEY EXISTING MODULES:
- basSeqAllocator: ReserveSeq(scope, BaseToken, QualifierCode, SystemLetter) - reserves next sequence number
- basBatchWizard: BuildPreview(), CommitBatch(), CreateOneOrder() - batch operations
- basOrderNumbering: NextSalesBaseToken() - generates next SALES base token

EXISTING FORMS:
- frmOrderList: Main order list view
- dlgBatchGenerateOrders: Batch order creation dialog (needs modification)
- fsubQualifierQty: Subform for qualifier quantity entry

I'll give you specific tasks one at a time. Please provide VBA code for Access.
```

---

### Step 2: Fix Known Issues First

Before implementing new features, fix this known bug:

---

## Known Issues (Fix First)

### Issue: "Enter Parameter Value" for BatchID

**Symptom:** After committing a batch and returning to frmOrderList, Access shows "Enter Parameter Value" dialog asking for BatchID.

**Cause:** The frmOrderList query doesn't include BatchID field, so when the code tries to filter by BatchID, Access can't find it.

**Copilot Prompt (copy and paste):**

```
I have a bug in my Access application. When I commit a batch of orders and return to frmOrderList, I get an "Enter Parameter Value" dialog asking for BatchID.

THE PROBLEM:
The frmOrderList form's RecordSource is a query that doesn't include BatchID. When my code tries to filter by BatchID, Access can't find it and prompts for a parameter.

CURRENT QUERY (frmOrderList):
SELECT S.SOID, S.OrderNumber, S.CustomerName, S.PONumber, S.DateReceived, F.QualifierCode, F.SequenceNo, S.BaseToken, S.SystemLetter, S.BackorderNo
FROM SalesOrders AS S LEFT JOIN qryFirstQualifierPerSO AS F ON S.SOID = F.SOID
WHERE S.ActiveFlag = True
ORDER BY S.SOID DESC;

THE CODE THAT SETS THE FILTER (in cmdCommit_Click):
If CurrentProject.AllForms("frmOrderList").IsLoaded Then
    Forms("frmOrderList").Filter = "BatchID='" & Replace(batchID, "'", "''") & "'"
    Forms("frmOrderList").FilterOn = True
    Forms("frmOrderList").Requery
Else
    DoCmd.OpenForm "frmOrderList", , , "BatchID='" & Replace(batchID, "'", "''") & "'"
End If

WHAT I NEED:
1. The corrected SQL for the frmOrderList query that includes S.BatchID
2. Confirm the filter syntax is correct

The SalesOrders table does have a BatchID field (Text 36 characters).
```

**The Fix:**

Update the `frmOrderList` query to include `S.BatchID`:

```sql
SELECT S.SOID, S.OrderNumber, S.CustomerName, S.PONumber, S.DateReceived,
       F.QualifierCode, F.SequenceNo, S.BaseToken, S.SystemLetter,
       S.BackorderNo, S.BatchID
FROM SalesOrders AS S
LEFT JOIN qryFirstQualifierPerSO AS F ON S.SOID = F.SOID
WHERE S.ActiveFlag = True
ORDER BY S.SOID DESC;
```

**How to Apply:**
1. Open the Navigation Pane in Access
2. Find the query named `frmOrderList`
3. Open it in Design View (or SQL View)
4. Add `S.BatchID` to the SELECT clause
5. Save and close the query
6. Test by committing a batch

---

### Step 3: Feed Tasks One at a Time

After fixing known issues and pasting the context, give Copilot these tasks in order:

---

## Task List for Copilot

### Task 1: Create the Form Layout

```
TASK 1: Help me design the control layout for dlgBatchGenerateOrders.

I need these controls:
- fraOrderType: Option group with optSales and optProject radio buttons
- cboProject: Combo box for project selection (visible only for PROJECT)
- txtBaseToken: Text box showing the base token (auto-filled)
- cboSystemLetter: Combo box for system letter selection
- cboCustomerCode: Combo box for customer selection (bound to Customers table)
- txtCustomerName: Text box (auto-fills from customer selection)
- txtPONumber: Text box for PO entry
- txtDateReceived: Text box with default value of Date()
- subQualifierQty: Subform control linked to fsubQualifierQty
- lstPreview: List box for order number preview
- cmdPreview: Command button "Preview Numbers"
- cmdCommit: Command button "Commit Batch"
- cmdCancel: Command button "Cancel"

Please provide:
1. Recommended control positions (Left, Top, Width, Height in twips)
2. Control properties to set in the property sheet
```

---

### Task 2: Option Group Logic

```
TASK 2: Write the VBA code for the Order Type option group (fraOrderType).

When user selects SALES:
- Hide cboProject
- Call NextSalesBaseToken() and put result in txtBaseToken
- txtBaseToken should be locked (read-only)

When user selects PROJECT:
- Show cboProject
- Clear txtBaseToken
- txtBaseToken gets populated when user selects from cboProject

Please write the fraOrderType_AfterUpdate event procedure.
```

---

### Task 3: Project Combo Box

```
TASK 3: Write the VBA code for cboProject combo box.

Row Source should be:
SELECT BaseToken, Program FROM Projects WHERE ActiveFlag=True ORDER BY BaseToken;

Column Count: 2
Column Widths: 1";2"
Bound Column: 1

When user selects a project:
- Put the BaseToken (column 0) into txtBaseToken

Please write:
1. The Row Source SQL
2. The cboProject_AfterUpdate event procedure
```

---

### Task 4: Customer Combo Box

```
TASK 4: Write the VBA code for cboCustomerCode combo box.

Row Source should query Customers table for active customers.
Show CustomerCode and CustomerName.
When selected, auto-fill txtCustomerName with the customer name.

Please write:
1. The Row Source SQL
2. The cboCustomerCode_AfterUpdate event procedure
```

---

### Task 5: Form Open Event

```
TASK 5: Write the Form_Open event for dlgBatchGenerateOrders.

On open:
1. Default fraOrderType to PROJECT (value = 2)
2. Call the option group AfterUpdate to set initial visibility
3. Default txtDateReceived to today's date
4. Call Reset_Qty_From_Qualifiers (existing sub in basResetRoutine)
5. Requery the subQualifierQty subform
6. Clear tmpBatchPreview table
7. Requery lstPreview

Please write the Form_Open event procedure.
```

---

### Task 6: Preview Button

```
TASK 6: Write the cmdPreview_Click event procedure.

This should:
1. Validate required fields (OrderType, BaseToken, SystemLetter, at least one qualifier with Qty > 0)
2. Get the scope ("SALES" or "PROJECT" based on fraOrderType)
3. Call BuildPreview(scope, txtBaseToken, cboSystemLetter) from basBatchWizard module
4. Requery lstPreview to show the preview

The BuildPreview sub already exists and populates tmpBatchPreview table.

Please write the cmdPreview_Click event procedure with proper validation.
```

---

### Task 7: Commit Button

```
TASK 7: Write the cmdCommit_Click event procedure.

This should:
1. Validate all required fields
2. Get values: scope, BaseToken, SystemLetter, CustomerCode, CustomerName, PONumber, DateReceived
3. Call CommitBatch() from basBatchWizard module
4. On success: show message with count, close this form, open/requery frmOrderList filtered to the new BatchID
5. On failure: show error message

The CommitBatch function signature is:
CommitBatch(scope, BaseToken, SystemLetter, CustomerCode, CustomerName, PONumber, DateReceived, ByRef createdCount, ByRef batchID) As Boolean

Please write the cmdCommit_Click event procedure.
```

---

### Task 8: Cancel Button and Form Unload

```
TASK 8: Write the cmdCancel_Click and Form_Unload events.

cmdCancel_Click:
- Close the form without saving

Form_Unload:
- Clean up tmpBatchPreview table
- Clean up tmpQualifierQty table (optional)

Please write both event procedures.
```

---

### Task 9: Add Button to frmOrderList

```
TASK 9: I need to add a "New Batch" button to frmOrderList that opens dlgBatchGenerateOrders.

Please write:
1. The cmdNewBatch_Click event procedure for frmOrderList
2. It should open dlgBatchGenerateOrders as a dialog (acDialog)
3. After the dialog closes, requery frmOrderList
```

---

### Task 10: Validation Helper Function

```
TASK 10: Write a helper function to validate the form before preview/commit.

ValidateForm() As Boolean should check:
1. fraOrderType has a value (1 or 2)
2. txtBaseToken is not empty
3. cboSystemLetter is not empty
4. At least one qualifier in tmpQualifierQty has Qty > 0

Return True if valid, False if not. Show appropriate MsgBox for each validation failure.

Please write the ValidateForm function.
```

---

## Tips for Working with Copilot

1. **One task at a time**: Don't overwhelm Copilot. Give one task, get the code, test it, then move to the next.

2. **Provide error messages**: If code doesn't work, paste the exact error message back to Copilot.

3. **Ask for explanations**: If you don't understand something, ask "Can you explain what this line does?"

4. **Request modifications**: Say "That's close, but I need it to also do X" to refine the code.

5. **Test incrementally**: After each task, test that piece before moving on.

---

## Files to Reference

If Copilot needs more context, you can share these files from the repo:

| File | Purpose |
|------|---------|
| `numberResContext` | Full database context export |
| `USER_FLOW.md` | Current user flow diagram |
| `RECOMMENDED_CLEANUP.md` | Cleanup tasks and planned changes |
| `INSTRUCTIONS_SALES_BATCH.md` | Detailed implementation instructions |

---

## Quick Reference: Key Function Signatures

```vba
' basOrderNumbering
Public Function NextSalesBaseToken() As Long

' basSeqAllocator
Public Function ReserveSeq(scope, BaseToken, QualifierCode, SystemLetter) As Long

' basBatchWizard
Public Sub BuildPreview(scope As String, BaseToken As String, SystemLetter As String)
Public Function CommitBatch(scope, BaseToken, SystemLetter, CustomerCode, CustomerName, _
                            PONumber, DateReceived, ByRef createdCount, ByRef batchID) As Boolean

' basResetRoutine
Public Sub Reset_Qty_From_Qualifiers()
```

---

## Completion Checklist

- [ ] Task 1: Form layout designed
- [ ] Task 2: Order type toggle working
- [ ] Task 3: Project combo working
- [ ] Task 4: Customer combo working
- [ ] Task 5: Form opens correctly
- [ ] Task 6: Preview button working
- [ ] Task 7: Commit button working
- [ ] Task 8: Cancel/cleanup working
- [ ] Task 9: New Batch button on frmOrderList
- [ ] Task 10: Validation helper complete
- [ ] End-to-end test: SALES flow
- [ ] End-to-end test: PROJECT flow
