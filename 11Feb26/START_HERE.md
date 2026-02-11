# START HERE - Daily Work Guide

**Date Created:** January 23, 2026
**Project:** Q1019 Order Management System
**Database:** `Q1019_23Jan26_Start_travellaptop.accdb`

---

## Before You Begin

Make sure you have:
- [ ] Access database open
- [ ] Microsoft Copilot open in browser
- [ ] This repo pulled to latest (`git pull origin main`)

---

## Today's Goal

Enable batch order generation for both SALES and PROJECT orders with unified order number format:

```
<BaseToken>-<QualifierCode><SystemLetter><Seq>-<BackorderNo>

Examples:
  SALES:   576001-CMN001-00
  PROJECT: TAGO25-CMN001-00
```

---

## Step-by-Step Checklist

### Phase 1: Fix Known Bug (Do This First!)

- [ ] **1.1** Open Access database
- [ ] **1.2** In Navigation Pane, find query: `frmOrderList`
- [ ] **1.3** Open in SQL View
- [ ] **1.4** Replace with this SQL:

```sql
SELECT S.SOID, S.OrderNumber, S.CustomerName, S.PONumber, S.DateReceived,
       F.QualifierCode, F.SequenceNo, S.BaseToken, S.SystemLetter,
       S.BackorderNo, S.BatchID
FROM SalesOrders AS S
LEFT JOIN qryFirstQualifierPerSO AS F ON S.SOID = F.SOID
WHERE S.ActiveFlag = True
ORDER BY S.SOID DESC;
```

- [ ] **1.5** Save and close the query
- [ ] **1.6** Test: Open `frmOrderList`, verify no errors

---

### Phase 2: Set Up Copilot Session

- [ ] **2.1** Open Microsoft Copilot
- [ ] **2.2** Start a new conversation
- [ ] **2.3** Copy the **CONTEXT MESSAGE** from `COPILOT_GUIDE.md` (lines 84-118)
- [ ] **2.4** Paste into Copilot and send

**Quick copy - Context Message:**
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

### Phase 3: Implement Tasks (One at a Time)

Work through these tasks in order. For each task:
1. Copy the task prompt from `COPILOT_GUIDE.md`
2. Paste to Copilot
3. Get the code
4. Implement in Access
5. Test
6. Check off and move to next

| # | Task | File Reference | Status |
|---|------|----------------|--------|
| 1 | Form layout design | COPILOT_GUIDE.md lines 132-153 | [ ] |
| 2 | Order type toggle (SALES/PROJECT) | COPILOT_GUIDE.md lines 159-173 | [ ] |
| 3 | Project combo box | COPILOT_GUIDE.md lines 179-195 | [ ] |
| 4 | Customer combo box | COPILOT_GUIDE.md lines 201-211 | [ ] |
| 5 | Form_Open event | COPILOT_GUIDE.md lines 217-230 | [ ] |
| 6 | Preview button | COPILOT_GUIDE.md lines 236-248 | [ ] |
| 7 | Commit button | COPILOT_GUIDE.md lines 254-268 | [ ] |
| 8 | Cancel/cleanup | COPILOT_GUIDE.md lines 274-285 | [ ] |
| 9 | New Batch button on frmOrderList | COPILOT_GUIDE.md lines 291-298 | [ ] |
| 10 | Validation helper | COPILOT_GUIDE.md lines 304-316 | [ ] |

---

### Phase 4: Testing

- [ ] **4.1** Test SALES flow:
  - Open frmOrderList
  - Click New Batch
  - Select SALES → BaseToken auto-generates (576xxx)
  - Select System Letter, Customer, enter PO#
  - Enter qualifier quantities
  - Preview Numbers → verify format
  - Commit Batch → returns to filtered list

- [ ] **4.2** Test PROJECT flow:
  - Open frmOrderList
  - Click New Batch
  - Select PROJECT → Select from Projects dropdown
  - Complete remaining fields
  - Preview and Commit

- [ ] **4.3** Verify no "Enter Parameter Value" prompts

---

## Quick Reference Files

| File | What It Contains |
|------|------------------|
| `START_HERE.md` | You are here! Daily checklist |
| `COPILOT_GUIDE.md` | Context message + all task prompts |
| `USER_FLOW.md` | Visual diagram of the user flow |
| `INSTRUCTIONS_SALES_BATCH.md` | Detailed VBA code examples |
| `RECOMMENDED_CLEANUP.md` | Cleanup tasks + planned changes |
| `numberResContext` | Full database context export |
| `modAI_ContextExport.bas` | VBA module to regenerate context |

---

## If You Get Stuck

### Copilot not understanding?
- Paste the error message exactly
- Say: "That didn't work. I got this error: [paste error]"
- Ask: "Can you explain what this code does?"

### Need to regenerate context?
In Access Immediate Window (Ctrl+G):
```
Call ExportAIContext
```
This creates a fresh context file in the database folder.

### Need to see current database state?
Open `numberResContext` in the repo - it has full schema, queries, forms, and code.

---

## End of Day

- [ ] Test both SALES and PROJECT flows work
- [ ] Commit any working code to a backup
- [ ] Note any issues for tomorrow in this file or a new issue

---

## Notes / Issues Found Today

_(Add notes here as you work)_

```
-
-
-
```
