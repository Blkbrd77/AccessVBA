# User Flow - Order Management System

Based on review of `numberResContext` export from `Q1019_23Jan26_Start_travellaptop.accdb`.

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           ORDER MANAGEMENT SYSTEM                            │
└─────────────────────────────────────────────────────────────────────────────┘

┌─────────────────┐
│  frmOrderList   │ ◄─── Main entry point: View all active orders
│  (Order List)   │      - Filtered by BatchID after batch creation
└────────┬────────┘      - Shows: OrderNumber, Customer, PO#, Date, Qualifier
         │
         │ [New Order]
         ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                         frmSalesOrderEntry                                   │
│                      (Sales Order Entry Form)                                │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  ┌─── HEADER SECTION ───────────────────────────────────────────────────┐   │
│  │  Order Type:    [SALES ▼] / [PROJECT ▼]                              │   │
│  │  Order Base:    [txtBaseToken]  ◄── 576001 (SALES) or TAGO25 (PROJ)  │   │
│  │  SysLtr:        [N ▼] [P ▼] [J ▼]  ◄── From SystemLetter table       │   │
│  │  Customer:      [cboCustomerCode ▼]  ◄── Lookup from Customers       │   │
│  │  Customer Name: [txtCustomerName]    ◄── Auto-fills from combo       │   │
│  │  PO #:          [txtPONumber]                                        │   │
│  │  Date Received: [txtDateReceived]                                    │   │
│  └──────────────────────────────────────────────────────────────────────┘   │
│                                                                              │
│  ┌─── QUALIFIER QTY SUBFORM (fsubQualifierQty) ─────────────────────────┐   │
│  │  QualifierCode │ Description        │ Qty                            │   │
│  │  ──────────────┼────────────────────┼─────                           │   │
│  │  CM            │ Construction Mgmt  │ [5]                            │   │
│  │  CE            │ Civil Engineering  │ [3]                            │   │
│  │  EL            │ Electrical         │ [0]                            │   │
│  │  ...           │ ...                │ ...                            │   │
│  └──────────────────────────────────────────────────────────────────────┘   │
│                                                                              │
│  ┌─── PREVIEW LIST (lstPreview) ────────────────────────────────────────┐   │
│  │  TAGO25-CMN001-00                                                    │   │
│  │  TAGO25-CMN002-00                                                    │   │
│  │  TAGO25-CMN003-00                                                    │   │
│  │  TAGO25-CEN001-00                                                    │   │
│  │  ...                                                                 │   │
│  └──────────────────────────────────────────────────────────────────────┘   │
│                                                                              │
│                    ┌──────────────┐  ┌─────────────────┐                    │
│                    │ Reset Qty    │  │ Preview Numbers │                    │
│                    └──────────────┘  └────────┬────────┘                    │
│                    ┌──────────────┐           │                             │
│                    │ Commit Batch │ ◄─────────┘                             │
│                    └──────┬───────┘                                         │
│                    ┌──────────────┐                                         │
│                    │    Cancel    │                                         │
│                    └──────────────┘                                         │
└─────────────────────────────────────────────────────────────────────────────┘
         │
         │ [Commit Batch]
         ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                           BATCH COMMIT PROCESS                               │
│                           (basBatchWizard)                                   │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  1. Generate BatchID (GUID)                                                  │
│  2. Begin Transaction                                                        │
│  3. For each Qualifier with Qty > 0:                                         │
│     └─► For i = 1 to Qty:                                                    │
│         ├─► ReserveSeq() ─► Get next sequence from OrderSeq table            │
│         ├─► Build OrderNumber: BaseToken-QualifierCode+SysLtr+Seq-00         │
│         ├─► INSERT into SalesOrders (header)                                 │
│         └─► INSERT into SalesOrderEntry (qualifier detail)                   │
│  4. Commit Transaction                                                       │
│  5. Show success message with count                                          │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
         │
         │ [Success]
         ▼
┌─────────────────┐
│  frmOrderList   │ ◄─── Returns here, filtered to show new batch
│  (Filtered)     │      Filter: BatchID = '{new-guid}'
└─────────────────┘


┌─────────────────────────────────────────────────────────────────────────────┐
│                         SUPPORTING DIALOGS                                   │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  dlgProjectSelection          dlgNewOrderType (usage unclear)                │
│  ┌────────────────────┐       ┌────────────────────┐                        │
│  │ Select Project:    │       │ [May be unused]    │                        │
│  │ [TAGO25 - Desc ▼]  │       │                    │                        │
│  │                    │       │                    │                        │
│  │ [OK] [Cancel]      │       │                    │                        │
│  └────────────────────┘       └────────────────────┘                        │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘


┌─────────────────────────────────────────────────────────────────────────────┐
│                         DATA FLOW                                            │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  QualifierType ──► tmpQualifierQty ──► User enters Qty ──► tmpBatchPreview   │
│       │                                                           │          │
│       │ (seed descriptions)              (preview only)           │          │
│       ▼                                                           ▼          │
│  ┌─────────┐     ┌──────────┐     ┌─────────────┐     ┌─────────────────┐   │
│  │OrderSeq │ ◄───│ Reserve  │ ◄───│ Commit      │ ───►│ SalesOrders     │   │
│  │ (Next#) │     │ Sequence │     │ Batch       │     │ SalesOrderEntry │   │
│  └─────────┘     └──────────┘     └─────────────┘     └─────────────────┘   │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

## Flow Summary

1. **View Orders** → `frmOrderList`
2. **Create Batch** → `frmSalesOrderEntry`
   - Fill header (Type, Base, SysLtr, Customer, PO, Date)
   - Enter quantities per qualifier code
   - Click "Preview Numbers" to see what will be created
   - Click "Commit Batch" to create orders
3. **Behind the scenes**:
   - `basSeqAllocator.ReserveSeq()` reserves sequence numbers
   - `basBatchWizard.CommitBatch()` creates orders in a transaction
4. **Return** → `frmOrderList` filtered to show new batch

## Key Tables

| Table | Purpose |
|-------|---------|
| `SalesOrders` | Order header (OrderNumber, Customer, PO, Dates) |
| `SalesOrderEntry` | Qualifier details per order |
| `OrderSeq` | Sequence counters per domain (Scope/Base/Qualifier/SysLtr) |
| `QualifierType` | Master list of qualifier codes (CM, CE, EL, etc.) |
| `Customers` | Customer lookup |
| `Projects` | Project base tokens (TAGO25, etc.) |
| `SystemLetter` | System letter codes (N, P, J, etc.) |
| `tmpQualifierQty` | Temp table for batch entry quantities |
| `tmpBatchPreview` | Temp table for preview display |

## Key Modules

| Module | Purpose |
|--------|---------|
| `basSeqAllocator` | Transactional sequence reservation |
| `basBatchWizard` | Preview and commit batch operations |
| `basOrderNumberGen` | Order number string building |
| `basResetRoutine` | Reset qualifier quantities from master |
