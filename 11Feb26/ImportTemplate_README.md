# Comprehensive Data Import Templates

These CSV templates allow you to import data into the database with ALL fields, including the new features not covered by LegacyDataImport.

## Files Included

| File | Target Table | Purpose |
|------|--------------|---------|
| `ImportTemplate_SalesOrders.csv` | SalesOrders | Main order headers |
| `ImportTemplate_SalesOrderEntry.csv` | SalesOrderEntry | Qualifier details per order |
| `ImportTemplate_OrderSeq.csv` | OrderSeq | Sequence tracking (required for future order creation) |

---

## ImportTemplate_SalesOrders.csv

### Field Definitions

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| **OrderNumber** | Text(50) | YES | Full derived order number (e.g., "576001-CMN001-00"). Must be UNIQUE. |
| **OrderType** | Text(10) | YES | "SALES" or "PROJECT" |
| **BaseToken** | Text(50) | YES | For SALES: numeric token (576001). For PROJECT: project code (TAGO25) |
| **SystemLetter** | Text(1) | YES | Must exist in SystemLetter table (P, N, J, etc.) |
| **BackorderNo** | Integer | YES | Last segment of order number (0 for original, 1+ for backorders) |
| **CustomerName** | Text(255) | YES | Customer name |
| **CustomerCode** | Text(255) | NO | Customer code (should match Customers table) |
| **PONumber** | Text(55) | NO | Purchase order number |
| **DateReceived** | Date | NO | Order received date (YYYY-MM-DD format) |
| **DateBilled** | Date | NO | Date billed |
| **ActiveFlag** | Boolean | NO | True/False (defaults to True) |
| **BatchID** | Text(36) | NO | GUID linking orders created in same batch |
| **DateCreated** | Date | NO | Order creation timestamp |
| **DateBackorderCreated** | Date | NO | When backorder was created (for BackorderNo > 0) |
| **DateCanceled** | Date | NO | Cancellation date |
| **CancelReason** | Text(255) | NO | Why order was canceled |
| **CanceledBy** | Text(64) | NO | User who canceled |

### New Fields Not in LegacyDataImport

These fields are available in the new system but were not imported by the legacy process:

- `CustomerCode` - Links to Customers table
- `BackorderNo` - Was parsed from DerivedOrderID, now explicit
- `BatchID` - For tracking batch-created orders
- `DateBackorderCreated` - Backorder creation date
- `DateCanceled`, `CancelReason`, `CanceledBy` - Cancellation tracking

---

## ImportTemplate_SalesOrderEntry.csv

### Field Definitions

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| **SOID** | Long | YES | Foreign key to SalesOrders. Must match an imported order. |
| **QualifierCode** | Text(10) | YES | Qualifier code (CMN, CEN, EL, etc.). Must exist in QualifierType table. |
| **SequenceNo** | Integer | NO | Sequence within this qualifier for this order (defaults to 0) |
| **OrderNumberDisplay** | Text(30) | YES | The derived order number (e.g., "576001-CMN001-00") |
| **IsDeleted** | Boolean | NO | Soft delete flag (defaults to False) |
| **CreatedOn** | Date | NO | When entry was created |

### Notes
- The SOID field references SalesOrders - you'll need to know the SOID values after importing SalesOrders
- For initial import, you can use row numbers (1, 2, 3...) matching the order in SalesOrders.csv

---

## ImportTemplate_OrderSeq.csv

### Field Definitions

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| **Scope** | Text(10) | YES | "SALES" or "PROJECT" |
| **BaseToken** | Text(32) | YES | The base token (576001 for SALES, TAGO25 for PROJECT) |
| **QualifierCode** | Text(10) | YES | Qualifier code |
| **SystemLetter** | Text(1) | YES | System letter |
| **NextSeq** | Long | YES | Next sequence number to allocate |
| **LastUpdated** | Date | NO | Last update timestamp |

### Why This Table Matters

The OrderSeq table tracks sequence numbers. If you import orders but don't populate OrderSeq, the system won't know what sequence numbers are already used and may create duplicates.

**Example:** If you import order "576001-CMN001-00" (sequence 1), you must set OrderSeq.NextSeq = 2 for (SALES, 576001, CMN, P) so the next order gets sequence 2.

---

## Import Process

### Step 1: Prepare Your Data

1. Open each CSV in Excel
2. Delete the example data rows (keep headers)
3. Populate with your actual data
4. Ensure dates are in YYYY-MM-DD format
5. Ensure Boolean fields are True/False

### Step 2: Import into Access

**Option A: Using Access Import Wizard**

1. Open Access database
2. External Data > Text File > Browse to CSV
3. Select "Append to existing table" or "Import to new table"
4. Follow wizard to map fields

**Option B: Using VBA Import Module**

Create a new module with this code:

```vba
Public Sub ImportFromCSV()
    ' Import SalesOrders
    DoCmd.TransferText acImportDelim, , "SalesOrders", _
        "C:\Path\To\ImportTemplate_SalesOrders.csv", True

    ' Import SalesOrderEntry
    DoCmd.TransferText acImportDelim, , "SalesOrderEntry", _
        "C:\Path\To\ImportTemplate_SalesOrderEntry.csv", True

    ' Import OrderSeq
    DoCmd.TransferText acImportDelim, , "OrderSeq", _
        "C:\Path\To\ImportTemplate_OrderSeq.csv", True

    MsgBox "Import complete!", vbInformation
End Sub
```

### Step 3: Post-Import Verification

Run these queries to verify data integrity:

```sql
-- Check for orphaned SalesOrderEntry records
SELECT * FROM SalesOrderEntry
WHERE SOID NOT IN (SELECT SOID FROM SalesOrders);

-- Check for missing OrderSeq entries
SELECT DISTINCT s.BaseToken, e.QualifierCode, s.SystemLetter
FROM SalesOrders s
INNER JOIN SalesOrderEntry e ON s.SOID = e.SOID
WHERE NOT EXISTS (
    SELECT 1 FROM OrderSeq q
    WHERE q.BaseToken = s.BaseToken
    AND q.QualifierCode = e.QualifierCode
    AND q.SystemLetter = s.SystemLetter
);

-- Verify order counts
SELECT 'SalesOrders' AS TableName, COUNT(*) AS Records FROM SalesOrders
UNION ALL
SELECT 'SalesOrderEntry', COUNT(*) FROM SalesOrderEntry
UNION ALL
SELECT 'OrderSeq', COUNT(*) FROM OrderSeq;
```

---

## Reference Data

### Valid SystemLetter Values
| Letter | Description |
|--------|-------------|
| P | Primary |
| N | Secondary |
| J | Project |

### Valid QualifierCode Values
| Code | Name |
|------|------|
| CMN | Construction Mgmt |
| CEN | Central Engineering |
| EL | Electrical |
| ... | (check QualifierType table for full list) |

### OrderType Values
| Type | BaseToken Format |
|------|------------------|
| SALES | Numeric (576001, 576002, etc.) |
| PROJECT | Alpha code (TAGO25, etc.) |

---

## Differences from LegacyDataImport

| Feature | LegacyDataImport | This Template |
|---------|------------------|---------------|
| CustomerCode | Not imported | Supported |
| BackorderNo | Parsed from DerivedOrderID | Explicit field |
| BatchID | Not imported | Supported |
| DateBackorderCreated | Not imported | Supported |
| Cancellation fields | Not imported | Supported |
| OrderSeq population | Not handled | Included template |
| Notes field | Lost during import | Still not in schema |

---

## Tips

1. **SOID Assignment**: Access auto-generates SOID values. For SalesOrderEntry, import SalesOrders first, then query to get the assigned SOID values before importing entries.

2. **Batch Import**: If importing many orders from the same batch, generate a GUID and use it for all orders in that batch.

3. **BackorderNo Logic**:
   - Original order: BackorderNo = 0
   - First backorder: BackorderNo = 1
   - The order number format includes this: `576001-CMN001-00` vs `576001-CMN001-01`

4. **Sequence Numbers**: After import, verify OrderSeq.NextSeq is set to MAX(existing sequence) + 1 for each combination.
