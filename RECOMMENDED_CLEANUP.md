# Recommended Cleanup Actions

Based on review of `numberResContext` export from `Q1019_23Jan26_Start_travellaptop.accdb`.

## Unused Tables

| Table | Rows | Issue |
|-------|------|-------|
| `SalesOrderQualifier` | 0 | **Superseded** - Appears to be an older design replaced by `SalesOrderEntry`. Has relationships defined but no data. |
| `SalesOrderQualifierAudit` | 0 | **Orphaned** - Audit table for `SalesOrderQualifier` which isn't used. |

Both tables have relationships and indexes defined but zero rows, suggesting they were part of an earlier architecture that was replaced.

**Action**: Drop these tables after confirming no code references them.

---

## Redundant/Unnecessary Modules

| Module | Issue |
|--------|-------|
| `Module1` | **Generic name** - Likely scratch/test code. Should be renamed or removed. |
| `CoPilotContextDump` | **Superseded** - The original context dump; replaced by `modAI_ContextExport`. |
| `basSchemaDump` | **Redundant** - Schema dumping now handled by context export module. |
| `basFormDump` | **Redundant** - Form dumping now handled by context export module. |
| `basOrderNumberGen` | **Partially redundant** - Scans `OrderNumber` strings to find max sequence. `basSeqAllocator` uses `OrderSeq` table instead. These represent two different approaches - pick one. |
| `basTrace` | **Minimal utility** - Only 2 functions (`DEBUG_UI` flag and `Trace`). Could be inlined or removed if not actively used. |
| `basSalesOrdersIndexes` | **One-time scripts** - Contains index patches (`Allow_Duplicate_PONumbers`, `Drop_Blocking_Composite_For_Project`, etc.). These are maintenance scripts that should be run once and removed. |

**Action**: Delete `Module1`, `CoPilotContextDump`, `basSchemaDump`, `basFormDump`. Archive `basSalesOrdersIndexes`. Consolidate sequence generation approach.

---

## Potential Form Issues

| Form | Issue |
|------|-------|
| `dlgNewOrderType` | **Verify usage** - Check if this is actually called from anywhere or if functionality was merged into `frmSalesOrderEntry`. |

**Action**: Search codebase for references to `dlgNewOrderType`. Remove if unused.

---

## Summary of Recommended Actions

### Priority 1: Safe to Remove
- [ ] Delete `Module1` (generic scratch module)
- [ ] Delete `CoPilotContextDump` (replaced by `modAI_ContextExport`)
- [ ] Delete `basSchemaDump` (redundant with context export)
- [ ] Delete `basFormDump` (redundant with context export)

### Priority 2: Verify Before Removing
- [ ] Verify `dlgNewOrderType` is not referenced, then remove
- [ ] Verify `SalesOrderQualifier` table is not referenced in code, then drop
- [ ] Verify `SalesOrderQualifierAudit` table is not referenced in code, then drop

### Priority 3: Architectural Decisions
- [ ] Choose between `basOrderNumberGen` (string scanning) and `basSeqAllocator` (OrderSeq table) for sequence management
- [ ] Decide whether to keep `basTrace` or inline its minimal functionality
- [ ] Archive `basSalesOrdersIndexes` (one-time maintenance scripts)

---

## Planned Changes

### Unify SALES and PROJECT Order Number Formats

**Current State:**

| Order Type | Format | Example |
|------------|--------|---------|
| SALES | `<BaseToken>-<SystemLetter><Seq>-<BackorderNo>` | `576001-P001-00` |
| PROJECT | `<BaseToken>-<QualifierCode><SystemLetter><Seq>-<BackorderNo>` | `TAGO25-CMN001-00` |

**Desired State:**

| Order Type | Format | Example |
|------------|--------|---------|
| SALES | `<BaseToken>-<QualifierCode><SystemLetter><Seq>-<BackorderNo>` | `576001-CMN001-00` |
| PROJECT | `<BaseToken>-<QualifierCode><SystemLetter><Seq>-<BackorderNo>` | `TAGO25-CMN001-00` |

**Requirements:**
- SALES order numbers should match PROJECT order number format (include QualifierCode)
- SALES BaseTokens continue to use year-based numbering scheme:
  - 2026: starts at 576000
  - 2027: starts at 577000
  - Pattern: `576000 + (year - 2026) * 1000`

**Affected Code:**
- `basOrderNumberGen.BuildSalesOrderNumber` - needs QualifierCode parameter added
- `basOrderNumberGen.NextSeqForSales` - needs to account for qualifier in sequence lookup
- `basBatchWizard.CreateOneOrder` - may need updates for SALES flow
- `frmSalesOrderEntry` - UI needs to capture QualifierCode for SALES orders

---

## Notes

- The `basSeqAllocator` approach using the `OrderSeq` table is more robust for concurrent multi-user access
- The `basOrderNumberGen` approach of scanning existing `OrderNumber` strings is simpler but less efficient and prone to race conditions
- Consider keeping one backup of removed modules before deletion in case functionality needs to be recovered
