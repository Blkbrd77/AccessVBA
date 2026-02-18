-- =============================================================================
-- Migration: Add DateInvoiceSent column to SalesOrders
-- Purpose  : Tracks the date a proforma invoice was sent to the customer.
--            Mirrors the existing DateBilled column pattern.
-- Run on   : Access BE (.accdb) via the Access query designer (DDL mode)
--            OR via SSMA / Azure SQL after migration.
-- =============================================================================

-- ACCESS (.accdb) - run in Access Query Designer (set query type to Data Definition)
ALTER TABLE SalesOrders
    ADD COLUMN DateInvoiceSent DATETIME;


-- =============================================================================
-- AZURE SQL equivalent (run after SSMA migration)
-- =============================================================================
-- ALTER TABLE dbo.SalesOrders
--     ADD DateInvoiceSent DATE NULL;


-- =============================================================================
-- Optional: tblOrderAudit already handles STAMP_INVOICE_SENT events via
-- LogOrderAction(), so no schema change is needed to the audit table.
-- =============================================================================
