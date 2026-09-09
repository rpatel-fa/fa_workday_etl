-- ============================================================================
-- VALIDATION RESULTS TABLE
-- ============================================================================
DROP TABLE IF EXISTS validation_supplier_result
;
CREATE TABLE validation_supplier_result (
    validation_type VARCHAR
  , table_name      VARCHAR
  , data_count      BIGINT
)
;

-- ============================================================================
-- INGESTION SOURCE COUNTS (Raw tables from each workbook/sheet)
-- ============================================================================
-- FA workbook raw tables
INSERT INTO validation_supplier_result
SELECT 'ingestion_raw', 'fin_supplier_creditor_created_date_fa', COUNT(*) FROM fin_supplier_creditor_created_date_fa
;
INSERT INTO validation_supplier_result
SELECT 'ingestion_raw', 'fin_supplier_creditor_last_payment_date_fa', COUNT(*) FROM fin_supplier_creditor_last_payment_date_fa
;
INSERT INTO validation_supplier_result
SELECT 'ingestion_raw', 'fin_supplier_creditor_last_purchase_date_fa', COUNT(*) FROM fin_supplier_creditor_last_purchase_date_fa
;
-- NFC workbook raw tables
INSERT INTO validation_supplier_result
SELECT 'ingestion_raw', 'fin_supplier_creditor_created_date_nfc', COUNT(*) FROM fin_supplier_creditor_created_date_nfc
;
INSERT INTO validation_supplier_result
SELECT 'ingestion_raw', 'fin_supplier_creditor_last_payment_date_nfc', COUNT(*) FROM fin_supplier_creditor_last_payment_date_nfc
;
INSERT INTO validation_supplier_result
SELECT 'ingestion_raw', 'fin_supplier_creditor_last_purchase_date_nfc', COUNT(*) FROM fin_supplier_creditor_last_purchase_date_nfc
;
-- WNSL workbook raw tables
INSERT INTO validation_supplier_result
SELECT 'ingestion_raw', 'fin_supplier_creditor_created_date_wnsl', COUNT(*) FROM fin_supplier_creditor_created_date_wnsl
;
INSERT INTO validation_supplier_result
SELECT 'ingestion_raw', 'fin_supplier_creditor_last_payment_date_wnsl', COUNT(*) FROM fin_supplier_creditor_last_payment_date_wnsl
;
INSERT INTO validation_supplier_result
SELECT 'ingestion_raw', 'fin_supplier_creditor_last_purchase_date_wnsl', COUNT(*) FROM fin_supplier_creditor_last_purchase_date_wnsl
;

-- ============================================================================
-- INGESTION TOTALS BY BUSINESS UNIT (Sum of raw tables per workbook)
-- ============================================================================
INSERT INTO validation_supplier_result
SELECT 'ingestion_total', 'FA_workbook_total',
       (SELECT COUNT(*) FROM fin_supplier_creditor_created_date_fa) +
       (SELECT COUNT(*) FROM fin_supplier_creditor_last_payment_date_fa) +
       (SELECT COUNT(*) FROM fin_supplier_creditor_last_purchase_date_fa)
;
INSERT INTO validation_supplier_result
SELECT 'ingestion_total', 'NFC_workbook_total',
       (SELECT COUNT(*) FROM fin_supplier_creditor_created_date_nfc) +
       (SELECT COUNT(*) FROM fin_supplier_creditor_last_payment_date_nfc) +
       (SELECT COUNT(*) FROM fin_supplier_creditor_last_purchase_date_nfc)
;
INSERT INTO validation_supplier_result
SELECT 'ingestion_total', 'WNSL_workbook_total',
       (SELECT COUNT(*) FROM fin_supplier_creditor_created_date_wnsl) +
       (SELECT COUNT(*) FROM fin_supplier_creditor_last_payment_date_wnsl) +
       (SELECT COUNT(*) FROM fin_supplier_creditor_last_purchase_date_wnsl)
;

-- ============================================================================
-- MERGED TABLE COUNTS BY BUSINESS UNIT
-- ============================================================================
INSERT INTO validation_supplier_result
SELECT 'merged_by_bu', 'src_fin_supplier_FA', COUNT(*) FROM src_fin_supplier WHERE business_unit = 'FA'
;
INSERT INTO validation_supplier_result
SELECT 'merged_by_bu', 'src_fin_supplier_NFC', COUNT(*) FROM src_fin_supplier WHERE business_unit = 'NFC'
;
INSERT INTO validation_supplier_result
SELECT 'merged_by_bu', 'src_fin_supplier_WNSL', COUNT(*) FROM src_fin_supplier WHERE business_unit = 'WNSL'
;

-- Source table count (total)
INSERT INTO validation_supplier_result
SELECT 'record_count', 'src_fin_supplier', COUNT(*) FROM src_fin_supplier
;

-- Record counts by workday table
INSERT INTO validation_supplier_result
SELECT 'record_count', 'workday_supplier_name', COUNT(*) FROM workday_supplier_name
;
INSERT INTO validation_supplier_result
SELECT 'record_count', 'workday_supplier_groups', COUNT(*) FROM workday_supplier_groups
;
INSERT INTO validation_supplier_result
SELECT 'record_count', 'workday_supplier_alternate_name', COUNT(*) FROM workday_supplier_alternate_name
;
INSERT INTO validation_supplier_result
SELECT 'record_count', 'workday_supplier_currencies', COUNT(*) FROM workday_supplier_currencies
;
INSERT INTO validation_supplier_result
SELECT 'record_count', 'workday_supplier_tax', COUNT(*) FROM workday_supplier_tax
;
INSERT INTO validation_supplier_result
SELECT 'record_count', 'workday_supplier_payment', COUNT(*) FROM workday_supplier_payment
;
INSERT INTO validation_supplier_result
SELECT 'record_count', 'workday_supplier_procurement', COUNT(*) FROM workday_supplier_procurement
;
INSERT INTO validation_supplier_result
SELECT 'record_count', 'workday_supplier_company_restrictions', COUNT(*) FROM workday_supplier_company_restrictions
;
INSERT INTO validation_supplier_result
SELECT 'record_count', 'workday_supplier_settlement_account', COUNT(*) FROM workday_supplier_settlement_account
;
INSERT INTO validation_supplier_result
SELECT 'record_count', 'workday_supplier_intermediary_account', COUNT(*) FROM workday_supplier_intermediary_account
;
INSERT INTO validation_supplier_result
SELECT 'record_count', 'workday_supplier_address', COUNT(*) FROM workday_supplier_address
;
INSERT INTO validation_supplier_result
SELECT 'record_count', 'workday_supplier_phone', COUNT(*) FROM workday_supplier_phone
;
INSERT INTO validation_supplier_result
SELECT 'record_count', 'workday_supplier_email', COUNT(*) FROM workday_supplier_email
;

-- Data quality checks
INSERT INTO validation_supplier_result
SELECT 'data_quality', 'orphan_groups', COUNT(*)
  FROM workday_supplier_groups g
  LEFT JOIN workday_supplier_name n ON g.supplier_id = n.supplier_id
 WHERE n.supplier_id IS NULL
;
INSERT INTO validation_supplier_result
SELECT 'data_quality', 'missing_required_supplier_name', COUNT(*)
  FROM workday_supplier_name
 WHERE supplier_id IS NULL OR supplier_name IS NULL
;
INSERT INTO validation_supplier_result
SELECT 'data_quality', 'missing_required_currencies', COUNT(*)
  FROM workday_supplier_currencies
 WHERE supplier_id IS NULL OR accepted_currencies_plus IS NULL
;
INSERT INTO validation_supplier_result
SELECT 'data_quality', 'invalid_email_addresses', COUNT(*)
  FROM src_fin_supplier
 WHERE email_to_address IS NOT NULL
   AND email_to_address !~ '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'
;

-- ============================================================================
-- INGESTION RECONCILIATION SUMMARY
-- This compares row counts from raw ingested tables vs merged table by business unit
-- Note: Merged counts may differ due to DISTINCT and deduplication logic
-- ============================================================================
DROP TABLE IF EXISTS validation_supplier_reconciliation
;
CREATE TABLE validation_supplier_reconciliation AS
WITH ingestion_totals AS (
    SELECT 'FA' AS business_unit,
           (SELECT COUNT(*) FROM fin_supplier_creditor_created_date_fa) +
           (SELECT COUNT(*) FROM fin_supplier_creditor_last_payment_date_fa) +
           (SELECT COUNT(*) FROM fin_supplier_creditor_last_purchase_date_fa) AS ingested_rows
    UNION ALL
    SELECT 'NFC' AS business_unit,
           (SELECT COUNT(*) FROM fin_supplier_creditor_created_date_nfc) +
           (SELECT COUNT(*) FROM fin_supplier_creditor_last_payment_date_nfc) +
           (SELECT COUNT(*) FROM fin_supplier_creditor_last_purchase_date_nfc) AS ingested_rows
    UNION ALL
    SELECT 'WNSL' AS business_unit,
           (SELECT COUNT(*) FROM fin_supplier_creditor_created_date_wnsl) +
           (SELECT COUNT(*) FROM fin_supplier_creditor_last_payment_date_wnsl) +
           (SELECT COUNT(*) FROM fin_supplier_creditor_last_purchase_date_wnsl) AS ingested_rows
),
merged_totals AS (
    SELECT business_unit, COUNT(*) AS merged_rows
      FROM src_fin_supplier
     GROUP BY business_unit
)
SELECT i.business_unit
     , i.ingested_rows
     , COALESCE(m.merged_rows, 0) AS merged_rows
     , i.ingested_rows - COALESCE(m.merged_rows, 0) AS deduped_rows
     , CASE
         WHEN COALESCE(m.merged_rows, 0) > 0 THEN 'OK - Data loaded'
         WHEN i.ingested_rows = 0 THEN 'OK - No source data'
         ELSE 'ERROR - Missing data'
       END AS status
  FROM ingestion_totals i
  LEFT JOIN merged_totals m ON i.business_unit = m.business_unit
 ORDER BY i.business_unit
;
