-- ============================================================================
-- VALIDATION RESULTS TABLE
-- ============================================================================
DROP TABLE IF EXISTS validation_customer_result
;
CREATE TABLE validation_customer_result (
    validation_type VARCHAR
  , table_name      VARCHAR
  , data_count      BIGINT
)
;

-- ============================================================================
-- INGESTION SOURCE COUNTS (Raw tables from each workbook/sheet)
-- ============================================================================
-- FA workbook raw tables
INSERT INTO validation_customer_result
SELECT 'ingestion_raw', 'fin_customer_debtor_last_3_years_fa', COUNT(*) FROM fin_customer_debtor_last_3_years_fa
;
INSERT INTO validation_customer_result
SELECT 'ingestion_raw', 'fin_customer_debtor_created_date_fa', COUNT(*) FROM fin_customer_debtor_created_date_fa
;
INSERT INTO validation_customer_result
SELECT 'ingestion_raw', 'fin_customer_debtor_last_payment_date_fa', COUNT(*) FROM fin_customer_debtor_last_payment_date_fa
;
-- NFC workbook raw tables
INSERT INTO validation_customer_result
SELECT 'ingestion_raw', 'fin_customer_debtor_last_3_years_nfc', COUNT(*) FROM fin_customer_debtor_last_3_years_nfc
;
INSERT INTO validation_customer_result
SELECT 'ingestion_raw', 'fin_customer_debtor_created_date_nfc', COUNT(*) FROM fin_customer_debtor_created_date_nfc
;
INSERT INTO validation_customer_result
SELECT 'ingestion_raw', 'fin_customer_debtor_last_payment_date_nfc', COUNT(*) FROM fin_customer_debtor_last_payment_date_nfc
;
-- WNSL workbook raw tables
INSERT INTO validation_customer_result
SELECT 'ingestion_raw', 'fin_customer_debtor_last_3_years_wnsl', COUNT(*) FROM fin_customer_debtor_last_3_years_wnsl
;
INSERT INTO validation_customer_result
SELECT 'ingestion_raw', 'fin_customer_debtor_created_date_wnsl', COUNT(*) FROM fin_customer_debtor_created_date_wnsl
;
INSERT INTO validation_customer_result
SELECT 'ingestion_raw', 'fin_customer_debtor_last_payment_date_wnsl', COUNT(*) FROM fin_customer_debtor_last_payment_date_wnsl
;

-- ============================================================================
-- INGESTION TOTALS BY BUSINESS UNIT (Sum of raw tables per workbook)
-- ============================================================================
INSERT INTO validation_customer_result
SELECT 'ingestion_total', 'FA_workbook_total',
       (SELECT COUNT(*) FROM fin_customer_debtor_last_3_years_fa) +
       (SELECT COUNT(*) FROM fin_customer_debtor_created_date_fa) +
       (SELECT COUNT(*) FROM fin_customer_debtor_last_payment_date_fa)
;
INSERT INTO validation_customer_result
SELECT 'ingestion_total', 'NFC_workbook_total',
       (SELECT COUNT(*) FROM fin_customer_debtor_last_3_years_nfc) +
       (SELECT COUNT(*) FROM fin_customer_debtor_created_date_nfc) +
       (SELECT COUNT(*) FROM fin_customer_debtor_last_payment_date_nfc)
;
INSERT INTO validation_customer_result
SELECT 'ingestion_total', 'WNSL_workbook_total',
       (SELECT COUNT(*) FROM fin_customer_debtor_last_3_years_wnsl) +
       (SELECT COUNT(*) FROM fin_customer_debtor_created_date_wnsl) +
       (SELECT COUNT(*) FROM fin_customer_debtor_last_payment_date_wnsl)
;

-- ============================================================================
-- MERGED TABLE COUNTS BY BUSINESS UNIT
-- ============================================================================
INSERT INTO validation_customer_result
SELECT 'merged_by_bu', 'src_fin_customer_FA', COUNT(*) FROM src_fin_customer WHERE business_unit = 'FA'
;
INSERT INTO validation_customer_result
SELECT 'merged_by_bu', 'src_fin_customer_NFC', COUNT(*) FROM src_fin_customer WHERE business_unit = 'NFC'
;
INSERT INTO validation_customer_result
SELECT 'merged_by_bu', 'src_fin_customer_WNSL', COUNT(*) FROM src_fin_customer WHERE business_unit = 'WNSL'
;

-- Source table count (total)
INSERT INTO validation_customer_result
SELECT 'record_count', 'src_fin_customer', COUNT(*) FROM src_fin_customer
;

-- Record counts by workday table
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_name', COUNT(*) FROM workday_customer_name
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_groups', COUNT(*) FROM workday_customer_groups
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_related_worktags', COUNT(*) FROM workday_customer_related_worktags
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_credit', COUNT(*) FROM workday_customer_credit
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_payment', COUNT(*) FROM workday_customer_payment
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_currencies', COUNT(*) FROM workday_customer_currencies
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_company_restrictions', COUNT(*) FROM workday_customer_company_restrictions
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_document_delivery', COUNT(*) FROM workday_customer_document_delivery
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_settlement_account', COUNT(*) FROM workday_customer_settlement_account
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_intermediary_account', COUNT(*) FROM workday_customer_intermediary_account
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_bank_account_address', COUNT(*) FROM workday_customer_bank_account_address
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_credit_card', COUNT(*) FROM workday_customer_credit_card
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_tax', COUNT(*) FROM workday_customer_tax
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_address', COUNT(*) FROM workday_customer_address
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_phone', COUNT(*) FROM workday_customer_phone
;
INSERT INTO validation_customer_result
SELECT 'record_count', 'workday_customer_email', COUNT(*) FROM workday_customer_email
;

-- Data quality checks
INSERT INTO validation_customer_result
SELECT 'data_quality', 'orphan_groups', COUNT(*)
  FROM workday_customer_groups g
  LEFT JOIN workday_customer_name n ON g.customer_id = n.customer_id
 WHERE n.customer_id IS NULL
;
INSERT INTO validation_customer_result
SELECT 'data_quality', 'missing_required_customer_name', COUNT(*)
  FROM workday_customer_name
 WHERE customer_id IS NULL OR customer_name IS NULL
;
INSERT INTO validation_customer_result
SELECT 'data_quality', 'missing_required_currencies', COUNT(*)
  FROM workday_customer_currencies
 WHERE customer_id IS NULL OR accepted_currencies_plus IS NULL
;
INSERT INTO validation_customer_result
SELECT 'data_quality', 'invalid_email_addresses', COUNT(*)
  FROM src_fin_customer
 WHERE email_to_address IS NOT NULL
   AND email_to_address !~ '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'
;

-- ============================================================================
-- INGESTION RECONCILIATION SUMMARY
-- This compares row counts from raw ingested tables vs merged table by business unit
-- Note: Merged counts may differ due to DISTINCT and deduplication logic
-- ============================================================================
DROP TABLE IF EXISTS validation_customer_reconciliation
;
CREATE TABLE validation_customer_reconciliation AS
WITH ingestion_totals AS (
    SELECT 'FA' AS business_unit,
           (SELECT COUNT(*) FROM fin_customer_debtor_last_3_years_fa) +
           (SELECT COUNT(*) FROM fin_customer_debtor_created_date_fa) +
           (SELECT COUNT(*) FROM fin_customer_debtor_last_payment_date_fa) AS ingested_rows
    UNION ALL
    SELECT 'NFC' AS business_unit,
           (SELECT COUNT(*) FROM fin_customer_debtor_last_3_years_nfc) +
           (SELECT COUNT(*) FROM fin_customer_debtor_created_date_nfc) +
           (SELECT COUNT(*) FROM fin_customer_debtor_last_payment_date_nfc) AS ingested_rows
    UNION ALL
    SELECT 'WNSL' AS business_unit,
           (SELECT COUNT(*) FROM fin_customer_debtor_last_3_years_wnsl) +
           (SELECT COUNT(*) FROM fin_customer_debtor_created_date_wnsl) +
           (SELECT COUNT(*) FROM fin_customer_debtor_last_payment_date_wnsl) AS ingested_rows
),
merged_totals AS (
    SELECT business_unit, COUNT(*) AS merged_rows
      FROM src_fin_customer
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
