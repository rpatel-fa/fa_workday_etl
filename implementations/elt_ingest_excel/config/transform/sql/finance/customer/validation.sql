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
-- PFF club mapped to a non-Active Workday customer (expected 0; fix via a
-- pff_workday_manual_mapped remap row with current_workday_id = the inactive id)
INSERT INTO validation_customer_result
SELECT 'data_quality', 'pff_os_mapped_wd_customer_not_active', COUNT(*)
  FROM workday_pff_os_customer_load_base
 WHERE load_status = 'MAPPED'
   AND COALESCE(wd_customer_status, '') != 'Active'
;
