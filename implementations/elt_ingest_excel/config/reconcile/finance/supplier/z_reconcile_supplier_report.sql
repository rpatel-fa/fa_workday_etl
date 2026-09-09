-- ============================================================================
-- RECONCILE: Supplier reconciliation report (consolidated)
-- Business sign-off report: confirms GP source data ingested reconciles to
-- the target data created for the Workday supplier EIBs, with no supplier
-- lost or duplicated by GP vendor_id.
--
-- Every metric below is a numerator/denominator pair (source count vs target
-- count, or covered vs total), not a bare pass/fail flag - so retention can
-- be independently verified by eye rather than trusted on faith.
--
-- Single table combining:
--   1. GP vendor_id reconciliation - source count vs target-traced count vs
--      duplication (backed by validate_supplier_id_coverage.sql).
--   2. Row-count funnel (source side) - GP ingest through the final,
--      deduplicated supplier record.
--   3. Attribute coverage (address/bank/email/phone/payment/tax) - covered
--      count vs total source count, so a shortfall is visible as a ratio.
--   4. Workday EIB output counts (target side) - row count per
--      workday_supplier_* table, i.e. what was actually created for each
--      EIB tab.
--   5. Name traceability - audit trail, not a pass/fail check.
--
-- Depends on validate_supplier_id_coverage.sql having been run beforehand
-- (i.e. examples/fin_supplier_creditor_validate_sql.py), so that
-- validation_supplier_issues exists in this database.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_supplier_report
;
CREATE TABLE reconcile_supplier_report
    AS
  WITH cte_source_ids
    AS (
SELECT DISTINCT TRIM(b.vendor_id)                                                                vendor_id
  FROM src_fin_supplier_base                                                                     b
 WHERE NULLIF(TRIM(b.vendor_id), '')                                                              IS NOT NULL
       )
     , cte_target_ids
    AS (
SELECT DISTINCT TRIM(u.vendor_id)                                                                vendor_id
  FROM workday_supplier_name                                                                     n
     , UNNEST(STRING_SPLIT(n.reference_id, '|'))                                                 u(vendor_id)
       )
     , cte_source_addresses
    AS (
SELECT s.supplier_id                                                                              supplier_id
     , u.address                                                                                 address
  FROM src_fin_supplier                                                                          s
     , UNNEST(s.nrm_array_address)                                                               u(address)
 WHERE COALESCE(
           NULLIF(TRIM(u.address.nrm_address_line_1), '')
         , NULLIF(TRIM(u.address.nrm_address_line_2), '')
         , NULLIF(TRIM(u.address.nrm_address_line_3), '')
         , NULLIF(TRIM(u.address.nrm_address_line_4), '')
       )                                                                                         IS NOT NULL
       )
     , cte_source_bank
    AS (
SELECT s.supplier_id                                                                              supplier_id
     , COALESCE(u.bank.nrm_bank_account_number, u.bank.nrm_bank_iban)                             bank_identifier
  FROM src_fin_supplier                                                                          s
     , UNNEST(s.nrm_array_bank)                                                                  u(bank)
 WHERE COALESCE(u.bank.nrm_bank_account_number, u.bank.nrm_bank_iban)                             IS NOT NULL
   AND NULLIF(TRIM(u.bank.nrm_bank_name), '')                                                     IS NOT NULL
       )
     , cte_source_emails
    AS (
SELECT DISTINCT
       s.supplier_id                                                                              supplier_id
     , TRIM(LOWER(u.email))                                                                       email_address
  FROM src_fin_supplier                                                                          s
     , UNNEST(STRING_SPLIT(
           COALESCE(s.nrm_agg_email_to_address, '')  || ';' ||
           COALESCE(s.nrm_agg_email_cc_address, '')  || ';' ||
           COALESCE(s.nrm_agg_email_bcc_address, '')
       , ';'))                                                                                   u(email)
 WHERE NULLIF(TRIM(u.email), '')                                                                  IS NOT NULL
   AND TRIM(u.email)                                                                              ~ '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'
       )
     , cte_source_phones
    AS (
SELECT s.supplier_id                                                                              supplier_id
     , u.phone                                                                                    phone
  FROM src_fin_supplier                                                                          s
     , UNNEST(STRING_SPLIT(
           COALESCE(s.nrm_agg_phone_1, '') || ';' ||
           COALESCE(s.nrm_agg_phone_2, '') || ';' ||
           COALESCE(s.nrm_agg_phone_3, '') || ';' ||
           COALESCE(s.nrm_agg_fax,     '')
       , ';'))                                                                                   u(phone)
 WHERE NULLIF(TRIM(u.phone), '')                                                                  IS NOT NULL
       )
SELECT '1_gp_id_reconciliation'                                                        section
     , '1a_distinct_gp_vendor_id_in_source (src_fin_supplier_base)'                    metric
     , (SELECT COUNT(*) FROM cte_source_ids)::VARCHAR                                  metric_value
     , CAST(NULL AS VARCHAR)                                                          status

 UNION ALL
SELECT '1_gp_id_reconciliation'
     , '1b_distinct_gp_vendor_id_traced_to_target (workday_supplier_name.reference_id)'
     , (SELECT COUNT(*) FROM cte_target_ids)::VARCHAR
     , CAST(NULL AS VARCHAR)

 UNION ALL
SELECT '1_gp_id_reconciliation'
     , '1c_coverage/duplicate_issues (src_fin_supplier + workday_supplier_name)'
     , COUNT(*)::VARCHAR
     , CASE
         WHEN COUNT(*) = 0
         THEN 'CONFIRMED - every GP vendor_id in source is retained exactly once in target'
         ELSE 'REVIEW - see validation_supplier_issues'
       END
  FROM validation_supplier_issues
 WHERE (table_name, rule_type, col) IN (
       ('src_fin_supplier',      'COVERAGE',  'vendor_id')
     , ('workday_supplier_name', 'COVERAGE',  'supplier_id')
     , ('src_fin_supplier',      'DUPLICATE', 'vendor_id')
     , ('workday_supplier_name', 'DUPLICATE', 'vendor_id')
       )

 UNION ALL
SELECT '1_gp_id_reconciliation'
     , '1d_row_level_proof (see reconcile_supplier_gp_id_proof for every GP vendor_id and its resulting supplier_id)'
     , (SELECT COUNT(*) FROM reconcile_supplier_gp_id_proof WHERE found_in_workday_supplier_name = 'Y')::VARCHAR ||
       ' found / ' ||
       (SELECT COUNT(*) FROM reconcile_supplier_gp_id_proof WHERE found_in_workday_supplier_name = 'N')::VARCHAR ||
       ' not found'
     , CAST(NULL AS VARCHAR)

 UNION ALL
SELECT '2_source_row_funnel'                                                           section
     , stage                                                                           metric
     , row_count::VARCHAR                                                             metric_value
     , CAST(NULL AS VARCHAR)                                                          status
  FROM reconcile_supplier_row_funnel

 UNION ALL
SELECT '3_attribute_coverage'
     , '3a_address: source non-blank addresses vs present in workday_supplier_address_base'
     , (SELECT COUNT(*) FROM cte_source_addresses)::VARCHAR || ' source / ' ||
       (SELECT COUNT(*) FROM workday_supplier_address_base)::VARCHAR || ' target'
     , CAST(NULL AS VARCHAR)

 UNION ALL
SELECT '3_attribute_coverage'
     , '3b_bank: identifiable source bank accounts vs present in workday_supplier_settlement_account'
     , (SELECT COUNT(*) FROM cte_source_bank)::VARCHAR || ' source / ' ||
       (SELECT COUNT(*) FROM workday_supplier_settlement_account)::VARCHAR || ' target'
     , CAST(NULL AS VARCHAR)

 UNION ALL
SELECT '3_attribute_coverage'
     , '3c_email: valid distinct source emails vs present in workday_supplier_email'
     , (SELECT COUNT(*) FROM cte_source_emails)::VARCHAR || ' source / ' ||
       (SELECT COUNT(*) FROM (SELECT DISTINCT supplier_id, email_address FROM workday_supplier_email))::VARCHAR || ' target'
     , (SELECT COUNT(*) FROM reconciliation_supplier_issues WHERE table_name = 'workday_supplier_email')::VARCHAR ||
       ' lost to cross-supplier dedup (same email shared by another supplier)'

 UNION ALL
SELECT '3_attribute_coverage'
     , '3d_phone: distinct source numbers aggregated vs rows in workday_supplier_phone'
     , (SELECT COUNT(*) FROM cte_source_phones)::VARCHAR || ' source / ' ||
       (SELECT COUNT(*) FROM workday_supplier_phone)::VARCHAR || ' target'
     , (SELECT COUNT(*) FROM reconciliation_supplier_issues WHERE table_name = 'workday_supplier_phone')::VARCHAR ||
       ' suppliers with a shortfall (see reconciliation_supplier_issues for detail)'

 UNION ALL
SELECT '3_attribute_coverage'
     , '3e_address_row_level_proof (see reconcile_supplier_address_proof - exact field match, address_line_1-4/city/postal_code)'
     , (SELECT COUNT(*) FROM reconcile_supplier_address_proof WHERE found_in_workday_supplier_address_base = 'Y')::VARCHAR ||
       ' found / ' ||
       (SELECT COUNT(*) FROM reconcile_supplier_address_proof WHERE found_in_workday_supplier_address_base = 'N')::VARCHAR ||
       ' not found'
     , 'Row-level check catches content mismatches that the count-only check in 3a cannot see - see reconcile_supplier_address_proof for detail'

 UNION ALL
SELECT '3_attribute_coverage'
     , '3f_bank_row_level_proof (see reconcile_supplier_bank_proof - exact match on account number/IBAN, same supplier_id)'
     , (SELECT COUNT(*) FROM reconcile_supplier_bank_proof WHERE found_in_workday_supplier_settlement_account = 'Y')::VARCHAR ||
       ' found / ' ||
       (SELECT COUNT(*) FROM reconcile_supplier_bank_proof WHERE found_in_workday_supplier_settlement_account = 'N')::VARCHAR ||
       ' not found'
     , CAST(NULL AS VARCHAR)

 UNION ALL
SELECT '3_attribute_coverage'
     , '3g_email_row_level_proof (see reconcile_supplier_email_proof - exact match, same supplier_id)'
     , (SELECT COUNT(*) FROM reconcile_supplier_email_proof WHERE found_in_workday_supplier_email = 'Y')::VARCHAR ||
       ' found / ' ||
       (SELECT COUNT(*) FROM reconcile_supplier_email_proof WHERE found_in_workday_supplier_email = 'N')::VARCHAR ||
       ' not found'
     , CAST(NULL AS VARCHAR)

 UNION ALL
SELECT '3_attribute_coverage'
     , '3h_phone_row_level_proof_APPROX (see reconcile_supplier_phone_proof - best-effort digit match, not exact)'
     , (SELECT COUNT(*) FROM reconcile_supplier_phone_proof WHERE found_in_workday_supplier_phone_approx = 'Y')::VARCHAR ||
       ' found / ' ||
       (SELECT COUNT(*) FROM reconcile_supplier_phone_proof WHERE found_in_workday_supplier_phone_approx = 'N')::VARCHAR ||
       ' not found'
     , 'APPROXIMATE match only - workday_supplier_phone.sql reformats numbers, so this cannot be an exact proof like address/email'

 UNION ALL
SELECT '3_attribute_coverage'
     , '3i_payment_coverage (see reconcile_supplier_payment_proof - 1:1 passthrough, coverage guaranteed by construction)'
     , (SELECT COUNT(*) FROM reconcile_supplier_payment_proof WHERE found_in_workday_supplier_payment = 'Y')::VARCHAR ||
       ' found / ' ||
       (SELECT COUNT(*) FROM reconcile_supplier_payment_proof WHERE found_in_workday_supplier_payment = 'N')::VARCHAR ||
       ' not found'
     , CAST(NULL AS VARCHAR)

 UNION ALL
SELECT '3_attribute_coverage'
     , '3j_tax_coverage (see reconcile_supplier_tax_proof - 1:1 passthrough, coverage guaranteed by construction)'
     , (SELECT COUNT(*) FROM reconcile_supplier_tax_proof WHERE found_in_workday_supplier_tax = 'Y')::VARCHAR ||
       ' found / ' ||
       (SELECT COUNT(*) FROM reconcile_supplier_tax_proof WHERE found_in_workday_supplier_tax = 'N')::VARCHAR ||
       ' not found'
     , CAST(NULL AS VARCHAR)

 UNION ALL
-- Excludes tables permanently disabled via a hardcoded `1 = 2` filter in
-- their transform (never populated by design, not a coverage gap):
-- workday_supplier_company_restrictions, workday_supplier_alternate_name,
-- workday_supplier_intermediary_account.
SELECT '4_workday_eib_output_counts'                                                   section
     , table_name                                                                     metric
     , row_count::VARCHAR                                                             metric_value
     , CAST(NULL AS VARCHAR)                                                          status
  FROM (
        SELECT 'workday_supplier_name'                    table_name, COUNT(*) row_count FROM workday_supplier_name
  UNION ALL
        SELECT 'workday_supplier_groups'                  table_name, COUNT(*) row_count FROM workday_supplier_groups
  UNION ALL
        SELECT 'workday_supplier_currencies'              table_name, COUNT(*) row_count FROM workday_supplier_currencies
  UNION ALL
        SELECT 'workday_supplier_tax'                     table_name, COUNT(*) row_count FROM workday_supplier_tax
  UNION ALL
        SELECT 'workday_supplier_payment'                 table_name, COUNT(*) row_count FROM workday_supplier_payment
  UNION ALL
        SELECT 'workday_supplier_procurement'             table_name, COUNT(*) row_count FROM workday_supplier_procurement
  UNION ALL
        SELECT 'workday_supplier_settlement_account'      table_name, COUNT(*) row_count FROM workday_supplier_settlement_account
  UNION ALL
        SELECT 'workday_supplier_address'                 table_name, COUNT(*) row_count FROM workday_supplier_address
  UNION ALL
        SELECT 'workday_supplier_phone'                   table_name, COUNT(*) row_count FROM workday_supplier_phone
  UNION ALL
        SELECT 'workday_supplier_email'                   table_name, COUNT(*) row_count FROM workday_supplier_email
       )                                                                              t

 UNION ALL
SELECT '5_name_traceability'
     , 'suppliers with a full source-name audit trail (see reconcile_supplier_name_traceability - not a pass/fail check, names are deliberately conformed)'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconcile_supplier_name_traceability

 ORDER BY 1, 2
;
