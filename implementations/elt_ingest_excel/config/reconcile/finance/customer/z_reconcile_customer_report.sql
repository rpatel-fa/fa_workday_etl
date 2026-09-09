-- ============================================================================
-- RECONCILE: Customer reconciliation report (consolidated)
-- Business sign-off report: confirms GP source data ingested reconciles to
-- the target data created for the Workday customer EIBs, with no customer
-- lost or duplicated by GP customer_number.
--
-- Every metric below is a numerator/denominator pair (source count vs target
-- count, or covered vs total), not a bare pass/fail flag - so retention can
-- be independently verified by eye rather than trusted on faith.
--
-- Single table combining:
--   1. GP customer_number reconciliation - source count vs target-traced
--      count vs duplication (backed by validate_customer_id_coverage.sql).
--   2. Row-count funnel (source side) - GP ingest through the final,
--      deduplicated customer record.
--   3. Attribute coverage (address/email/phone) - covered count vs total
--      source count, so a shortfall is visible as a ratio, not just a count.
--   4. Workday EIB output counts (target side) - row count per workday_customer_*
--      table, i.e. what was actually created for each EIB tab.
--
-- Depends on validate_customer_id_coverage.sql having been run beforehand
-- (i.e. examples/fin_customer_debtor_validate_sql.py), so that
-- validation_customer_issues exists in this database.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_customer_report
;
CREATE TABLE reconcile_customer_report
    AS
  WITH cte_source_ids
    AS (
SELECT DISTINCT TRIM(b.customer_number)                                                         customer_number
  FROM src_fin_customer_base                                                                    b
 WHERE NULLIF(TRIM(b.customer_number), '')                                                      IS NOT NULL
       )
     , cte_target_ids
    AS (
SELECT DISTINCT TRIM(u.customer_number)                                                         customer_number
  FROM workday_customer_name                                                                    n
     , UNNEST(STRING_SPLIT(n.reference_id, '|'))                                                u(customer_number)
       )
     , cte_source_addresses
    AS (
SELECT c.customer_id                                                                            customer_id
     , u.address                                                                                address
  FROM src_fin_customer                                                                         c
     , UNNEST(c.nrm_array_customer_address)                                                     u(address)
 WHERE COALESCE(
           NULLIF(TRIM(u.address.nrm_address_line_1), '')
         , NULLIF(TRIM(u.address.nrm_address_line_2), '')
         , NULLIF(TRIM(u.address.nrm_address_line_3), '')
         , NULLIF(TRIM(u.address.nrm_address_line_4), '')
       )                                                                                        IS NOT NULL
       )
     , cte_source_emails
    AS (
SELECT DISTINCT
       c.customer_id                                                                            customer_id
     , TRIM(LOWER(u.email))                                                                     email_address
  FROM src_fin_customer                                                                         c
     , UNNEST(STRING_SPLIT(
           COALESCE(c.nrm_agg_email_to_address, '')  || ';' ||
           COALESCE(c.nrm_agg_email_cc_address, '')  || ';' ||
           COALESCE(c.nrm_agg_email_bcc_address, '')
       , ';'))                                                                                  u(email)
 WHERE NULLIF(TRIM(u.email), '')                                                                IS NOT NULL
   AND TRIM(u.email)                                                                            ~ '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'
       )
     , cte_source_phones
    AS (
SELECT c.customer_id                                                                            customer_id
     , u.phone                                                                                  phone
  FROM src_fin_customer                                                                         c
     , UNNEST(STRING_SPLIT(
           COALESCE(c.nrm_agg_phone_1, '') || ';' ||
           COALESCE(c.nrm_agg_phone_2, '') || ';' ||
           COALESCE(c.nrm_agg_phone_3, '') || ';' ||
           COALESCE(c.nrm_agg_fax,     '')
       , ';'))                                                                                  u(phone)
 WHERE NULLIF(TRIM(u.phone), '')                                                                IS NOT NULL
       )
SELECT '1_gp_id_reconciliation'                                                        section
     , '1a_distinct_gp_customer_number_in_source (src_fin_customer_base)'              metric
     , (SELECT COUNT(*) FROM cte_source_ids)::VARCHAR                                  metric_value
     , CAST(NULL AS VARCHAR)                                                          status

 UNION ALL
SELECT '1_gp_id_reconciliation'
     , '1b_distinct_gp_customer_number_traced_to_target (workday_customer_name.reference_id)'
     , (SELECT COUNT(*) FROM cte_target_ids)::VARCHAR
     , CAST(NULL AS VARCHAR)

 UNION ALL
SELECT '1_gp_id_reconciliation'
     , '1c_coverage/duplicate_issues (src_fin_customer + workday_customer_name)'
     , COUNT(*)::VARCHAR
     , CASE
         WHEN COUNT(*) = 0
         THEN 'CONFIRMED - every GP customer_number in source is retained exactly once in target'
         ELSE 'REVIEW - see validation_customer_issues'
       END
  FROM validation_customer_issues
 WHERE (table_name, rule_type, col) IN (
       ('src_fin_customer',      'COVERAGE',  'customer_number')
     , ('workday_customer_name', 'COVERAGE',  'customer_id')
     , ('src_fin_customer',      'DUPLICATE', 'customer_number')
     , ('workday_customer_name', 'DUPLICATE', 'customer_number')
       )

 UNION ALL
SELECT '1_gp_id_reconciliation'
     , '1d_row_level_proof (see reconcile_customer_gp_id_proof for every GP customer_number and its resulting customer_id)'
     , (SELECT COUNT(*) FROM reconcile_customer_gp_id_proof WHERE found_in_workday_customer_name = 'Y')::VARCHAR ||
       ' found / ' ||
       (SELECT COUNT(*) FROM reconcile_customer_gp_id_proof WHERE found_in_workday_customer_name = 'N')::VARCHAR ||
       ' not found'
     , CAST(NULL AS VARCHAR)

 UNION ALL
SELECT '2_source_row_funnel'                                                           section
     , stage                                                                           metric
     , row_count::VARCHAR                                                             metric_value
     , CAST(NULL AS VARCHAR)                                                          status
  FROM reconcile_customer_row_funnel

 UNION ALL
SELECT '3_attribute_coverage'
     , '3a_address: source non-blank addresses vs present in workday_customer_address_base'
     , (SELECT COUNT(*) FROM cte_source_addresses)::VARCHAR || ' source / ' ||
       (SELECT COUNT(*) FROM workday_customer_address_base)::VARCHAR || ' target'
     , CAST(NULL AS VARCHAR)

 UNION ALL
SELECT '3_attribute_coverage'
     , '3b_email: valid distinct source emails vs present in workday_customer_email'
     , (SELECT COUNT(*) FROM cte_source_emails)::VARCHAR || ' source / ' ||
       (SELECT COUNT(*) FROM (SELECT DISTINCT customer_id, email_address FROM workday_customer_email))::VARCHAR || ' target'
     , (SELECT COUNT(*) FROM reconciliation_customer_issues WHERE table_name = 'workday_customer_email')::VARCHAR ||
       ' lost to cross-customer dedup (same email shared by another customer)'

 UNION ALL
SELECT '3_attribute_coverage'
     , '3c_phone: distinct source numbers aggregated vs rows in workday_customer_phone'
     , (SELECT COUNT(*) FROM cte_source_phones)::VARCHAR || ' source / ' ||
       (SELECT COUNT(*) FROM workday_customer_phone)::VARCHAR || ' target'
     , (SELECT COUNT(*) FROM reconciliation_customer_issues WHERE table_name = 'workday_customer_phone')::VARCHAR ||
       ' customers with a shortfall (see reconciliation_customer_issues for detail)'

 UNION ALL
SELECT '3_attribute_coverage'
     , '3d_address_row_level_proof (see reconcile_customer_address_proof - exact field match, address_line_1-4/city/postal_code)'
     , (SELECT COUNT(*) FROM reconcile_customer_address_proof WHERE found_in_workday_customer_address_base = 'Y')::VARCHAR ||
       ' found / ' ||
       (SELECT COUNT(*) FROM reconcile_customer_address_proof WHERE found_in_workday_customer_address_base = 'N')::VARCHAR ||
       ' not found'
     , 'Row-level check catches content mismatches (e.g. a dropped address_line_3/4) that the count-only check in 3a cannot see - see reconcile_customer_address_proof for detail'

 UNION ALL
SELECT '3_attribute_coverage'
     , '3e_email_row_level_proof (see reconcile_customer_email_proof - exact match, same customer_id)'
     , (SELECT COUNT(*) FROM reconcile_customer_email_proof WHERE found_in_workday_customer_email = 'Y')::VARCHAR ||
       ' found / ' ||
       (SELECT COUNT(*) FROM reconcile_customer_email_proof WHERE found_in_workday_customer_email = 'N')::VARCHAR ||
       ' not found'
     , CAST(NULL AS VARCHAR)

 UNION ALL
SELECT '3_attribute_coverage'
     , '3f_phone_row_level_proof_APPROX (see reconcile_customer_phone_proof - best-effort digit match, not exact)'
     , (SELECT COUNT(*) FROM reconcile_customer_phone_proof WHERE found_in_workday_customer_phone_approx = 'Y')::VARCHAR ||
       ' found / ' ||
       (SELECT COUNT(*) FROM reconcile_customer_phone_proof WHERE found_in_workday_customer_phone_approx = 'N')::VARCHAR ||
       ' not found'
     , 'APPROXIMATE match only - workday_customer_phone.sql reformats numbers, so this cannot be an exact proof like address/email'

 UNION ALL
SELECT '3_attribute_coverage'
     , '3g_payment_coverage (see reconcile_customer_payment_proof - 1:1 passthrough, coverage guaranteed by construction)'
     , (SELECT COUNT(*) FROM reconcile_customer_payment_proof WHERE found_in_workday_customer_payment = 'Y')::VARCHAR ||
       ' found / ' ||
       (SELECT COUNT(*) FROM reconcile_customer_payment_proof WHERE found_in_workday_customer_payment = 'N')::VARCHAR ||
       ' not found'
     , CAST(NULL AS VARCHAR)

 UNION ALL
SELECT '3_attribute_coverage'
     , '3h_tax_coverage (see reconcile_customer_tax_proof - 1:1 passthrough, coverage guaranteed by construction)'
     , (SELECT COUNT(*) FROM reconcile_customer_tax_proof WHERE found_in_workday_customer_tax = 'Y')::VARCHAR ||
       ' found / ' ||
       (SELECT COUNT(*) FROM reconcile_customer_tax_proof WHERE found_in_workday_customer_tax = 'N')::VARCHAR ||
       ' not found'
     , CAST(NULL AS VARCHAR)

 UNION ALL
-- Excludes tables permanently disabled via a hardcoded `1 = 2` filter in
-- their transform (never populated by design, not a coverage gap):
-- workday_customer_company_restrictions, workday_customer_bank_account_address,
-- workday_customer_credit, workday_customer_credit_card,
-- workday_customer_intermediary_account, workday_customer_settlement_account,
-- workday_customer_related_worktags.
SELECT '4_workday_eib_output_counts'                                                   section
     , table_name                                                                     metric
     , row_count::VARCHAR                                                             metric_value
     , CAST(NULL AS VARCHAR)                                                          status
  FROM (
        SELECT 'workday_customer_name'                    table_name, COUNT(*) row_count FROM workday_customer_name
  UNION ALL
        SELECT 'workday_customer_groups'                  table_name, COUNT(*) row_count FROM workday_customer_groups
  UNION ALL
        SELECT 'workday_customer_payment'                 table_name, COUNT(*) row_count FROM workday_customer_payment
  UNION ALL
        SELECT 'workday_customer_currencies'              table_name, COUNT(*) row_count FROM workday_customer_currencies
  UNION ALL
        SELECT 'workday_customer_document_delivery'       table_name, COUNT(*) row_count FROM workday_customer_document_delivery
  UNION ALL
        SELECT 'workday_customer_tax'                     table_name, COUNT(*) row_count FROM workday_customer_tax
  UNION ALL
        SELECT 'workday_customer_address'                 table_name, COUNT(*) row_count FROM workday_customer_address
  UNION ALL
        SELECT 'workday_customer_phone'                   table_name, COUNT(*) row_count FROM workday_customer_phone
  UNION ALL
        SELECT 'workday_customer_email'                   table_name, COUNT(*) row_count FROM workday_customer_email
       )                                                                              t

 UNION ALL
SELECT '5_name_traceability'
     , 'customers with a full source-name audit trail (see reconcile_customer_name_traceability - not a pass/fail check, names are deliberately conformed)'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconcile_customer_name_traceability

 ORDER BY 1, 2
;
