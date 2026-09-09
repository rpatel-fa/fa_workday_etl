-- ============================================================================
-- RECONCILE: Supplier attribute coverage
-- Confirms that address, bank, and email details aggregated across merged
-- duplicate records actually made it through to the workday_supplier_* tables.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- COVERAGE_COUNT: distinct non-blank addresses aggregated per supplier must
-- all appear in workday_supplier_address_base (pre country-specific filter,
-- so this isolates aggregation loss from the deliberate IT/BR/DO exclusion).
-- ----------------------------------------------------------------------------
INSERT INTO reconciliation_supplier_issues
  WITH cte_source_address_count
    AS (
SELECT s.supplier_id                                                                            supplier_id
     , COUNT(DISTINCT u.address)                                                                source_address_count
  FROM src_fin_supplier                                                                         s
     , UNNEST(s.nrm_array_address)                                                              u(address)
 WHERE COALESCE(
           NULLIF(TRIM(u.address.nrm_address_line_1), '')
         , NULLIF(TRIM(u.address.nrm_address_line_2), '')
         , NULLIF(TRIM(u.address.nrm_address_line_3), '')
         , NULLIF(TRIM(u.address.nrm_address_line_4), '')
       )                                                                                        IS NOT NULL
 GROUP BY s.supplier_id
       )
     , cte_final_address_count
    AS (
SELECT supplier_id
     , COUNT(*)                                                                                 final_address_count
  FROM workday_supplier_address_base
 GROUP BY supplier_id
       )
SELECT
       'workday_supplier_address_base'                                                          table_name
     , 'COVERAGE_COUNT'                                                                         rule_type
     , 'address'                                                                                col
     , s.supplier_id                                                                            supplier_id
     , s.source_address_count || ' source vs ' || COALESCE(f.final_address_count, 0) || ' final' detail
     , 'Distinct addresses aggregated from merged duplicate records do not all appear in workday_supplier_address_base.' message
  FROM cte_source_address_count                                                                 s
  LEFT JOIN cte_final_address_count                                                             f
         ON f.supplier_id                                                                       = s.supplier_id
 WHERE s.source_address_count > COALESCE(f.final_address_count, 0)
;

-- ----------------------------------------------------------------------------
-- COVERAGE_COUNT: distinct identifiable bank accounts aggregated per supplier
-- must all appear in workday_supplier_settlement_account.
-- ----------------------------------------------------------------------------
INSERT INTO reconciliation_supplier_issues
  WITH cte_source_bank_count
    AS (
SELECT s.supplier_id                                                                            supplier_id
     , COUNT(DISTINCT COALESCE(u.bank.nrm_bank_account_number, u.bank.nrm_bank_iban))           source_bank_count
  FROM src_fin_supplier                                                                         s
     , UNNEST(s.nrm_array_bank)                                                                 u(bank)
 WHERE COALESCE(u.bank.nrm_bank_account_number, u.bank.nrm_bank_iban)                           IS NOT NULL
   AND NULLIF(TRIM(u.bank.nrm_bank_name), '')                                                    IS NOT NULL
 GROUP BY s.supplier_id
       )
     , cte_final_bank_count
    AS (
SELECT supplier_id
     , COUNT(*)                                                                                 final_bank_count
  FROM workday_supplier_settlement_account
 GROUP BY supplier_id
       )
SELECT
       'workday_supplier_settlement_account'                                                    table_name
     , 'COVERAGE_COUNT'                                                                         rule_type
     , 'bank_account'                                                                           col
     , s.supplier_id                                                                            supplier_id
     , s.source_bank_count || ' source vs ' || COALESCE(f.final_bank_count, 0) || ' final'      detail
     , 'Distinct bank accounts aggregated from merged duplicate records do not all appear in workday_supplier_settlement_account.' message
  FROM cte_source_bank_count                                                                    s
  LEFT JOIN cte_final_bank_count                                                                f
         ON f.supplier_id                                                                       = s.supplier_id
 WHERE s.source_bank_count > COALESCE(f.final_bank_count, 0)
;

-- ----------------------------------------------------------------------------
-- CROSS_ID_DEDUP: workday_supplier_email keeps each email address under only
-- one supplier_id (RANK() PARTITION BY email_address in workday_supplier_email.sql).
-- If two suppliers share an email, the losing supplier's copy disappears from
-- the final table entirely. This flags every supplier where that happened.
-- ----------------------------------------------------------------------------
INSERT INTO reconciliation_supplier_issues
  WITH cte_source_emails
    AS (
SELECT DISTINCT
       s.supplier_id                                                                            supplier_id
     , TRIM(LOWER(u.email))                                                                     email_address
  FROM src_fin_supplier                                                                         s
     , UNNEST(STRING_SPLIT(
           COALESCE(s.nrm_agg_email_to_address, '')  || ';' ||
           COALESCE(s.nrm_agg_email_cc_address, '')  || ';' ||
           COALESCE(s.nrm_agg_email_bcc_address, '')
       , ';'))                                                                                  u(email)
 WHERE NULLIF(TRIM(u.email), '')                                                                IS NOT NULL
   AND TRIM(u.email)                                                                            ~ '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'
       )
SELECT
       'workday_supplier_email'                                                                 table_name
     , 'CROSS_ID_DEDUP'                                                                         rule_type
     , 'email_address'                                                                          col
     , se.supplier_id                                                                           supplier_id
     , se.email_address                                                                         detail
     , 'Valid source email is not present under this supplier_id in workday_supplier_email - likely lost to the cross-supplier dedup by email_address, or dropped outright.' message
  FROM cte_source_emails                                                                        se
  LEFT JOIN workday_supplier_email                                                               w
         ON w.supplier_id                                                                       = se.supplier_id
        AND w.email_address                                                                     = se.email_address
 WHERE w.email_address IS NULL
;
