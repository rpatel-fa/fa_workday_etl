-- Attribute reconciliation at GP Vendor ID grain (the key that matches 100% on
-- both sides). Compares what the transform sent (workday_supplier_*) against what
-- the Workday extract now holds, then classifies address / bank / payment gaps as
-- EXPECTED (explained by the go-live manual-update list) or UNEXPLAINED.

-- Expected attributes per GP Vendor ID (inherited from its merged supplier).
DROP TABLE IF EXISTS reconcile_supplier_ccs_expected_supplier
;
CREATE TABLE reconcile_supplier_ccs_expected_supplier
    AS
SELECT mb.supplier_number
     , mb.expected_supplier_id
     , mb.expected_reference_id
     , n.supplier_name                                                               expected_supplier_name
     , n.supplier_category                                                           expected_supplier_category
     , p.payment_terms                                                               expected_payment_terms
     , p.default_payment_type                                                        expected_default_payment_type
     , p.payment_types_accepted                                                       expected_payment_types_accepted
     , c.default_currency                                                            expected_default_currency
     , t.tax_id_country                                                              expected_tax_id_country
     , t.tax_id_type                                                                 expected_tax_id_type
     , t.tax_id                                                                      expected_tax_id
     , CASE WHEN a.supplier_id IS NOT NULL THEN TRUE ELSE FALSE END                  expected_address_present
     , CASE WHEN s.supplier_id IS NOT NULL THEN TRUE ELSE FALSE END                  expected_bank_present
     , CASE
         WHEN LEN(LIST_FILTER(STRING_SPLIT(REGEXP_REPLACE(UPPER(COALESCE(p.payment_types_accepted, '')), '[^A-Z]+', ' ', 'g'), ' '),
                              x -> x <> '' AND x <> 'MANUAL')) > 0
         THEN TRUE ELSE FALSE
       END                                                                           expected_payment_multi
  FROM reconcile_supplier_ccs_source_merge_bridge                                    mb
       INNER JOIN
       workday_supplier_name                                                         n
          ON n.supplier_id                                                           = mb.expected_supplier_id
       LEFT OUTER JOIN
       workday_supplier_payment                                                      p
          ON p.supplier_id                                                           = mb.expected_supplier_id
       LEFT OUTER JOIN
       workday_supplier_currencies                                                   c
          ON c.supplier_id                                                           = mb.expected_supplier_id
       LEFT OUTER JOIN
       workday_supplier_tax                                                          t
          ON t.supplier_id                                                           = mb.expected_supplier_id
       LEFT OUTER JOIN
       (
         SELECT DISTINCT supplier_id
           FROM workday_supplier_address_base
          WHERE is_primary                                                           = 'Yes'
       )                                                                             a
          ON a.supplier_id                                                           = mb.expected_supplier_id
       LEFT OUTER JOIN
       (
         SELECT DISTINCT supplier_id
           FROM workday_supplier_settlement_account
          WHERE COALESCE(NULLIF(TRIM(bank_account_number), ''), NULLIF(TRIM(bank_iban), '')) IS NOT NULL
       )                                                                             s
          ON s.supplier_id                                                           = mb.expected_supplier_id
;

-- Actual attributes per GP Vendor ID, aggregated over the extract row(s) that
-- carry it (a supplier can occupy multiple extract rows for split address/bank).
DROP TABLE IF EXISTS reconcile_supplier_ccs_actual_supplier
;
CREATE TABLE reconcile_supplier_ccs_actual_supplier
    AS
  WITH cte_wx
    AS (
SELECT TRIM(u.supplier_number)                                                       supplier_number
     , w.supplier
     , w.reference_id
     , w.supplier_category
     , w.default_payment_terms
     , w.default_payment_type
     , w.payment_types_accepted
     , w.default_currency
     , w.country_for_tax_id
     , w.company_tax_id_type
     , w.tax_id
     , COALESCE(NULLIF(TRIM(w.primary_address), '')
              , NULLIF(TRIM(w.primary_remit_to_address), '')
              , NULLIF(TRIM(w.remit_to_address_for_payee), '')
              , NULLIF(TRIM(w.payee_address), '')
              , NULLIF(TRIM(w.order_from_address), '')
              , NULLIF(TRIM(w.all_addresses), ''))                                    address_any
     , w.account_number
     , w.iban
  FROM fin_workday_extract_supplier                                                  w
     , UNNEST(STRING_SPLIT(w.reference_id, '|'))                                     u(supplier_number)
 WHERE NULLIF(TRIM(u.supplier_number), '')                                           IS NOT NULL
       )
SELECT supplier_number
     , STRING_AGG(DISTINCT supplier,       '|' ORDER BY supplier)                    actual_supplier_name
     , STRING_AGG(DISTINCT reference_id,   '|' ORDER BY reference_id)                actual_reference_id
     , STRING_AGG(DISTINCT supplier_category, '|' ORDER BY supplier_category)        actual_supplier_category
     , STRING_AGG(DISTINCT default_payment_terms, '|' ORDER BY default_payment_terms)
                                                                                    actual_payment_terms
     , STRING_AGG(DISTINCT default_payment_type,  '|' ORDER BY default_payment_type)
                                                                                    actual_default_payment_type
     , STRING_AGG(DISTINCT default_currency, '|' ORDER BY default_currency)          actual_default_currency
     , STRING_AGG(DISTINCT country_for_tax_id, '|' ORDER BY country_for_tax_id)      actual_tax_id_country
     , STRING_AGG(DISTINCT company_tax_id_type, '|' ORDER BY company_tax_id_type)    actual_tax_id_type
     , STRING_AGG(DISTINCT tax_id,          '|' ORDER BY tax_id)                     actual_tax_id
     , BOOL_OR(address_any IS NOT NULL)                                              actual_address_present
     , BOOL_OR(COALESCE(NULLIF(TRIM(account_number), ''), NULLIF(TRIM(iban), '')) IS NOT NULL)
                                                                                    actual_bank_present
     , CASE
         WHEN BOOL_AND(LEN(LIST_FILTER(STRING_SPLIT(REGEXP_REPLACE(UPPER(COALESCE(payment_types_accepted, '')), '[^A-Z]+', ' ', 'g'), ' '),
                                       x -> x <> '' AND x <> 'MANUAL')) = 0)
         THEN TRUE ELSE FALSE
       END                                                                           actual_payment_manual_only
  FROM cte_wx
 GROUP BY
       supplier_number
;

-- Manual-update exception buckets per GP Vendor ID.
DROP TABLE IF EXISTS reconcile_supplier_ccs_attribute_detail
;
CREATE TABLE reconcile_supplier_ccs_attribute_detail
    AS
  WITH cte_manual_by_number
    AS (
SELECT TRIM(u.supplier_number)                                                       supplier_number
     , STRING_AGG(DISTINCT mu.exception_type, ',' ORDER BY mu.exception_type)        manual_exception_types
  FROM reconcile_supplier_ccs_manual_update                                          mu
     , UNNEST(STRING_SPLIT(mu.workday_reference_id, '|'))                            u(supplier_number)
 WHERE mu.matched_to_extract                                                         = 'Y'
   AND NULLIF(TRIM(u.supplier_number), '')                                           IS NOT NULL
 GROUP BY
       TRIM(u.supplier_number)
       )
SELECT e.supplier_number
     , e.expected_supplier_id
     , e.expected_reference_id
     , a.actual_reference_id
     , e.expected_supplier_name
     , a.actual_supplier_name
     , e.expected_supplier_category
     , a.actual_supplier_category
     , e.expected_default_payment_type
     , a.actual_default_payment_type
     , e.expected_default_currency
     , a.actual_default_currency
     , e.expected_tax_id_country
     , a.actual_tax_id_country
     , e.expected_address_present
     , a.actual_address_present
     , e.expected_bank_present
     , a.actual_bank_present
     , e.expected_payment_multi
     , a.actual_payment_manual_only
     , COALESCE(mb.manual_exception_types, '')                                       manual_exception_types
  FROM reconcile_supplier_ccs_expected_supplier                                      e
       LEFT OUTER JOIN
       reconcile_supplier_ccs_actual_supplier                                        a
          ON a.supplier_number                                                       = e.supplier_number
       LEFT OUTER JOIN
       cte_manual_by_number                                                          mb
          ON mb.supplier_number                                                      = e.supplier_number
;

-- One classified row per (GP Vendor ID, gap type) for the three go-live gaps.
DROP TABLE IF EXISTS reconcile_supplier_ccs_attribute_gap
;
CREATE TABLE reconcile_supplier_ccs_attribute_gap
    AS
SELECT supplier_number
     , expected_reference_id
     , expected_supplier_name
     , 'ADDRESS'                                                                     gap_type
     , manual_exception_types
     , CASE WHEN manual_exception_types LIKE '%ADDRESS%' THEN 'EXPECTED' ELSE 'UNEXPLAINED' END
                                                                                    gap_status
  FROM reconcile_supplier_ccs_attribute_detail
 WHERE expected_address_present                                                      = TRUE
   AND COALESCE(actual_address_present, FALSE)                                       = FALSE
UNION ALL
SELECT supplier_number
     , expected_reference_id
     , expected_supplier_name
     , 'BANK'
     , manual_exception_types
     , CASE WHEN manual_exception_types LIKE '%BANK%' THEN 'EXPECTED' ELSE 'UNEXPLAINED' END
  FROM reconcile_supplier_ccs_attribute_detail
 WHERE expected_bank_present                                                         = TRUE
   AND COALESCE(actual_bank_present, FALSE)                                          = FALSE
UNION ALL
SELECT supplier_number
     , expected_reference_id
     , expected_supplier_name
     , 'PAYMENT'
     , manual_exception_types
     , CASE WHEN manual_exception_types LIKE '%PAYMENT%' OR manual_exception_types LIKE '%BANK%'
            THEN 'EXPECTED' ELSE 'UNEXPLAINED' END
  FROM reconcile_supplier_ccs_attribute_detail
 WHERE expected_payment_multi                                                        = TRUE
   AND COALESCE(actual_payment_manual_only, FALSE)                                   = TRUE
;

DROP TABLE IF EXISTS reconcile_supplier_ccs_attribute_summary
;
CREATE TABLE reconcile_supplier_ccs_attribute_summary
    AS
SELECT gap_type
     , gap_status
     , COUNT(*)                                                                       supplier_count
  FROM reconcile_supplier_ccs_attribute_gap
 GROUP BY
       gap_type
     , gap_status
 ORDER BY
       gap_type
     , gap_status
;

INSERT INTO reconciliation_supplier_ccs_issues
SELECT '4_attribute_gaps'
     , g.gap_type || '_GAP_' || g.gap_status
     , g.supplier_number
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , g.expected_reference_id
     , CAST(NULL AS VARCHAR)
     , g.gap_type
     , 'present in transform'
     , 'absent / reduced in Workday'
     , CASE WHEN g.gap_status = 'EXPECTED'
            THEN 'Explained by go-live manual-update list (' || g.manual_exception_types || ')'
            ELSE 'Not on the manual-update list — review' END
  FROM reconcile_supplier_ccs_attribute_gap                                          g
;

-- Merge-boundary differences: GP vendor ids whose transform merge group differs
-- from the Workday merge group (surfaced, not treated as a data defect).
DROP TABLE IF EXISTS reconcile_supplier_ccs_merge_group_diff
;
CREATE TABLE reconcile_supplier_ccs_merge_group_diff
    AS
SELECT DISTINCT
       d.supplier_number
     , d.expected_supplier_name
     , d.expected_reference_id                                                        transform_merge_group
     , d.actual_reference_id                                                          workday_merge_group
  FROM reconcile_supplier_ccs_attribute_detail                                       d
 WHERE d.actual_reference_id                                                         IS NOT NULL
   AND d.expected_reference_id                                                       <> d.actual_reference_id
 ORDER BY
       d.supplier_number
;

INSERT INTO reconciliation_supplier_ccs_issues
SELECT '4_attribute_gaps'
     , 'MERGE_GROUP_DIFF'
     , t.supplier_number
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , t.transform_merge_group
     , t.workday_merge_group
     , 'merge_group'
     , t.transform_merge_group
     , t.workday_merge_group
     , 'GP vendor id merged differently in transform vs Workday'
  FROM reconcile_supplier_ccs_merge_group_diff                                        t
;
