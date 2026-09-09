-- ----------------------------------------------------------------------------
-- COVERAGE: vendor_id from base must exist in src_fin_supplier
--
-- vendor_id is scoped per business_unit in GP - the same code can be reused for
-- unrelated vendors in different company books. Coverage therefore matches the
-- exact bare vendor_id OR the exact vendor_id-business_unit token for that SAME
-- base row's business_unit, not any arbitrary vendor_id-% suffix.
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'src_fin_supplier'                                                         table_name
     , 'COVERAGE'                                                                 rule_type
     , 'vendor_id'                                                                col
     , CAST(NULL AS VARCHAR)                                                      supplier_id
     , s.vendor_id || ' (' || s.business_unit || ')'                              detail
     , 'vendor_id from src_fin_supplier_base is missing from src_fin_supplier.'   message
  FROM (
        SELECT DISTINCT
               TRIM(b.vendor_id)                                                  vendor_id
             , UPPER(TRIM(b.business_unit))                                       business_unit
          FROM src_fin_supplier_base                                              b
         WHERE NULLIF(TRIM(b.vendor_id), '') IS NOT NULL
       )                                                                          s
       LEFT OUTER JOIN (
        SELECT TRIM(UNNEST(STRING_SPLIT(nrm_agg_supplier_number, '|')))           vendor_id
          FROM src_fin_supplier
       )                                                                          t
          ON (
                 t.vendor_id = s.vendor_id
              OR t.vendor_id = s.vendor_id || '-' || s.business_unit
             )
 WHERE t.vendor_id IS NULL
;

-- ----------------------------------------------------------------------------
-- COVERAGE: vendor_id from base must appear in workday_supplier_name
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_name'                                                             table_name
     , 'COVERAGE'                                                                         rule_type
     , 'supplier_id'                                                                      col
     , t.supplier_id                                                                      supplier_id
     , s.vendor_id || ' (' || s.business_unit || ')'                                      detail
     , 'vendor_id from src_fin_supplier_base is missing from workday_supplier_name.'      message
  FROM (
        SELECT DISTINCT
               TRIM(b.vendor_id)                                                          vendor_id
             , UPPER(TRIM(b.business_unit))                                               business_unit
          FROM src_fin_supplier_base                                                      b
         WHERE NULLIF(TRIM(b.vendor_id), '') IS NOT NULL
       )                                                                                  s
       LEFT OUTER JOIN (
        SELECT supplier_id
             , TRIM(UNNEST(STRING_SPLIT(reference_id, '|')))                             vendor_id
          FROM workday_supplier_name
       )                                                                                  t
          ON (
                 t.vendor_id = s.vendor_id
              OR t.vendor_id = s.vendor_id || '-' || s.business_unit
             )
 WHERE t.supplier_id IS NULL
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: each canonical nrm_supplier_number must map to exactly one row in
-- src_fin_supplier.
--
-- Checked on nrm_supplier_number (the value in nrm_agg_supplier_number), not the
-- raw GP vendor_id: a business-unit-collision split intentionally makes one
-- vendor_id resolve to several nrm_supplier_number values (e.g. 'HAL003-FA',
-- 'HAL003-NFC') by design, so checking raw vendor_id 1:1 here would flag every
-- expected split as a false-positive duplicate. Each nrm_supplier_number itself
-- must still map to exactly one supplier_id - that invariant is unaffected by
-- the split and is what this check verifies.
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'src_fin_supplier'                                                                        table_name
     , 'DUPLICATE'                                                                              rule_type
     , 'nrm_supplier_number'                                                                    col
     , CAST(NULL AS VARCHAR)                                                                    supplier_id
     , t.nrm_supplier_number                                                                    detail
     , 'nrm_supplier_number matches multiple rows in src_fin_supplier.'                         message
  FROM (
        SELECT supplier_id
             , TRIM(UNNEST(STRING_SPLIT(nrm_agg_supplier_number, '|')))                        nrm_supplier_number
          FROM src_fin_supplier
       )                                                                                        t
 GROUP BY t.nrm_supplier_number
HAVING COUNT(DISTINCT t.supplier_id) > 1
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: each canonical nrm_supplier_number must map to exactly one row in
-- workday_supplier_name. Same business-unit-collision reasoning as above.
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_name'                                                                  table_name
     , 'DUPLICATE'                                                                              rule_type
     , 'nrm_supplier_number'                                                                    col
     , CAST(NULL AS VARCHAR)                                                                    supplier_id
     , t.nrm_supplier_number                                                                    detail
     , 'nrm_supplier_number matches multiple rows in workday_supplier_name.'                    message
  FROM (
        SELECT supplier_id
             , TRIM(UNNEST(STRING_SPLIT(reference_id, '|')))                                   nrm_supplier_number
          FROM workday_supplier_name
       )                                                                                        t
 GROUP BY t.nrm_supplier_number
HAVING COUNT(DISTINCT t.supplier_id) > 1
;
