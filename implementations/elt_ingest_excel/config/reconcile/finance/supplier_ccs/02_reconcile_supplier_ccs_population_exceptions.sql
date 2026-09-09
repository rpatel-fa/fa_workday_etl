-- Population exceptions: GP suppliers absent from Workday, Workday rows with no GP
-- lineage (dummy/test rows), and suppliers landing as multiple extract rows.

DROP TABLE IF EXISTS reconcile_supplier_ccs_missing_in_workday
;
CREATE TABLE reconcile_supplier_ccs_missing_in_workday
    AS
SELECT s.supplier_number
     , s.source_business_units
     , s.source_supplier_name
     , s.source_country
     , s.expected_supplier_id
     , s.expected_reference_id
  FROM reconcile_supplier_ccs_source_profile                                         s
       LEFT OUTER JOIN
       (
         SELECT DISTINCT supplier_number
           FROM reconcile_supplier_ccs_workday_reference_bridge
       )                                                                             w
          ON w.supplier_number                                                       = s.supplier_number
 WHERE w.supplier_number                                                             IS NULL
;

DROP TABLE IF EXISTS reconcile_supplier_ccs_extra_in_workday
;
CREATE TABLE reconcile_supplier_ccs_extra_in_workday
    AS
SELECT w.supplier_number
     , w.workday_supplier_id
     , w.workday_supplier_name
     , w.workday_reference_id
  FROM reconcile_supplier_ccs_workday_reference_bridge                               w
       LEFT OUTER JOIN
       reconcile_supplier_ccs_source_profile                                         s
          ON s.supplier_number                                                       = w.supplier_number
 WHERE s.supplier_number                                                             IS NULL
;

-- Workday extract rows whose reference_id maps to no GP source vendor number
-- (blank reference_id or tenant test rows).
DROP TABLE IF EXISTS reconcile_supplier_ccs_workday_dummy_row
;
CREATE TABLE reconcile_supplier_ccs_workday_dummy_row
    AS
SELECT w.supplier                                                                    workday_supplier_name
     , w.reference_id                                                                workday_reference_id
     , w.supplier_id                                                                 workday_supplier_id
     , w.supplier_status                                                             workday_supplier_status
  FROM fin_workday_extract_supplier                                                  w
 WHERE NOT EXISTS
       (
         SELECT 1
           FROM UNNEST(STRING_SPLIT(w.reference_id, '|'))                            u(supplier_number)
                INNER JOIN
                reconcile_supplier_ccs_source_profile                                s
                   ON s.supplier_number                                              = TRIM(u.supplier_number)
       )
;

-- Suppliers materialising as more than one Workday extract row (multi address /
-- multi bank blocks). Expected for split banking; surfaced for eyeball review.
DROP TABLE IF EXISTS reconcile_supplier_ccs_multi_row_supplier
;
CREATE TABLE reconcile_supplier_ccs_multi_row_supplier
    AS
SELECT w.reference_id                                                                workday_reference_id
     , STRING_AGG(DISTINCT w.supplier, ' || ' ORDER BY w.supplier)                   workday_supplier_name
     , COUNT(*)                                                                       extract_row_count
     , COUNT(DISTINCT w.supplier_id)                                                  workday_supplier_id_count
  FROM fin_workday_extract_supplier                                                  w
 WHERE NULLIF(TRIM(w.reference_id), '')                                              IS NOT NULL
 GROUP BY
       w.reference_id
HAVING COUNT(*)                                                                      > 1
;

-- One GP vendor number expanding to more than one distinct Workday supplier row.
DROP TABLE IF EXISTS reconcile_supplier_ccs_source_multi_workday
;
CREATE TABLE reconcile_supplier_ccs_source_multi_workday
    AS
SELECT w.supplier_number
     , COUNT(DISTINCT w.workday_supplier_id)                                          workday_supplier_count
     , STRING_AGG(DISTINCT w.workday_supplier_id,   '|' ORDER BY w.workday_supplier_id)
                                                                                    workday_supplier_ids
     , STRING_AGG(DISTINCT w.workday_reference_id, '|' ORDER BY w.workday_reference_id)
                                                                                    workday_reference_ids
  FROM reconcile_supplier_ccs_workday_reference_bridge                               w
 GROUP BY
       w.supplier_number
HAVING COUNT(DISTINCT w.workday_supplier_id)                                         > 1
;

INSERT INTO reconciliation_supplier_ccs_issues
SELECT '2_population_exceptions'
     , 'MISSING_IN_WORKDAY'
     , t.supplier_number
     , t.expected_supplier_id
     , CAST(NULL AS VARCHAR)
     , t.expected_reference_id
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , 'GP supplier vendor_id is not present in any exploded Workday reference_id'
  FROM reconcile_supplier_ccs_missing_in_workday                                     t
;

INSERT INTO reconciliation_supplier_ccs_issues
SELECT '2_population_exceptions'
     , 'EXTRA_IN_WORKDAY'
     , t.supplier_number
     , CAST(NULL AS VARCHAR)
     , t.workday_supplier_id
     , CAST(NULL AS VARCHAR)
     , t.workday_reference_id
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , 'Exploded Workday reference_id is not present in GP source'
  FROM reconcile_supplier_ccs_extra_in_workday                                       t
;

INSERT INTO reconciliation_supplier_ccs_issues
SELECT '2_population_exceptions'
     , 'WORKDAY_DUMMY_ROW'
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , t.workday_supplier_id
     , CAST(NULL AS VARCHAR)
     , t.workday_reference_id
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , 'Workday extract row has no matching GP source vendor number'
  FROM reconcile_supplier_ccs_workday_dummy_row                                      t
;

INSERT INTO reconciliation_supplier_ccs_issues
SELECT '2_population_exceptions'
     , 'SOURCE_NUMBER_LINKED_TO_MULTIPLE_WORKDAY_ROWS'
     , t.supplier_number
     , CAST(NULL AS VARCHAR)
     , t.workday_supplier_ids
     , CAST(NULL AS VARCHAR)
     , t.workday_reference_ids
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , 'One GP vendor number expands to more than one Workday supplier row'
  FROM reconcile_supplier_ccs_source_multi_workday                                   t
;
