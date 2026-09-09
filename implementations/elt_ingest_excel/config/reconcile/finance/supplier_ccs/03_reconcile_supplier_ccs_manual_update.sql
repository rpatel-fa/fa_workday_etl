-- Known go-live manual-update exceptions. Gayle stripped address / bank / payment
-- details from these suppliers to force the Workday load; they must be updated
-- manually in Production. The list is name-keyed only, so it is matched to the
-- Workday extract on normalised supplier name to recover the reference_id.

DROP TABLE IF EXISTS reconcile_supplier_ccs_manual_update
;
CREATE TABLE reconcile_supplier_ccs_manual_update
    AS
  WITH cte_manual
    AS (
SELECT m.seq
     , m.bucket
     , m.supplier_name_raw
     , nrm_name(m.supplier_name_raw)                                                 nrm_supplier_name
     , CASE
         WHEN m.bucket ILIKE 'REMOVED ADDRESS%'      THEN 'ADDRESS'
         WHEN m.bucket ILIKE 'REMOVED BANK%'         THEN 'BANK'
         WHEN m.bucket ILIKE 'REMOVED ALL PAYMENT%'  THEN 'PAYMENT'
         ELSE 'OTHER'
       END                                                                           exception_type
  FROM fin_supplier_manual_update                                                    m
       )
     , cte_extract
    AS (
SELECT nrm_name(w.supplier)                                                          nrm_supplier_name
     , STRING_AGG(DISTINCT w.reference_id, '|' ORDER BY w.reference_id)              workday_reference_id
     , STRING_AGG(DISTINCT w.supplier_id,  '|' ORDER BY w.supplier_id)               workday_supplier_id
     , STRING_AGG(DISTINCT w.supplier,     '|' ORDER BY w.supplier)                  workday_supplier_name
  FROM fin_workday_extract_supplier                                                  w
 WHERE NULLIF(TRIM(w.reference_id), '')                                              IS NOT NULL
 GROUP BY
       nrm_name(w.supplier)
       )
SELECT c.seq
     , c.bucket
     , c.exception_type
     , c.supplier_name_raw
     , e.workday_supplier_name
     , e.workday_reference_id
     , e.workday_supplier_id
     , CASE WHEN e.nrm_supplier_name IS NOT NULL THEN 'Y' ELSE 'N' END               matched_to_extract
  FROM cte_manual                                                                    c
       LEFT OUTER JOIN
       cte_extract                                                                   e
          ON e.nrm_supplier_name                                                     = c.nrm_supplier_name
 ORDER BY
       c.exception_type
     , c.supplier_name_raw
;

DROP TABLE IF EXISTS reconcile_supplier_ccs_manual_update_unmatched
;
CREATE TABLE reconcile_supplier_ccs_manual_update_unmatched
    AS
SELECT bucket
     , exception_type
     , supplier_name_raw
  FROM reconcile_supplier_ccs_manual_update
 WHERE matched_to_extract                                                            = 'N'
 ORDER BY
       exception_type
     , supplier_name_raw
;

INSERT INTO reconciliation_supplier_ccs_issues
SELECT '3_known_exceptions'
     , 'MANUAL_UPDATE_' || t.exception_type
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , t.workday_supplier_id
     , CAST(NULL AS VARCHAR)
     , t.workday_reference_id
     , t.exception_type
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , 'Known go-live manual-update supplier (' || t.bucket
       || CASE WHEN t.matched_to_extract = 'N' THEN '; NOT matched to extract' ELSE '' END || ')'
  FROM reconcile_supplier_ccs_manual_update                                          t
;
