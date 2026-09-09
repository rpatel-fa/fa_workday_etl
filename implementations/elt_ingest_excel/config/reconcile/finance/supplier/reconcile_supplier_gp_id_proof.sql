-- ============================================================================
-- RECONCILE: Supplier GP ID proof (row-level)
-- One row per distinct GP vendor_id found anywhere in the source GP extracts
-- (src_fin_supplier_base), showing whether it is present in the final
-- workday_supplier_name table and, if so, the supplier_id it landed under.
-- workday_supplier_name.reference_id carries the pipe-delimited list of every
-- GP vendor_id merged into that record (see workday_supplier_name.sql:
-- reference_id := nrm_agg_supplier_number).
--
-- vendor_id is scoped per business_unit in GP - the same code can be reused for
-- unrelated vendors in different company books. src_fin_supplier_raw.sql detects
-- this and suffixes '-{business_unit}' onto nrm_supplier_number for those cases,
-- so the match below tolerates an optional business_unit suffix on the target
-- side, not just an exact vendor_id.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_supplier_gp_id_proof
;
CREATE TABLE reconcile_supplier_gp_id_proof
    AS
  WITH cte_source_ids
    AS (
SELECT DISTINCT TRIM(b.vendor_id)                                                               gp_vendor_id
  FROM src_fin_supplier_base                                                                    b
 WHERE NULLIF(TRIM(b.vendor_id), '')                                                             IS NOT NULL
       )
     , cte_target_ids
    AS (
SELECT n.supplier_id                                                                             supplier_id
     , TRIM(u.vendor_id)                                                                         gp_vendor_id
  FROM workday_supplier_name                                                                     n
     , UNNEST(STRING_SPLIT(n.reference_id, '|'))                                                 u(vendor_id)
       )
SELECT s.gp_vendor_id                                                                             gp_vendor_id
     , t.supplier_id                                                                              supplier_id
     , CASE WHEN t.supplier_id IS NOT NULL THEN 'Y' ELSE 'N' END                                  found_in_workday_supplier_name
  FROM cte_source_ids                                                                             s
  LEFT JOIN cte_target_ids                                                                        t
         ON (t.gp_vendor_id = s.gp_vendor_id OR t.gp_vendor_id LIKE s.gp_vendor_id || '-%')
 ORDER BY found_in_workday_supplier_name
     , gp_vendor_id
;
