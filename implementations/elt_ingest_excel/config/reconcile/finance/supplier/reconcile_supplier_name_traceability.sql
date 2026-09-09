-- ============================================================================
-- RECONCILE: Supplier name traceability (audit trail, not pass/fail)
-- Names are NOT reconciled by equality: the final nrm_supplier_name is a
-- deliberately conformed/cleaned name (case, "AND"->"&", suffix stripping,
-- fuzzy merge across duplicate GP records), so it is expected to differ from
-- the raw source name. There is nothing to "prove" by matching.
--
-- What CAN be proven is traceability: every final supplier_id here is shown
-- alongside every distinct raw source vendor_name (from src_fin_supplier_base)
-- across every GP vendor_id that was merged into it, so a reviewer can see
-- exactly which source records the conformed name was derived from.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_supplier_name_traceability
;
CREATE TABLE reconcile_supplier_name_traceability
    AS
  WITH cte_vendor_id_list
    AS (
SELECT s.supplier_id                                                                             supplier_id
     , s.nrm_supplier_name                                                                       final_workday_name
     , TRIM(u.vendor_id)                                                                         gp_vendor_id
  FROM src_fin_supplier                                                                          s
     , UNNEST(STRING_SPLIT(s.nrm_agg_supplier_number, '|'))                                      u(vendor_id)
       )
SELECT n.supplier_id                                                                             supplier_id
     , n.final_workday_name                                                                      final_workday_name
     , COUNT(DISTINCT n.gp_vendor_id)                                                            gp_ids_merged
     , STRING_AGG(DISTINCT b.vendor_name, ' | ' ORDER BY b.vendor_name)                          raw_source_names
  FROM cte_vendor_id_list                                                                        n
  LEFT JOIN src_fin_supplier_base                                                                b
         ON TRIM(b.vendor_id)                                                                    = n.gp_vendor_id
 GROUP BY n.supplier_id
     , n.final_workday_name
 ORDER BY n.supplier_id
;
