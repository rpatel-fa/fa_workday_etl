-- ============================================================================
-- RECONCILE: Customer name traceability (audit trail, not pass/fail)
-- Names are NOT reconciled by equality: the final nrm_customer_name is a
-- deliberately conformed/cleaned name (case, "AND"->"&", suffix stripping,
-- fuzzy merge across duplicate GP records), so it is expected to differ from
-- the raw source name. There is nothing to "prove" by matching.
--
-- What CAN be proven is traceability: every final customer_id here is shown
-- alongside every distinct raw source customer_name (from src_fin_customer_base)
-- across every GP customer_number that was merged into it, so a reviewer can
-- see exactly which source records the conformed name was derived from.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_customer_name_traceability
;
CREATE TABLE reconcile_customer_name_traceability
    AS
  WITH cte_customer_number_list
    AS (
SELECT c.customer_id                                                                            customer_id
     , c.nrm_customer_name                                                                      final_workday_name
     , TRIM(u.customer_number)                                                                  gp_customer_number
  FROM src_fin_customer                                                                         c
     , UNNEST(STRING_SPLIT(c.nrm_agg_customer_number, '|'))                                     u(customer_number)
       )
SELECT n.customer_id                                                                            customer_id
     , n.final_workday_name                                                                     final_workday_name
     , COUNT(DISTINCT n.gp_customer_number)                                                      gp_ids_merged
     , STRING_AGG(DISTINCT b.customer_name, ' | ' ORDER BY b.customer_name)                     raw_source_names
  FROM cte_customer_number_list                                                                 n
  LEFT JOIN src_fin_customer_base                                                                b
         ON TRIM(b.customer_number)                                                              = n.gp_customer_number
 GROUP BY n.customer_id
     , n.final_workday_name
 ORDER BY n.customer_id
;
