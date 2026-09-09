-- ============================================================================
-- RECONCILE: Customer GP ID proof (row-level)
-- One row per distinct GP customer_number found anywhere in the source GP
-- extracts (src_fin_customer_base), showing whether it is present in the
-- final workday_customer_name table and, if so, the customer_id it landed
-- under. workday_customer_name.reference_id carries the pipe-delimited list
-- of every GP customer_number merged into that record
-- (see workday_customer_name.sql: reference_id := nrm_agg_customer_number).
-- This is the row-level backing for the aggregate counts in
-- reconcile_customer_report section 1.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_customer_gp_id_proof
;
CREATE TABLE reconcile_customer_gp_id_proof
    AS
  WITH cte_source_ids
    AS (
SELECT DISTINCT TRIM(b.customer_number)                                                         gp_customer_number
  FROM src_fin_customer_base                                                                    b
 WHERE NULLIF(TRIM(b.customer_number), '')                                                      IS NOT NULL
       )
     , cte_target_ids
    AS (
SELECT n.customer_id                                                                             customer_id
     , TRIM(u.customer_number)                                                                   gp_customer_number
  FROM workday_customer_name                                                                     n
     , UNNEST(STRING_SPLIT(n.reference_id, '|'))                                                 u(customer_number)
       )
SELECT s.gp_customer_number                                                                      gp_customer_number
     , t.customer_id                                                                             customer_id
     , CASE WHEN t.customer_id IS NOT NULL THEN 'Y' ELSE 'N' END                                found_in_workday_customer_name
  FROM cte_source_ids                                                                            s
  LEFT JOIN cte_target_ids                                                                       t
         ON t.gp_customer_number                                                                = s.gp_customer_number
 ORDER BY found_in_workday_customer_name
     , gp_customer_number
;
