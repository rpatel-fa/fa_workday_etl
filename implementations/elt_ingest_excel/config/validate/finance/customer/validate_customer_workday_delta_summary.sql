DROP TABLE IF EXISTS validate_customer_workday_delta_summary
;
CREATE TABLE validate_customer_workday_delta_summary
    AS
  WITH cte_entity
    AS (
SELECT 'FA'   entity, 1 sort_order
 UNION ALL
SELECT 'NFC'  entity, 2 sort_order
 UNION ALL
SELECT 'WNSL' entity, 3 sort_order
       )
     , cte_workday_current
    AS (
-- `fin_workday_extract_customer` does not retain a clean entity breakdown.
-- Use the preserved source-to-Workday bridge, which already explodes pipe-delimited
-- reference IDs back to source customer_number and carries source_business_units.
SELECT t.source_business_units                          entity
     , COUNT(*)                                         workday_current
  FROM reconcile_customer_ccs_bridge                    t
 WHERE t.in_source_gp                                   = 'Y'
   AND t.source_business_units                          IN ('FA', 'NFC', 'WNSL')
 GROUP BY
       t.source_business_units
       )
     , cte_workday_reference_token
    AS (
SELECT DISTINCT
       UPPER(TRIM(u.reference_token))                   reference_token
  FROM fin_workday_extract_customer                     t
     , UNNEST(STRING_SPLIT(t.reference_id, '|'))        u(reference_token)
 WHERE NULLIF(TRIM(t.reference_id), '')                 IS NOT NULL
   AND NULLIF(TRIM(u.reference_token), '')              IS NOT NULL
       )
     , cte_delta_source
    AS (
SELECT DISTINCT
       UPPER(TRIM(t.company))                           entity
     , TRIM(t.customer_number)                          customer_number
  FROM fin_customer_debtor_delta_1                      t
 WHERE NULLIF(TRIM(t.customer_number), '')              IS NOT NULL
   AND UPPER(TRIM(t.company))                           IN ('FA', 'NFC', 'WNSL')
       )
     , cte_delta
    AS (
SELECT s.entity
     , COUNT(*)                                         gp_next_delta
     , SUM(
           CASE
               WHEN EXISTS (
                    SELECT 1
                      FROM cte_workday_reference_token w
                     WHERE w.reference_token IN (
                               UPPER(s.customer_number)
                             , UPPER(s.customer_number || '-' || s.entity)
                           )
               )
               THEN 0
               ELSE 1
           END
       )                                                additional_delta_to_load
  FROM cte_delta_source                                 s
 GROUP BY
       s.entity
       )
     , cte_detail
    AS (
SELECT e.sort_order
     , e.entity
     , COALESCE(w.workday_current, 0)                                      gp_source_initial
     , COALESCE(d.gp_next_delta, 0)                                        gp_next_delta
     , COALESCE(w.workday_current, 0) + COALESCE(d.gp_next_delta, 0)
                                                                          gp_source_total
     , COALESCE(w.workday_current, 0)                                      workday_current
     , COALESCE(d.additional_delta_to_load, 0)                             additional_delta_to_load
     , COALESCE(w.workday_current, 0) + COALESCE(d.additional_delta_to_load, 0)
                                                                          workday_plus_delta
  FROM cte_entity                                         e
       LEFT OUTER JOIN
       cte_workday_current                                w
          ON w.entity                                     = e.entity
       LEFT OUTER JOIN
       cte_delta                                          d
          ON d.entity                                     = e.entity
       )
     , cte_output
    AS (
SELECT entity
     , gp_source_initial
     , gp_next_delta
     , gp_source_total
     , workday_current
     , additional_delta_to_load
     , workday_plus_delta
     , CASE
           WHEN gp_source_total = workday_plus_delta THEN 'Match'
           ELSE 'Mismatch'
       END                                               status
  FROM cte_detail

 UNION ALL

SELECT 'Total'                                                                  entity
     , SUM(gp_source_initial)                                                   gp_source_initial
     , SUM(gp_next_delta)                                                       gp_next_delta
     , SUM(gp_source_total)                                                     gp_source_total
     , SUM(workday_current)                                                     workday_current
     , SUM(additional_delta_to_load)                                            additional_delta_to_load
     , SUM(workday_plus_delta)                                                  workday_plus_delta
     , CASE
           WHEN SUM(gp_source_total) = SUM(workday_plus_delta) THEN 'Match'
           ELSE 'Mismatch'
       END                                                                      status
  FROM cte_detail
       )
SELECT *
  FROM cte_output
 ORDER BY
       CASE entity
           WHEN 'FA'    THEN 1
           WHEN 'NFC'   THEN 2
           WHEN 'WNSL'  THEN 3
           WHEN 'Total' THEN 4
           ELSE 99
       END
;
