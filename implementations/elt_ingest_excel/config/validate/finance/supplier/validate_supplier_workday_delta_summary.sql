DROP TABLE IF EXISTS validate_supplier_workday_delta_summary
;
CREATE TABLE validate_supplier_workday_delta_summary
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
-- Supplier current-state entity counts come from the preserved CCS bridge, which
-- already carries the source business unit alongside each exploded supplier_number.
SELECT CASE t.source_business_units
           WHEN 'The Football Association'          THEN 'FA'
           WHEN 'National Football Centre Limited'  THEN 'NFC'
           WHEN 'Wembley National Stadium Ltd'      THEN 'WNSL'
       END                                             entity
     , COUNT(*)                                        workday_current
  FROM reconcile_supplier_ccs_bridge                   t
 WHERE t.in_source_gp                                  = 'Y'
   AND t.source_business_units                         IN (
           'The Football Association'
         , 'National Football Centre Limited'
         , 'Wembley National Stadium Ltd'
       )
 GROUP BY
       CASE t.source_business_units
           WHEN 'The Football Association'          THEN 'FA'
           WHEN 'National Football Centre Limited'  THEN 'NFC'
           WHEN 'Wembley National Stadium Ltd'      THEN 'WNSL'
       END
       )
     , cte_workday_reference_token
    AS (
SELECT DISTINCT
       UPPER(TRIM(u.reference_token))                  reference_token
  FROM fin_supplier_creditor_existing                  t
     , UNNEST(STRING_SPLIT(t.reference_id, '|'))       u(reference_token)
 WHERE NULLIF(TRIM(t.reference_id), '')                IS NOT NULL
   AND NULLIF(TRIM(u.reference_token), '')             IS NOT NULL
       )
     , cte_delta_source
    AS (
-- Use the transformed supplier_number, not the raw vendor_id, because supplier
-- collisions can suffix '-{business_unit}' onto the authoritative key.
SELECT DISTINCT
       CASE UPPER(TRIM(s.nrm_business_unit))
           WHEN 'THE FOOTBALL ASSOCIATION'         THEN 'FA'
           WHEN 'NATIONAL FOOTBALL CENTRE LIMITED' THEN 'NFC'
           WHEN 'WEMBLEY NATIONAL STADIUM LTD'     THEN 'WNSL'
       END                                             entity
     , TRIM(u.supplier_number)                         supplier_number
  FROM src_fin_supplier                                s
     , UNNEST(STRING_SPLIT(s.nrm_agg_supplier_number, '|')) u(supplier_number)
 WHERE s.rnk_score                                     = 1
   AND NULLIF(TRIM(u.supplier_number), '')             IS NOT NULL
       )
     , cte_delta
    AS (
SELECT s.entity
     , COUNT(*)                                        gp_next_delta
     , SUM(
           CASE
               WHEN EXISTS (
                    SELECT 1
                      FROM cte_workday_reference_token w
                     WHERE w.reference_token           = UPPER(s.supplier_number)
               )
               THEN 1
               ELSE 0
           END
       )                                               already_in_workday
     , SUM(
           CASE
               WHEN EXISTS (
                    SELECT 1
                      FROM cte_workday_reference_token w
                     WHERE w.reference_token           = UPPER(s.supplier_number)
               )
               THEN 0
               ELSE 1
           END
       )                                               additional_delta_to_load
  FROM cte_delta_source                                s
 GROUP BY
       s.entity
       )
     , cte_detail
    AS (
SELECT e.sort_order
     , e.entity
     , COALESCE(w.workday_current, 0)                                      gp_source_initial
     , COALESCE(d.gp_next_delta, 0)                                        gp_next_delta
     , COALESCE(d.already_in_workday, 0)                                   already_in_workday
     , COALESCE(w.workday_current, 0)
       + COALESCE(d.gp_next_delta, 0)
       - COALESCE(d.already_in_workday, 0)                                 gp_source_total
     , COALESCE(w.workday_current, 0)                                      workday_current
     , COALESCE(d.additional_delta_to_load, 0)                             additional_delta_to_load
     , COALESCE(w.workday_current, 0)
       + COALESCE(d.additional_delta_to_load, 0)                           workday_plus_delta
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
     , already_in_workday
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
     , SUM(already_in_workday)                                                  already_in_workday
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
