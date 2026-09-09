DROP TABLE IF EXISTS src_fin_supplier_match
;
CREATE TABLE src_fin_supplier_match
    AS
  WITH cte_src
    AS (
SELECT COALESCE(
           t.nrm_resolved_agg_supplier_number
         , t.nrm_agg_supplier_number
       )                                                          src_agg_reference_id
     , TRIM(UNNEST(STRING_SPLIT(
           COALESCE(
               t.nrm_resolved_agg_supplier_number
             , t.nrm_agg_supplier_number
           )
         , '|')))                                                 src_reference_id
  FROM src_fin_supplier                                            t
 WHERE t.rnk_score                                                 = 1
       )
     , cte_wd
    AS (
SELECT t.reference_id                                               wd_agg_reference_id
     , TRIM(UNNEST(STRING_SPLIT(t.reference_id, '|')))              wd_reference_id
  FROM fin_workday_extract_supplier    t
       )
     , cte_ccs_match
    AS (
SELECT DISTINCT
       s.src_agg_reference_id
  FROM cte_src                         s
       INNER JOIN
       cte_wd                          w
          ON w.wd_reference_id         = s.src_reference_id
       )
     , cte_missing
    AS (
SELECT s.*
  FROM src_fin_supplier               s
       LEFT OUTER JOIN
       cte_ccs_match                   m
          ON m.src_agg_reference_id    = COALESCE(
                                             s.nrm_resolved_agg_supplier_number
                                           , s.nrm_agg_supplier_number
                                         )
 WHERE 
       m.src_agg_reference_id          IS NOT NULL
       )
SELECT DISTINCT
       t.*
  FROM cte_missing                     t
;
