DROP TABLE IF EXISTS src_fin_customer_include
;
CREATE TABLE src_fin_customer_include
    AS
  WITH cte_src
    AS (
SELECT t.nrm_agg_customer_number                                  src_agg_reference_id
     , TRIM(UNNEST(STRING_SPLIT(t.nrm_agg_customer_number, '|'))) src_reference_id
  FROM src_fin_customer_rnk            t
 WHERE t.rnk_score                     = 1   
       )
     , cte_wd
    AS (
SELECT t.reference_id                                               wd_agg_reference_id
     , TRIM(UNNEST(STRING_SPLIT(t.reference_id, '|')))              wd_reference_id
  FROM fin_workday_extract_customer    t
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
SELECT s.nrm_agg_customer_number
  FROM src_fin_customer_rnk            s
       LEFT OUTER JOIN
       cte_ccs_match                   m
          ON m.src_agg_reference_id    = s.nrm_agg_customer_number
 WHERE 
       m.src_agg_reference_id          IS NULL
       )
SELECT DISTINCT
       t.*
  FROM cte_missing                     t
;