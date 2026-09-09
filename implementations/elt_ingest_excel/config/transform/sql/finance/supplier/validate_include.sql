    WITH cte_src
    AS (
SELECT t.nrm_agg_supplier_number                                  src_agg_reference_id
     , TRIM(UNNEST(STRING_SPLIT(t.nrm_agg_supplier_number, '|'))) src_reference_id
  FROM src_fin_supplier_include            t
       )
     , cte_wd
    AS (
SELECT t.reference_id                                               wd_agg_reference_id
     , TRIM(UNNEST(STRING_SPLIT(t.reference_id, '|')))              wd_reference_id
  FROM fin_supplier_creditor_existing    t
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
SELECT * FROM cte_ccs_match