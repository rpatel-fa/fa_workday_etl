DROP TABLE IF EXISTS src_fin_supplier
;
CREATE TABLE src_fin_supplier
    AS
  WITH cte_supplier
    AS (
SELECT t.*
  FROM src_fin_supplier_rnk                              t
 WHERE rnk_score = 1
       )
     , cte_supplier_rnk
    AS (
SELECT t.*
     , ROW_NUMBER() OVER(ORDER BY t.nrm_supplier_name, t.nrm_agg_supplier_number)   rnk
  FROM cte_supplier                                      t
 WHERE t.nrm_payment_terms_id IS NOT NULL
       )
     , cte_supplier_id
    AS (
SELECT
       'S-' || LPAD(CAST((t.rnk + 9000) AS VARCHAR), 6, '0')                supplier_id
     , t.*
  FROM cte_supplier_rnk                                  t
       )
     , cte_supplier_include
    AS (
SELECT t.nrm_agg_supplier_number
     , t.nrm_resolved_agg_supplier_number
  FROM src_fin_supplier_include                           t
 WHERE t.include_ind                                      = 1
       )
SELECT t.*
     , i.nrm_resolved_agg_supplier_number
  FROM cte_supplier_id                                      t
       INNER JOIN
       cte_supplier_include                                 i
          ON i.nrm_agg_supplier_number                      = t.nrm_agg_supplier_number
;
