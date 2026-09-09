DROP TABLE IF EXISTS src_fin_customer
;
CREATE TABLE src_fin_customer
    AS
  WITH cte_customer
    AS (
SELECT t.*
  FROM src_fin_customer_rnk                              t
 WHERE rnk_score = 1
       )
     , cte_customer_rnk
    AS (
SELECT t.*
     , ROW_NUMBER() OVER(ORDER BY t.nrm_customer_name, t.nrm_agg_customer_number)  rnk
  FROM cte_customer                                      t
       )
     , cte_customer_id
    AS (
SELECT
       'C-' || LPAD(CAST((t.rnk + 4400) AS VARCHAR), 6, '0')                customer_id
     , t.*
  FROM cte_customer_rnk                                  t
       )
SELECT t.* 
  FROM cte_customer_id                 t
 --WHERE t.nrm_agg_customer_number IN (SELECT i.nrm_agg_customer_number FROM src_fin_customer_include i)
;
