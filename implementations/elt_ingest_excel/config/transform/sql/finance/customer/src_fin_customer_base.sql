DROP TABLE IF EXISTS src_fin_customer_base;

CREATE TABLE src_fin_customer_base
    AS
  WITH cte_customer_src
    AS (
SELECT t.company business_unit, t.* FROM fin_customer_debtor_delta_1     t
       )
SELECT *
  FROM cte_customer_src
;  