DROP TABLE IF EXISTS src_fin_supplier_base
;
CREATE TABLE src_fin_supplier_base
    AS
  WITH cte_supplier_src
    AS (
SELECT t.company business_unit, t.* FROM fin_supplier_creditor_new      t
       )
SELECT * 
  FROM cte_supplier_src
;