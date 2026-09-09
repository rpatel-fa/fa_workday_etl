DROP TABLE IF EXISTS src_fin_supplier_purchase_order_base;

CREATE TABLE src_fin_supplier_purchase_order_base
    AS
  WITH cte_purchase_order
    AS (
SELECT t.* FROM fin_supplier_purchase_order     t
       )
SELECT *
     , accountingelement1 || '-' || accountingelement2 || '-' || accountingelement3 || '-' || accountingelement4 "22_digit_account_string"
  FROM cte_purchase_order
;  
