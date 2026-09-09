DROP TABLE IF EXISTS src_fin_supplier_invoice;

CREATE TABLE src_fin_supplier_invoice
    AS
  WITH cte_base
    AS (    
SELECT t.*
  FROM src_fin_supplier_invoice_nrm             t
 WHERE t.wd_supplier_id                         IS NOT NULL
   AND COALESCE(t.nrm_transaction_amt, 0)       <> 0
       )
SELECT 
       t.*
  FROM cte_base                                 t
;
