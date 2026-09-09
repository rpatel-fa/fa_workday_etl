DROP TABLE IF EXISTS src_fin_customer_invoice;

CREATE TABLE src_fin_customer_invoice
    AS
  WITH cte_base
    AS (    
SELECT t.*
  FROM src_fin_customer_invoice_nrm             t
 WHERE t.wd_customer_id                         IS NOT NULL
   AND COALESCE(CAST(t.nrm_transaction_amt AS DOUBLE), 0) <> 0
       )
SELECT 
       t.*
  FROM cte_base                                 t
;
