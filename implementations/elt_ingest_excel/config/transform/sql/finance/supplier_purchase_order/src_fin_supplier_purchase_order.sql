DROP TABLE IF EXISTS src_fin_supplier_purchase_order
;
CREATE TABLE src_fin_supplier_purchase_order
    AS
  WITH cte_base
    AS (    
SELECT t.*
  FROM src_fin_supplier_purchase_order_nrm   t
 WHERE t.wd_supplier_id                      IS NOT NULL
   AND t.nrm_amount_excl_vat                 <> 0
       )
SELECT 
       ROW_NUMBER() OVER (
         ORDER BY 
               t.nrm_company_id
             , t.wd_supplier_id
             , CAST(TRY_STRPTIME(NULLIF(TRIM(t.nrm_document_date), ''), '%d/%m/%Y') AS DATE)
             , t.orderno
             , t.displayno
             , t."22_digit_account_string"
             , t.description
                         )                   spreadsheet_key
     , t.*
  FROM cte_base                              t
;
