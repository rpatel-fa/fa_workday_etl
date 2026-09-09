DROP TABLE IF EXISTS src_fin_supplier_po_receipt
;
CREATE TABLE src_fin_supplier_po_receipt
    AS
  WITH cte_base
    AS (
SELECT t.*
  FROM src_fin_supplier_po_receipt_nrm                              t
 WHERE t.nrm_row_status                                            = 'KEEP'
   AND t.wd_supplier_id                                            IS NOT NULL
   AND t.receipted_quantity_text                                   != 'NULL'
   AND COALESCE(t.nrm_receipt_net_gbp, 0)                          >  0.005
       )
SELECT
       ROW_NUMBER() OVER (
         ORDER BY
               t.nrm_company_id
             , t.wd_supplier_id
             , t.order_number
             , t.receipt_number
             , t.line_number
             , t.split_position
       )                                                           spreadsheet_key
     , t.*
  FROM cte_base                                                     t
;
