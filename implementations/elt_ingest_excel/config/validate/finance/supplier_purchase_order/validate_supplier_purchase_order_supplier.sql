DROP TABLE IF EXISTS validate_supplier_purchase_order_supplier
;
CREATE TABLE validate_supplier_purchase_order_supplier 
    AS
SELECT t.nrm_company_name
     , t.nrm_supplier_code
     , t.company
     , t.supplier_code
     , t.supplier_title
     , t.orderno
     , t.displayno
     , t.po_date
  FROM src_fin_supplier_purchase_order_nrm   t
 WHERE t.wd_supplier_id                      IS NULL
;
