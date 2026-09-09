DROP TABLE IF EXISTS validate_supplier_invoice_supplier
;
CREATE TABLE validate_supplier_invoice_supplier 
    AS
SELECT t.business_unit
     , t.vendor_id
     , t.vendor_name
     , t.voucher_number
     , t.document_number
     , t.document_type
     , t.invoice_date
  FROM src_fin_supplier_invoice_nrm       t
 WHERE t.wd_supplier_id                   IS NULL
;
