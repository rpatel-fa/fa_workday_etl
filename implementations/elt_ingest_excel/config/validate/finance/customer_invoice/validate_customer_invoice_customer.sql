DROP TABLE IF EXISTS validate_customer_invoice_customer
;
CREATE TABLE validate_customer_invoice_customer
    AS
SELECT t.business_unit
     , t.customer_id
     , t.customer_name
     , t.document_number
     , t.document_type
     , t.document_date
  FROM src_fin_customer_invoice_nrm    t
 WHERE t.wd_customer_id                IS NULL
;
