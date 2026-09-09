
DROP TABLE IF EXISTS validate_customer_invoice_missing_payment_term
;
CREATE TABLE validate_customer_invoice_missing_payment_term
    AS
SELECT i.*
  FROM src_fin_supplier_purchase_order_base  i
       LEFT OUTER JOIN 
       src_fin_supplier_base                 s
         ON UPPER(TRIM(s.business_unit))                       = UPPER(TRIM(i.company))
        AND UPPER(TRIM(s.business_unit || '-' ||s.vendor_id))  = UPPER(TRIM(i.supplier_code))
 WHERE NULLIF(TRIM(UPPER(s.payment_terms_id)), '') IS NULL
 ORDER BY i.supplier_code
;
