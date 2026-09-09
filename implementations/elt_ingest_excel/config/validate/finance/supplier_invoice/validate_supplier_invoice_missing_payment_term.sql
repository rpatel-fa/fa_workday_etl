
DROP TABLE IF EXISTS validate_customer_invoice_missing_payment_term
;
CREATE TABLE validate_customer_invoice_missing_payment_term
    AS
SELECT i.*
  FROM src_fin_supplier_invoice_base   i
       LEFT OUTER JOIN 
       src_fin_supplier_base           s
         ON s.business_unit            = i.business_unit
        AND UPPER(TRIM(s.vendor_id))   = UPPER(TRIM(i.vendor_id))
 WHERE NULLIF(TRIM(UPPER(s.payment_terms_id)), '') IS NULL
 ORDER BY 1,2
;
