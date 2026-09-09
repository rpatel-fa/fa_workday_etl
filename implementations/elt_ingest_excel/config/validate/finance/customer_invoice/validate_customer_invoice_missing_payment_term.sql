
DROP TABLE IF EXISTS validate_customer_invoice_missing_payment_term
;
CREATE TABLE validate_customer_invoice_missing_payment_term
    AS
SELECT i.*
  FROM src_fin_customer_invoice_base i
       LEFT OUTER JOIN 
       src_fin_customer_base        c
         ON c.business_unit         = i.business_unit
        AND UPPER(TRIM(c.customer_number)) = UPPER(TRIM(i.customer_id))
 WHERE NULLIF(TRIM(UPPER(c.payment_terms_id)), '') IS NULL
 ORDER BY 2
;
