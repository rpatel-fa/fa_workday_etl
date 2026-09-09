DROP TABLE IF EXISTS src_fin_supplier_invoice_base;

CREATE TABLE src_fin_supplier_invoice_base
    AS
  WITH cte_supplier_invoice
    AS (
SELECT 'FA'   business_unit, t.* FROM fin_supplier_invoice_fa   t
UNION ALL
SELECT 'NFC'  business_unit, t.* FROM fin_supplier_invoice_nfc  t
UNION ALL
SELECT 'WNSL' business_unit, t.* FROM fin_supplier_invoice_wnsl t
       )
SELECT *
  FROM cte_supplier_invoice            t
 WHERE NULLIF(UPPER(TRIM(t.vendor_id)), '')  IS NOT NULL
;  