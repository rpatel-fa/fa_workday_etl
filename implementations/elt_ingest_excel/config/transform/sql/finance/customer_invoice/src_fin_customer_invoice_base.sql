DROP TABLE IF EXISTS src_fin_customer_invoice_base;

CREATE TABLE src_fin_customer_invoice_base
    AS
  WITH cte_customer_invoice
    AS (
SELECT 'FA'                      business_unit
     , t.*
     , CAST(t.customer_number AS VARCHAR)                  customer_id
     , t.document_amount                                    current_transaction_amount_gbp
     , t.local_currency_document_amount                     current_transaction_amount_in_originating_currncy
     , CAST(NULL AS VARCHAR)                               exchange_rate
  FROM fin_customer_invoice_fa   t
UNION ALL
SELECT 'NFC'                     business_unit
     , t.*
     , CAST(t.customer_number AS VARCHAR)                  customer_id
     , t.document_amount                                    current_transaction_amount_gbp
     , t.local_currency_document_amount                     current_transaction_amount_in_originating_currncy
     , CAST(NULL AS VARCHAR)                               exchange_rate
  FROM fin_customer_invoice_nfc  t
UNION ALL
SELECT 'WNSL'                    business_unit
     , t.*
     , CAST(t.customer_number AS VARCHAR)                  customer_id
     , t.document_amount                                    current_transaction_amount_gbp
     , t.local_currency_document_amount                     current_transaction_amount_in_originating_currncy
     , CAST(NULL AS VARCHAR)                               exchange_rate
  FROM fin_customer_invoice_wnsl t
       )
SELECT *
  FROM cte_customer_invoice            t
 WHERE NULLIF(UPPER(TRIM(t.customer_id)), '')  IS NOT NULL
;
