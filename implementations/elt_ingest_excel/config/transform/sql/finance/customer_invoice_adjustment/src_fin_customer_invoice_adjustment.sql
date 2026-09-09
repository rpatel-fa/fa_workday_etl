DROP TABLE IF EXISTS src_fin_customer_invoice_adjustment
;
-- Normalised source for the AR Customer Invoice Adjustment EIB.
--
-- Unlike the customer_invoice pipeline (three GP trial-balance sheets that must
-- be unioned and resolved against the Workday customer extract), this source is
-- a PRE-RESOLVED adjustment worklist: each row already carries the Workday
-- customer id (`workday_id`), the base invoice (`document_number`), the required
-- adjustment transaction (`wd_adjustment_transaction_needed`), a CREDIT/DEBIT
-- Type and the signed adjustment amount. No customer-map join is needed - every
-- workday_id resolves 1:1 in fin_workday_extract_customer (verified 21/08/26).
--
-- This transform therefore only: drops the trailing blank rows, derives the
-- Workday company id, formats the adjustment date, extracts the adjustment id
-- from the "Customer Invoice Adjustment: <id>" text, and exposes the amount both
-- as a signed DOUBLE (for the increase-amount-due direction) and as an unsigned
-- string (the EIB Extended Amount, which is always positive - direction is
-- carried by the Increase Amount Due flag).
CREATE TABLE src_fin_customer_invoice_adjustment
    AS
  WITH cte_base
    AS (
SELECT t.*
     , TRIM(t.workday_id)                                        wd_customer_id
     , TRIM(t.currency_id)                                       nrm_currency_id
     , CASE UPPER(TRIM(t.company))
           WHEN 'FA'   THEN 'THE FA GROUP - FA'
           WHEN 'NFC'  THEN 'THE FA GROUP - NFC'
           WHEN 'WNSL' THEN 'THE FA GROUP - WNSL'
           ELSE 'UNKNOWN'
       END                                                       nrm_company_id
     , strftime(CAST(
         COALESCE(
            TRY_STRPTIME(NULLIF(TRIM(t.document_date), ''), '%Y-%m-%d %H:%M:%S')
          , TRY_STRPTIME(NULLIF(TRIM(t.document_date), ''), '%Y-%m-%d')
         ) AS DATE), '%d/%m/%Y')                                 nrm_document_date
     , TRIM(SPLIT_PART(t.wd_adjustment_transaction_needed, ': ', 2)) nrm_adjustment_id
     , CAST(COALESCE(NULLIF(TRIM(t.adjustment_amount), ''), '0') AS DOUBLE) nrm_adjustment_amt
     , REGEXP_REPLACE(COALESCE(NULLIF(TRIM(t.adjustment_amount), ''), '0'), '^-', '') nrm_extended_amount
  FROM fin_customer_invoice_adjustment                           t
 WHERE NULLIF(TRIM(t.workday_id), '')                            IS NOT NULL
   AND NULLIF(TRIM(t.wd_adjustment_transaction_needed), '')      IS NOT NULL
       )
SELECT t.*
  FROM cte_base                                                  t
 WHERE COALESCE(t.nrm_adjustment_amt, 0)                         <> 0
;
