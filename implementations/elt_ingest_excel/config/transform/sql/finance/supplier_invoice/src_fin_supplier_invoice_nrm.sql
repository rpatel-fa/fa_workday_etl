DROP TABLE IF EXISTS src_fin_supplier_invoice_nrm;

CREATE TABLE src_fin_supplier_invoice_nrm
    AS
  WITH cte_src_supplier_invoice_base
    AS (
SELECT t.* 
  FROM src_fin_supplier_invoice_base         t
       )
SELECT t.*
     , COALESCE(s_exact.supplier_id, s_bare.supplier_id)                         wd_supplier_id
     , COALESCE(s_exact.nrm_supplier_name, s_bare.nrm_supplier_name)             wd_supplier_name
     , COALESCE(s_exact.nrm_agg_supplier_number, s_bare.nrm_agg_supplier_number) wd_supplier_reference_id
     , strftime(CAST(
       COALESCE(
          TRY_STRPTIME(NULLIF(TRIM(t.invoice_date), ''), '%Y-%m-%d %H:%M:%S')
        , TRY_STRPTIME(NULLIF(TRIM(t.invoice_date), ''), '%Y-%m-%d')
        , TRY_STRPTIME(NULLIF(TRIM(t.invoice_date), ''), '%d-%m-%Y')
       ) AS DATE), '%d/%m/%Y')               nrm_document_date
     , CASE (UPPER(TRIM(t.business_unit)))
           WHEN 'FA'   THEN 'THE FA GROUP - FA'
           WHEN 'NFC'  THEN 'THE FA GROUP - NFC'
           WHEN 'WNSL' THEN 'THE FA GROUP - WNSL'
           ELSE 'UNKNOWN'
       END                                   nrm_company_id
     , CASE
         WHEN UPPER(TRIM(t.currency_id))     = 'GBP'
         THEN CAST(COALESCE(NULLIF(TRIM(t.gbp_current_trx_amount), ''), '0') AS DOUBLE)
         WHEN UPPER(TRIM(t.currency_id))     = 'JPY'
         THEN ROUND(CAST(COALESCE(NULLIF(TRIM(t.local_currency_current_transaction_amount), ''), '0') AS DOUBLE), 0)
         ELSE CAST(COALESCE(NULLIF(TRIM(t.local_currency_current_transaction_amount), ''), '0') AS DOUBLE)
       END                                   nrm_transaction_amt
  FROM cte_src_supplier_invoice_base         t
       LEFT OUTER JOIN
       src_fin_workday_supplier_map                                             s_exact
           ON s_exact.reference_id_token     = TRIM(t.vendor_id) || '-' || TRIM(t.business_unit)
       LEFT OUTER JOIN
       src_fin_workday_supplier_map                                             s_bare
           ON s_bare.reference_id_token      = TRIM(t.vendor_id)
;
