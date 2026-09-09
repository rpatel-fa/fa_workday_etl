DROP TABLE IF EXISTS src_fin_supplier_purchase_order_nrm
;
CREATE TABLE src_fin_supplier_purchase_order_nrm
    AS
  WITH cte_src_purchase_order_base
    AS (
SELECT SPLIT_PART(t.supplier_code, '-', 1)      nrm_company_name
     , SPLIT_PART(t.supplier_code, '-', 2)      nrm_supplier_code
     , t.*
  FROM src_fin_supplier_purchase_order_base     t
       )
SELECT t.*
     , s.supplier_id                                                       wd_supplier_id
     , s.nrm_supplier_name
     , s.nrm_agg_supplier_number
     , UPPER(TRIM(t.accountingelement1))                                   nrm_accountingelement1
     , UPPER(TRIM(t.accountingelement2))                                   nrm_accountingelement2
     , UPPER(TRIM(t.accountingelement3))                                   nrm_accountingelement3
     , UPPER(TRIM(t.accountingelement4))                                   nrm_accountingelement4
     , CASE
           WHEN UPPER(TRIM(COALESCE(t.currency, ''))) = 'JPY'
           THEN ROUND(CAST(COALESCE(NULLIF(TRIM(t.amount_excl_vat), ''), '0') AS DOUBLE), 0)
           ELSE CAST(COALESCE(NULLIF(TRIM(t.amount_excl_vat), ''), '0') AS DOUBLE)
       END                                                                 nrm_amount_excl_vat
     , strftime(CAST(
       COALESCE(
          TRY_STRPTIME(NULLIF(TRIM(t.po_date), ''), '%Y-%m-%d %H:%M:%S')
        , TRY_STRPTIME(NULLIF(TRIM(t.po_date), ''), '%Y-%m-%d')
        , TRY_STRPTIME(NULLIF(TRIM(t.po_date), ''), '%d-%m-%Y')
       ) AS DATE), '%d/%m/%Y')                                             nrm_document_date
     , CASE (UPPER(TRIM(t.company)))
           WHEN 'FA'   THEN 'THE FA GROUP - FA'
           WHEN 'NFC'  THEN 'THE FA GROUP - NFC'
           WHEN 'WNSL' THEN 'THE FA GROUP - WNSL'
           ELSE 'UNKNOWN'
       END                                                                 nrm_company_id
     , CASE
            WHEN EXISTS (
                 SELECT 1
                   FROM workday_supplier_email  e
                  WHERE e.supplier_id           = s.supplier_id
                    AND e.email_address         IS NOT NULL
                    AND TRIM(e.email_address)   != ''
            )
            THEN 1
            ELSE 0
       END                                                                 nrm_email_exists_ind
  FROM cte_src_purchase_order_base              t
       LEFT OUTER JOIN
       (
       SELECT TRIM(UNNEST(STRING_SPLIT(s.nrm_agg_supplier_number, '|')))   vendor_id
            , s.supplier_id
            , s.nrm_supplier_name
            , s.nrm_agg_supplier_number
         FROM src_fin_supplier                  s
       )                                        s
          ON s.vendor_id                        = TRIM(t.nrm_supplier_code)
 WHERE UPPER(TRIM(t.nrm_company_name))          IN ('FA', 'NFC', 'WNSL')
;
