DROP TABLE IF EXISTS src_fin_supplier_po_commitment_stack_nrm;

CREATE TABLE src_fin_supplier_po_commitment_stack_nrm
    AS
  WITH cte_src_commitment_base
    AS (
SELECT SPLIT_PART(t.supplier_code, '-', 1)                                        nrm_company_name
     , CASE
           WHEN UPPER(TRIM(t.company_shortcode)) = 'FA'
            AND TRIM(t.order_number)       = 'FAPO200025244'
            AND TRIM(t.supplier_code)     = 'FA-STE096'
           THEN 'STE090'
           WHEN UPPER(TRIM(t.company_shortcode)) = 'NFC'
            AND TRIM(t.supplier_code)     = 'NFC-NAT008'
           THEN 'NAT015'
           ELSE SPLIT_PART(t.supplier_code, '-', 2)
       END                                                                        nrm_supplier_code
     , t.*
  FROM src_fin_supplier_po_commitment_stack_base                                  t
),
cte_joined
    AS (
SELECT
       t.*
     , COALESCE(s_exact.supplier_id, s_bare.supplier_id)                         wd_supplier_id
     , COALESCE(s_exact.nrm_supplier_name, s_bare.nrm_supplier_name)             wd_supplier_name
     , COALESCE(s_exact.nrm_agg_supplier_number, s_bare.nrm_agg_supplier_number) wd_supplier_reference_id
     , COALESCE(s_exact.nrm_email_exists_ind, s_bare.nrm_email_exists_ind, 0)    wd_supplier_email_exists_ind
     , CASE
          WHEN TRY_CAST(NULLIF(t.order_date_excel, '') AS TIMESTAMP) IS NOT NULL
          THEN CAST(TRY_CAST(NULLIF(t.order_date_excel, '') AS TIMESTAMP) AS DATE)
          WHEN TRY_CAST(NULLIF(t.order_date_excel, '') AS DOUBLE) IS NOT NULL
          THEN DATE '1899-12-30' + CAST(FLOOR(TRY_CAST(NULLIF(t.order_date_excel, '') AS DOUBLE)) AS INTEGER)
          ELSE NULL
       END                                                                        order_date
     , CASE
          WHEN TRY_CAST(NULLIF(t.commitment_date_excel, '') AS TIMESTAMP) IS NOT NULL
          THEN CAST(TRY_CAST(NULLIF(t.commitment_date_excel, '') AS TIMESTAMP) AS DATE)
          WHEN TRY_CAST(NULLIF(t.commitment_date_excel, '') AS DOUBLE) IS NOT NULL
          THEN DATE '1899-12-30' + CAST(FLOOR(TRY_CAST(NULLIF(t.commitment_date_excel, '') AS DOUBLE)) AS INTEGER)
          ELSE NULL
       END                                                                        commitment_date
     , UPPER(TRIM(t.accountingelement1))                                          nrm_accountingelement1
     , UPPER(TRIM(t.accountingelement2))                                          nrm_accountingelement2
     , UPPER(TRIM(t.accountingelement3))                                          nrm_accountingelement3
     , UPPER(TRIM(t.accountingelement4))                                          nrm_accountingelement4
     , CASE
           WHEN COALESCE(t.soft_commitment_value_gbp, 0) > 0
            AND COALESCE(t.hard_commitment_value_gbp, 0) > 0
           THEN 'MIXED'
           WHEN COALESCE(t.hard_commitment_value_gbp, 0) > 0
           THEN 'HARD'
           WHEN COALESCE(t.soft_commitment_value_gbp, 0) > 0
           THEN 'SOFT'
           ELSE 'NONE'
       END                                                                        commitment_class
     , CASE (UPPER(TRIM(t.company_shortcode)))
           WHEN 'FA'   THEN 'THE FA GROUP - FA'
           WHEN 'NFC'  THEN 'THE FA GROUP - NFC'
           WHEN 'WNSL' THEN 'THE FA GROUP - WNSL'
           ELSE 'UNKNOWN'
       END                                                                        nrm_company_id
  FROM cte_src_commitment_base                                                    t
       LEFT OUTER JOIN
       src_fin_workday_supplier_map                                               s_exact
          ON s_exact.reference_id_token                                           = t.nrm_supplier_code || '-' || t.nrm_company_name
       LEFT OUTER JOIN
       src_fin_workday_supplier_map                                               s_bare
          ON s_bare.reference_id_token                                            = t.nrm_supplier_code
)
SELECT j.*
  FROM cte_joined                                                                  j
;
