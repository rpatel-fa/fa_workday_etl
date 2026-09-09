DROP TABLE IF EXISTS src_fin_supplier_include
;
CREATE TABLE src_fin_supplier_include
    AS
  WITH cte_src
    AS (
SELECT t.nrm_agg_supplier_number                                  src_agg_reference_id
     , TRIM(UNNEST(STRING_SPLIT(t.nrm_agg_supplier_number, '|'))) src_reference_id
  FROM src_fin_supplier_rnk                                       t
 WHERE t.rnk_score                                                = 1
       )
     , cte_src_token
    AS (
SELECT s.src_agg_reference_id
     , i.nrm_supplier_number                                         src_reference_id
     , i.nrm_vendor_id                                               src_vendor_id
     , UPPER(TRIM(i.business_unit))                                  src_business_unit
     , CASE
           WHEN i.nrm_supplier_number = i.nrm_vendor_id
           THEN i.nrm_vendor_id || '-' || UPPER(TRIM(i.business_unit))
           ELSE i.nrm_supplier_number
       END                                                           src_company_reference_id
     , NULLIF(REGEXP_REPLACE(UPPER(TRIM(COALESCE(i.nrm_bank_account_number, ''))), '[^0-9A-Z]', '', 'g'), '')   src_bank_account_number
     , NULLIF(REGEXP_REPLACE(UPPER(TRIM(COALESCE(i.nrm_bank_sort_code     , ''))), '[^0-9A-Z]', '', 'g'), '')   src_bank_sort_code
     , NULLIF(REGEXP_REPLACE(UPPER(TRIM(COALESCE(i.nrm_bank_iban          , ''))), '[^0-9A-Z]', '', 'g'), '')   src_bank_iban
     , NULLIF(REGEXP_REPLACE(UPPER(TRIM(COALESCE(i.nrm_bank_swift_code    , ''))), '[^0-9A-Z]', '', 'g'), '')   src_bank_swift_code
  FROM cte_src                                                     s
       INNER JOIN
       src_fin_supplier_id_rnk                                     i
          ON i.nrm_supplier_number                                 = s.src_reference_id
         AND i.rnk_score                                           = 1
       )
     , cte_wd
    AS (
SELECT t.supplier_id                                                wd_supplier_id
     , t.reference_id                                               wd_agg_reference_id
     , TRIM(UNNEST(STRING_SPLIT(t.reference_id, '|')))              wd_reference_id
     , NULLIF(REGEXP_REPLACE(UPPER(TRIM(COALESCE(t.account_number           , ''))), '[^0-9A-Z]', '', 'g'), '')   wd_bank_account_number
     , NULLIF(REGEXP_REPLACE(UPPER(TRIM(COALESCE(t.bank_identification_code, ''))), '[^0-9A-Z]', '', 'g'), '')   wd_bank_identification_code
  FROM fin_workday_extract_supplier                                 t
 WHERE NULLIF(TRIM(t.reference_id), '')                             IS NOT NULL
       )
     , cte_token_resolution
    AS (
SELECT s.src_agg_reference_id
     , s.src_reference_id
     , s.src_vendor_id
     , s.src_business_unit
     , s.src_company_reference_id
     , CASE
           WHEN EXISTS (
                SELECT 1
                  FROM cte_wd                                             w
                 WHERE w.wd_reference_id                                  = s.src_company_reference_id
           )
           THEN 'exact_company'
           WHEN EXISTS (
                SELECT 1
                  FROM cte_wd                                             w
                 WHERE w.wd_reference_id                                  = s.src_vendor_id
                   AND (
                          (s.src_bank_account_number IS NOT NULL AND w.wd_bank_account_number        = s.src_bank_account_number)
                       OR (s.src_bank_iban           IS NOT NULL AND w.wd_bank_identification_code   = s.src_bank_iban)
                       OR (s.src_bank_swift_code     IS NOT NULL AND w.wd_bank_identification_code   = s.src_bank_swift_code)
                       OR (s.src_bank_sort_code      IS NOT NULL AND w.wd_bank_identification_code   = s.src_bank_sort_code)
                       )
           )
           THEN 'bare_vendor_same_bank'
           WHEN EXISTS (
                SELECT 1
                  FROM cte_wd                                             w
                 WHERE w.wd_reference_id                                  = s.src_vendor_id
           )
           THEN 'bare_vendor_different_supplier'
           ELSE 'no_match'
       END                                                                match_type
  FROM cte_src_token                                                   s
       )
     , cte_supplier_resolution
    AS (
SELECT t.src_agg_reference_id                                         nrm_agg_supplier_number
     , STRING_AGG(
           DISTINCT
           CASE
               WHEN t.match_type = 'bare_vendor_different_supplier'
               THEN t.src_company_reference_id
               ELSE t.src_reference_id
           END
         , '|'
           ORDER BY
           CASE
               WHEN t.match_type = 'bare_vendor_different_supplier'
               THEN t.src_company_reference_id
               ELSE t.src_reference_id
           END
       )                                                              nrm_resolved_agg_supplier_number
     , STRING_AGG(DISTINCT t.match_type, '|' ORDER BY t.match_type)   match_type
     , MAX(CASE WHEN t.match_type = 'bare_vendor_different_supplier' THEN 1 ELSE 0 END)  suffix_required_ind
     , MIN(CASE
               WHEN t.match_type IN ('exact_company', 'bare_vendor_same_bank')
               THEN 0
               ELSE 1
           END)                                                       include_ind
  FROM cte_token_resolution                                           t
 GROUP BY t.src_agg_reference_id
       )
SELECT t.*
  FROM cte_supplier_resolution                                          t
;
