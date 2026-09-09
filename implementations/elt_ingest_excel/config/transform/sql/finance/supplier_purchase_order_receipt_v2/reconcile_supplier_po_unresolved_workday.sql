DROP TABLE IF EXISTS reconcile_supplier_po_unresolved_workday
;
-- Second layer over reconcile_supplier_po_unresolved. Mirrors
-- reconcile_customer_invoice_unresolved_workday exactly:
--
--   GP enrichment (cte_gp_creditor_master / cte_gp_creditor_lookup):
--      Pulls the GP-side address/phone/email/payment-terms/tax from the two
--      GP creditor workbooks (fin_supplier_creditor_new = new creditors
--      delta workbook, and the flat active-supplier extract if any). Matches
--      EITHER on vendor_id = gp_supplier_code (exact) OR on normalised
--      vendor_name (fuzzy), because unresolved PO suppliers can live in the
--      creditor extract under a DIFFERENT vendor_id but with the same master
--      record (the same Simon-Polito pattern, on the supplier side).
--
--   Workday handling:
--     1. Collapses multiple GP supplier codes that name-match the SAME
--        Workday supplier into one row per wd_supplier_id (pipe-concat).
--     2. Splits name matches into MATCH_ADD_REFERENCE_ID (single WD candidate)
--        or AMBIGUOUS_NAME_MATCH (2+ WD candidates, disambiguated by
--        gp_creditor_on_wd_this_ref boolean = vendor_id already in the
--        WD candidate's pipe-split reference_id_current).
CREATE TABLE reconcile_supplier_po_unresolved_workday
    AS
  WITH cte_gp_creditor_master
    AS (
SELECT 'creditor_new'                               gp_creditor_source
     , TRIM(vendor_id)                              gp_creditor_vendor_id
     , TRIM(vendor_name)                            gp_creditor_vendor_name
     , TRIM(company)                                gp_creditor_company
     , CONCAT(
         TRIM(address_1)
       , CASE WHEN NULLIF(TRIM(address_2), '') IS NOT NULL
               THEN CHR(10) || TRIM(address_2) ELSE '' END
       , CASE WHEN NULLIF(TRIM(address_3), '') IS NOT NULL
               THEN CHR(10) || TRIM(address_3) ELSE '' END
       , CASE WHEN NULLIF(TRIM(city), '') IS NOT NULL
               THEN CHR(10) || TRIM(city)        ELSE '' END
       , CASE WHEN NULLIF(TRIM(county), '') IS NOT NULL
               THEN CHR(10) || TRIM(county)      ELSE '' END
       , CASE WHEN NULLIF(TRIM(post_code), '') IS NOT NULL
               THEN CHR(10) || TRIM(post_code)   ELSE '' END
       , CASE WHEN NULLIF(TRIM(country), '') IS NOT NULL
               THEN CHR(10) || TRIM(country)     ELSE '' END
       )                                            gp_creditor_address
     , TRIM(address_1)                              gp_creditor_address_1
     , TRIM(address_2)                              gp_creditor_address_2
     , TRIM(address_3)                              gp_creditor_address_3
     , TRIM(city)                                   gp_creditor_city
     , TRIM(county)                                 gp_creditor_county
     , TRIM(post_code)                              gp_creditor_post_code
     , TRIM(country)                                gp_creditor_country
     , COALESCE(NULLIF(TRIM(phone_number_1), ''),
                NULLIF(TRIM(phone_number_2), ''),
                NULLIF(TRIM(phone_3), ''))          gp_creditor_phone
     , NULLIF(TRIM(email_to_address), '')           gp_creditor_email
     , NULLIF(TRIM(payment_terms_id), '')           gp_creditor_payment_terms
     , COALESCE(NULLIF(TRIM(tax_registration_number), ''),
                NULLIF(TRIM(tax_id_number), ''))    gp_creditor_tax_reg_no
  FROM fin_supplier_creditor_new
 WHERE NULLIF(TRIM(vendor_id), '')                   IS NOT NULL
       )
     , cte_gp_creditor_lookup
    AS (
SELECT r.gp_business_unit
     , r.gp_supplier_code
     , r.gp_supplier_name
     , MIN(g.gp_creditor_source)                   gp_creditor_source
     , MIN(g.gp_creditor_vendor_id)              gp_creditor_vendor_id
     , MIN(g.gp_creditor_company)                   gp_creditor_company
     , MIN(g.gp_creditor_address)                   gp_creditor_address
     , MIN(g.gp_creditor_address_1)                 gp_creditor_address_1
     , MIN(g.gp_creditor_address_2)                 gp_creditor_address_2
     , MIN(g.gp_creditor_address_3)                 gp_creditor_address_3
     , MIN(g.gp_creditor_city)                      gp_creditor_city
     , MIN(g.gp_creditor_county)                    gp_creditor_county
     , MIN(g.gp_creditor_post_code)                 gp_creditor_post_code
     , MIN(g.gp_creditor_country)                   gp_creditor_country
     , MIN(g.gp_creditor_phone)                      gp_creditor_phone
     , MIN(g.gp_creditor_email)                      gp_creditor_email
     , MIN(g.gp_creditor_payment_terms)              gp_creditor_payment_terms
     , MIN(g.gp_creditor_tax_reg_no)                gp_creditor_tax_reg_no
  FROM reconcile_supplier_po_unresolved             r
       LEFT OUTER JOIN
       cte_gp_creditor_master                         g
          ON (
                g.gp_creditor_vendor_id                = TRIM(r.gp_supplier_code)
             OR UPPER(TRIM(g.gp_creditor_vendor_name))
              = UPPER(TRIM(r.gp_supplier_name))
             )
 GROUP BY r.gp_business_unit
        , r.gp_supplier_code
        , r.gp_supplier_name
       )
     , cte_base_match
    AS (
SELECT r.match_status
     , r.gp_business_unit
     , r.gp_supplier_code                             gp_supplier_id
     , r.gp_supplier_name
     , r.wd_supplier_id
     , r.wd_supplier_name
     , r.wd_reference_id_current
     , g.gp_creditor_source
     , g.gp_creditor_vendor_id
     , g.gp_creditor_company
     , g.gp_creditor_address
     , g.gp_creditor_address_1
     , g.gp_creditor_address_2
     , g.gp_creditor_address_3
     , g.gp_creditor_city
     , g.gp_creditor_county
     , g.gp_creditor_post_code
     , g.gp_creditor_country
     , g.gp_creditor_phone
     , g.gp_creditor_email
     , g.gp_creditor_payment_terms
     , g.gp_creditor_tax_reg_no
     , w.primary_remit_to_address                    wd_default_address
     , w.default_phone_number                         wd_default_phone
     , COALESCE(w.country, w.country_for_tax_status) wd_country
     , w.default_currency                           wd_currency
     , w.default_payment_terms                      wd_payment_terms
     , w.supplier_category                          wd_supplier_category
  FROM reconcile_supplier_po_unresolved             r
       LEFT OUTER JOIN
       cte_gp_creditor_lookup                         g
          ON g.gp_business_unit                        = r.gp_business_unit
         AND g.gp_supplier_code                        = r.gp_supplier_code
         AND g.gp_supplier_name                        = r.gp_supplier_name
       LEFT OUTER JOIN
       fin_workday_extract_supplier                   w
          ON w.supplier_id                             = r.wd_supplier_id
       )
     , cte_ambiguity_flag
    AS (
SELECT m.*
     , CASE
           WHEN m.wd_supplier_id                    IS NOT NULL
            AND COUNT(DISTINCT m.wd_supplier_id) OVER (
                    PARTITION BY m.gp_business_unit
                               , m.gp_supplier_id
                               , m.gp_supplier_name
                ) > 1
           THEN TRUE
           ELSE FALSE
       END                                          is_ambiguous
     , CASE
           WHEN m.wd_reference_id_current          IS NOT NULL
            AND m.gp_creditor_vendor_id             IS NOT NULL
            AND EXISTS (
                  SELECT 1
                    FROM (
                          SELECT TRIM(UNNEST(STRING_SPLIT(
                                         m.wd_reference_id_current, '|'
                          ))) AS tok
                    ) w
                   WHERE w.tok = TRIM(m.gp_creditor_vendor_id)
                )
           THEN TRUE
           ELSE FALSE
       END                                          gp_creditor_on_wd_this_ref
  FROM cte_base_match                                 m
       )
     , cte_consolidated_match
    AS (
SELECT 'MATCH_ADD_REFERENCE_ID'                     match_status
     , STRING_AGG(DISTINCT f.gp_business_unit,
                 '|' ORDER BY f.gp_business_unit)   gp_business_unit
     , STRING_AGG(DISTINCT f.gp_supplier_id,
                 '|' ORDER BY f.gp_supplier_id)     gp_supplier_id
     , STRING_AGG(DISTINCT f.gp_supplier_name,
                 '|' ORDER BY f.gp_supplier_name)   gp_supplier_name
     , f.wd_supplier_id
     , MIN(f.wd_supplier_name)                      wd_supplier_name
     , MIN(f.wd_reference_id_current)               wd_reference_id_current
     , MIN(f.gp_creditor_source)                     gp_creditor_source
     , MIN(f.gp_creditor_vendor_id)                  gp_creditor_vendor_id
     , MIN(f.gp_creditor_company)                    gp_creditor_company
     , MIN(f.gp_creditor_address)                    gp_creditor_address
     , MIN(f.gp_creditor_address_1)                  gp_creditor_address_1
     , MIN(f.gp_creditor_address_2)                  gp_creditor_address_2
     , MIN(f.gp_creditor_address_3)                  gp_creditor_address_3
     , MIN(f.gp_creditor_city)                      gp_creditor_city
     , MIN(f.gp_creditor_county)                    gp_creditor_county
     , MIN(f.gp_creditor_post_code)                 gp_creditor_post_code
     , MIN(f.gp_creditor_country)                    gp_creditor_country
     , MIN(f.gp_creditor_phone)                     gp_creditor_phone
     , MIN(f.gp_creditor_email)                     gp_creditor_email
     , MIN(f.gp_creditor_payment_terms)            gp_creditor_payment_terms
     , MIN(f.gp_creditor_tax_reg_no)              gp_creditor_tax_reg_no
     , CAST(NULL AS BOOLEAN)                         gp_creditor_on_wd_this_ref
     , MIN(f.wd_default_address)                    wd_default_address
     , MIN(f.wd_default_phone)                      wd_default_phone
     , MIN(f.wd_country)                            wd_country
     , MIN(f.wd_currency)                           wd_currency
     , MIN(f.wd_payment_terms)                      wd_payment_terms
     , MIN(f.wd_supplier_category)                  wd_supplier_category
  FROM cte_ambiguity_flag                           f
 WHERE f.match_status                               = 'MATCH_ADD_REFERENCE_ID'
   AND f.is_ambiguous                               = FALSE
 GROUP BY f.wd_supplier_id
       )
     , cte_ambiguous_match
    AS (
SELECT 'AMBIGUOUS_NAME_MATCH'                       match_status
     , f.gp_business_unit
     , f.gp_supplier_id
     , f.gp_supplier_name
     , f.wd_supplier_id
     , f.wd_supplier_name
     , f.wd_reference_id_current
     , f.gp_creditor_source
     , f.gp_creditor_vendor_id
     , f.gp_creditor_company
     , f.gp_creditor_address
     , f.gp_creditor_address_1
     , f.gp_creditor_address_2
     , f.gp_creditor_address_3
     , f.gp_creditor_city
     , f.gp_creditor_county
     , f.gp_creditor_post_code
     , f.gp_creditor_country
     , f.gp_creditor_phone
     , f.gp_creditor_email
     , f.gp_creditor_payment_terms
     , f.gp_creditor_tax_reg_no
     , f.gp_creditor_on_wd_this_ref
     , f.wd_default_address
     , f.wd_default_phone
     , f.wd_country
     , f.wd_currency
     , f.wd_payment_terms
     , f.wd_supplier_category
  FROM cte_ambiguity_flag                           f
 WHERE f.is_ambiguous                               = TRUE
 ORDER BY f.gp_business_unit
        , f.gp_supplier_name
        , f.wd_supplier_id
       )
     , cte_genuine_new
    AS (
SELECT 'GENUINE_NEW'                                match_status
     , f.gp_business_unit
     , f.gp_supplier_id
     , f.gp_supplier_name
     , CAST(NULL AS VARCHAR)                        wd_supplier_id
     , CAST(NULL AS VARCHAR)                        wd_supplier_name
     , CAST(NULL AS VARCHAR)                        wd_reference_id_current
     , MIN(f.gp_creditor_source)                     gp_creditor_source
     , MIN(f.gp_creditor_vendor_id)                  gp_creditor_vendor_id
     , MIN(f.gp_creditor_company)                    gp_creditor_company
     , MIN(f.gp_creditor_address)                    gp_creditor_address
     , MIN(f.gp_creditor_address_1)                  gp_creditor_address_1
     , MIN(f.gp_creditor_address_2)                  gp_creditor_address_2
     , MIN(f.gp_creditor_address_3)                  gp_creditor_address_3
     , MIN(f.gp_creditor_city)                      gp_creditor_city
     , MIN(f.gp_creditor_county)                    gp_creditor_county
     , MIN(f.gp_creditor_post_code)                 gp_creditor_post_code
     , MIN(f.gp_creditor_country)                    gp_creditor_country
     , MIN(f.gp_creditor_phone)                     gp_creditor_phone
     , MIN(f.gp_creditor_email)                     gp_creditor_email
     , MIN(f.gp_creditor_payment_terms)            gp_creditor_payment_terms
     , MIN(f.gp_creditor_tax_reg_no)              gp_creditor_tax_reg_no
     , CAST(NULL AS BOOLEAN)                         gp_creditor_on_wd_this_ref
     , CAST(NULL AS VARCHAR)                        wd_default_address
     , CAST(NULL AS VARCHAR)                        wd_default_phone
     , CAST(NULL AS VARCHAR)                        wd_country
     , CAST(NULL AS VARCHAR)                        wd_currency
     , CAST(NULL AS VARCHAR)                        wd_payment_terms
     , CAST(NULL AS VARCHAR)                        wd_supplier_category
  FROM cte_ambiguity_flag                           f
 WHERE f.match_status                               = 'GENUINE_NEW'
 GROUP BY f.gp_business_unit
        , f.gp_supplier_id
        , f.gp_supplier_name
       )
SELECT *
  FROM cte_consolidated_match
 UNION ALL
SELECT *
  FROM cte_ambiguous_match
 UNION ALL
SELECT *
  FROM cte_genuine_new
 ORDER BY match_status
        , gp_business_unit
        , gp_supplier_name
        , wd_supplier_id
;