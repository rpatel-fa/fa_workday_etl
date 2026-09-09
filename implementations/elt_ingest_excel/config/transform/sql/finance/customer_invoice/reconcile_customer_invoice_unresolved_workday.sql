DROP TABLE IF EXISTS reconcile_customer_invoice_unresolved_workday
;
-- Second layer over reconcile_customer_invoice_unresolved that adds both
-- GP-debtor-master enrichment AND Workday-side handling semantics:
--
--   GP enrichment (cte_gp_debtor_master / cte_gp_debtor_lookup):
--      Pulls the GP-side address/phone/email/payment-terms from the two GP
--      debtor workbooks (fin_customer_debtor_delta_1 new-debtor extract, and the
--      flat GP active-customer extract loaded into src_fin_customer_base).
--      Matches EITHER on customer_number = gp_customer_id (exact ID join)
--      OR on normalised customer_name (fuzzy), because the unresolved
--      invoice-only customers (e.g. CWN0616103074 Simon Polito) live in the
--      debtor extract under a DIFFERENT customer_number (CWN0726113776) but
--      share the same name and master record.
--
--   Workday handling:
--     1. Consolidates multiple GP ids that name-match the same Workday customer
--        into one row (gp_customer_id pipe-concatenated), so the review
--        pack shows "one action per Workday customer".
--     2. Splits MATCH_ADD_REFERENCE_ID into:
--          MATCH_ADD_REFERENCE_ID  = exactly one Workday customer per GP id;
--                                    GP debtor enrichment (address/phone) shown
--                                    for spot-check against Workday columns.
--          AMBIGUOUS_NAME_MATCH    = one GP id matched TWO+ Workday customers
--                                    on normalised name. Both Workday rows are
--                                    shown, and the GP debtor enrichment (from
--                                    delta_1 or src_fin_customer_base) is
--                                    compared side-by-side. For Simon Polito
--                                    this reveals that CWN0616103074 name-maps
--                                    to debtor CWN0726113776 in delta_1, and
--                                    CWN0726113776 is already a reference_id
--                                    on Workday C-004240 → hence that is the
--                                    correct WD record to attach to.
--
-- Lineage: reconcile_customer_invoice_unresolved
--       + fin_workday_extract_customer (WD enrichment)
--       + fin_customer_debtor_delta_1  (GP new-debtor enrichment)
--       + src_fin_customer_base      (GP active-debtor enrichment)
CREATE TABLE reconcile_customer_invoice_unresolved_workday
    AS
  WITH cte_gp_debtor_master
    AS (
SELECT 'delta_1'                                    gp_debtor_source
     , TRIM(customer_number)                        gp_debtor_customer_number
     , TRIM(customer_name)                          gp_debtor_customer_name
     , TRIM(company)                                gp_debtor_company
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
       )                                            gp_debtor_address
     , TRIM(address_1)                              gp_debtor_address_1
     , TRIM(address_2)                              gp_debtor_address_2
     , TRIM(address_3)                              gp_debtor_address_3
     , TRIM(city)                                   gp_debtor_city
     , TRIM(county)                                 gp_debtor_county
     , TRIM(post_code)                              gp_debtor_post_code
     , TRIM(country)                                gp_debtor_country
     , TRIM(phone_1)                                gp_debtor_phone
     , TRIM(email_to_address)                         gp_debtor_email
     , TRIM(payment_terms_id)                       gp_debtor_payment_terms
     , TRIM(tax_registration_number)             gp_debtor_tax_reg_no
  FROM fin_customer_debtor_delta_1
 WHERE NULLIF(TRIM(customer_number), '')             IS NOT NULL
 UNION ALL
SELECT 'customer_base'                              gp_debtor_source
     , TRIM(customer_number)                        gp_debtor_customer_number
     , TRIM(customer_name)                          gp_debtor_customer_name
     , CAST(NULL AS VARCHAR)                        gp_debtor_company
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
       )                                            gp_debtor_address
     , TRIM(address_1)                              gp_debtor_address_1
     , TRIM(address_2)                              gp_debtor_address_2
     , TRIM(address_3)                              gp_debtor_address_3
     , TRIM(city)                                   gp_debtor_city
     , TRIM(county)                                 gp_debtor_county
     , TRIM(post_code)                              gp_debtor_post_code
     , TRIM(country)                                gp_debtor_country
     , TRIM(phone_1)                                gp_debtor_phone
     , TRIM(email_to_address)                         gp_debtor_email
     , TRIM(payment_terms_id)                       gp_debtor_payment_terms
     , TRIM(tax_registration_number)             gp_debtor_tax_reg_no
  FROM src_fin_customer_base
 WHERE NULLIF(TRIM(customer_number), '')             IS NOT NULL
       )
     , cte_gp_debtor_lookup
    AS (
SELECT r.gp_business_unit
     , r.gp_customer_id
     , r.gp_customer_name
     , MIN(g.gp_debtor_source)                     gp_debtor_source
     , MIN(g.gp_debtor_customer_number)          gp_debtor_customer_number
     , MIN(g.gp_debtor_company)                   gp_debtor_company
     , MIN(g.gp_debtor_address)                   gp_debtor_address
     , MIN(g.gp_debtor_city)                      gp_debtor_city
     , MIN(g.gp_debtor_county)                    gp_debtor_county
     , MIN(g.gp_debtor_post_code)                 gp_debtor_post_code
     , MIN(g.gp_debtor_country)                   gp_debtor_country
     , MIN(g.gp_debtor_phone)                      gp_debtor_phone
     , MIN(g.gp_debtor_email)                      gp_debtor_email
     , MIN(g.gp_debtor_payment_terms)              gp_debtor_payment_terms
     , MIN(g.gp_debtor_tax_reg_no)                gp_debtor_tax_reg_no
  FROM reconcile_customer_invoice_unresolved        r
       LEFT OUTER JOIN
       cte_gp_debtor_master                         g
          ON (
                g.gp_debtor_customer_number             = TRIM(r.gp_customer_id)
             OR UPPER(TRIM(g.gp_debtor_customer_name))
              = UPPER(TRIM(r.gp_customer_name))
             )
 GROUP BY r.gp_business_unit
        , r.gp_customer_id
        , r.gp_customer_name
       )
     , cte_base_match
    AS (
SELECT r.match_status
     , r.gp_business_unit
     , r.gp_customer_id
     , r.gp_customer_name
     , r.wd_customer_id
     , r.wd_customer_name
     , r.wd_reference_id_current
     , g.gp_debtor_source
     , g.gp_debtor_customer_number
     , g.gp_debtor_company
     , g.gp_debtor_address
     , g.gp_debtor_city
     , g.gp_debtor_county
     , g.gp_debtor_post_code
     , g.gp_debtor_country
     , g.gp_debtor_phone
     , g.gp_debtor_email
     , g.gp_debtor_payment_terms
     , g.gp_debtor_tax_reg_no
     , w.default_address                            wd_default_address
     , w.default_phone_number                       wd_default_phone
     , w.country_for_tax_status                     wd_country
     , w.default_currency                           wd_currency
     , w.default_payment_terms                      wd_payment_terms
     , w.customer_category                          wd_customer_category
  FROM reconcile_customer_invoice_unresolved        r
       LEFT OUTER JOIN
       cte_gp_debtor_lookup                         g
          ON g.gp_business_unit                        = r.gp_business_unit
         AND g.gp_customer_id                          = r.gp_customer_id
         AND g.gp_customer_name                        = r.gp_customer_name
       LEFT OUTER JOIN
       fin_workday_extract_customer                 w
          ON w.customer_id                           = r.wd_customer_id
       )
     , cte_ambiguity_flag
    AS (
SELECT m.*
     , CASE
           WHEN m.wd_customer_id                  IS NOT NULL
            AND COUNT(DISTINCT m.wd_customer_id) OVER (
                    PARTITION BY m.gp_business_unit
                               , m.gp_customer_id
                               , m.gp_customer_name
                ) > 1
           THEN TRUE
           ELSE FALSE
       END                                          is_ambiguous
     , CASE
           WHEN m.wd_reference_id_current          IS NOT NULL
            AND m.gp_debtor_customer_number         IS NOT NULL
            AND EXISTS (
                  SELECT 1
                    FROM (
                          SELECT TRIM(UNNEST(STRING_SPLIT(
                                         m.wd_reference_id_current, '|'
                          ))) AS tok
                    ) w
                   WHERE w.tok = TRIM(m.gp_debtor_customer_number)
                )
           THEN TRUE
           ELSE FALSE
       END                                          gp_debtor_on_wd_this_ref
  FROM cte_base_match                               m
       )
     , cte_consolidated_match
    AS (
SELECT 'MATCH_ADD_REFERENCE_ID'                     match_status
     , STRING_AGG(DISTINCT f.gp_business_unit,
                 '|' ORDER BY f.gp_business_unit)   gp_business_unit
     , STRING_AGG(DISTINCT f.gp_customer_id,
                 '|' ORDER BY f.gp_customer_id)     gp_customer_id
     , STRING_AGG(DISTINCT f.gp_customer_name,
                 '|' ORDER BY f.gp_customer_name)   gp_customer_name
     , f.wd_customer_id
     , MIN(f.wd_customer_name)                      wd_customer_name
     , MIN(f.wd_reference_id_current)               wd_reference_id_current
     , MIN(f.gp_debtor_source)                       gp_debtor_source
     , MIN(f.gp_debtor_customer_number)              gp_debtor_customer_number
     , MIN(f.gp_debtor_company)                      gp_debtor_company
     , MIN(f.gp_debtor_address)                       gp_debtor_address
     , MIN(f.gp_debtor_city)                        gp_debtor_city
     , MIN(f.gp_debtor_county)                      gp_debtor_county
     , MIN(f.gp_debtor_post_code)                   gp_debtor_post_code
     , MIN(f.gp_debtor_country)                      gp_debtor_country
     , MIN(f.gp_debtor_phone)                       gp_debtor_phone
     , MIN(f.gp_debtor_email)                       gp_debtor_email
     , MIN(f.gp_debtor_payment_terms)              gp_debtor_payment_terms
     , MIN(f.gp_debtor_tax_reg_no)                gp_debtor_tax_reg_no
     , CAST(NULL AS BOOLEAN)                         gp_debtor_on_wd_this_ref
     , MIN(f.wd_default_address)                    wd_default_address
     , MIN(f.wd_default_phone)                      wd_default_phone
     , MIN(f.wd_country)                            wd_country
     , MIN(f.wd_currency)                           wd_currency
     , MIN(f.wd_payment_terms)                      wd_payment_terms
     , MIN(f.wd_customer_category)                  wd_customer_category
  FROM cte_ambiguity_flag                           f
 WHERE f.match_status                               = 'MATCH_ADD_REFERENCE_ID'
   AND f.is_ambiguous                               = FALSE
 GROUP BY f.wd_customer_id
       )
     , cte_ambiguous_match
    AS (
SELECT 'AMBIGUOUS_NAME_MATCH'                       match_status
     , f.gp_business_unit
     , f.gp_customer_id
     , f.gp_customer_name
     , f.wd_customer_id
     , f.wd_customer_name
     , f.wd_reference_id_current
     , f.gp_debtor_source
     , f.gp_debtor_customer_number
     , f.gp_debtor_company
     , f.gp_debtor_address
     , f.gp_debtor_city
     , f.gp_debtor_county
     , f.gp_debtor_post_code
     , f.gp_debtor_country
     , f.gp_debtor_phone
     , f.gp_debtor_email
     , f.gp_debtor_payment_terms
     , f.gp_debtor_tax_reg_no
     , f.gp_debtor_on_wd_this_ref
     , f.wd_default_address
     , f.wd_default_phone
     , f.wd_country
     , f.wd_currency
     , f.wd_payment_terms
     , f.wd_customer_category
  FROM cte_ambiguity_flag                           f
 WHERE f.is_ambiguous                               = TRUE
 ORDER BY f.gp_business_unit
        , f.gp_customer_name
        , f.wd_customer_id
       )
     , cte_genuine_new
    AS (
SELECT 'GENUINE_NEW'                                match_status
     , f.gp_business_unit
     , f.gp_customer_id
     , f.gp_customer_name
     , CAST(NULL AS VARCHAR)                        wd_customer_id
     , CAST(NULL AS VARCHAR)                        wd_customer_name
     , CAST(NULL AS VARCHAR)                        wd_reference_id_current
     , MIN(f.gp_debtor_source)                       gp_debtor_source
     , MIN(f.gp_debtor_customer_number)              gp_debtor_customer_number
     , MIN(f.gp_debtor_company)                      gp_debtor_company
     , MIN(f.gp_debtor_address)                       gp_debtor_address
     , MIN(f.gp_debtor_city)                        gp_debtor_city
     , MIN(f.gp_debtor_county)                      gp_debtor_county
     , MIN(f.gp_debtor_post_code)                   gp_debtor_post_code
     , MIN(f.gp_debtor_country)                      gp_debtor_country
     , MIN(f.gp_debtor_phone)                       gp_debtor_phone
     , MIN(f.gp_debtor_email)                       gp_debtor_email
     , MIN(f.gp_debtor_payment_terms)              gp_debtor_payment_terms
     , MIN(f.gp_debtor_tax_reg_no)                gp_debtor_tax_reg_no
     , CAST(NULL AS BOOLEAN)                         gp_debtor_on_wd_this_ref
     , CAST(NULL AS VARCHAR)                        wd_default_address
     , CAST(NULL AS VARCHAR)                        wd_default_phone
     , CAST(NULL AS VARCHAR)                        wd_country
     , CAST(NULL AS VARCHAR)                        wd_currency
     , CAST(NULL AS VARCHAR)                        wd_payment_terms
     , CAST(NULL AS VARCHAR)                        wd_customer_category
  FROM cte_ambiguity_flag                           f
 WHERE f.match_status                               = 'GENUINE_NEW'
 GROUP BY f.gp_business_unit
        , f.gp_customer_id
        , f.gp_customer_name
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
        , gp_customer_name
        , wd_customer_id
;