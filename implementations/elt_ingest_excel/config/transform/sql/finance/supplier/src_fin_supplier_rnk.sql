DROP TABLE IF EXISTS src_fin_supplier_rnk
;
CREATE TABLE src_fin_supplier_rnk
    AS
  WITH cte_canonical_name
    AS (
-- One row per nrm_supplier_number: the supplier name from the highest-scoring raw record.
-- All other rows for the same number are re-assigned to this canonical name so that
-- Level 2 (name-based) aggregation produces a single record per supplier number.
SELECT t.nrm_supplier_number
     , t.nrm_supplier_name     canonical_name
  FROM src_fin_supplier_id_rnk t
 WHERE t.rnk_score             = 1
       )
     , cte_supplier_score
    AS (
-- Strip Level-1 aggregate columns and rank columns before re-scoring at the name level.
-- The scalar nrm_* fields carried through from src_fin_supplier_raw are used for scoring
-- and re-aggregation below.
SELECT
       (
         (CASE WHEN t.nrm_payment_terms_id         IS NOT NULL THEN 100 ELSE 0 END)
       + (CASE WHEN t.nrm_bank_name                IS NOT NULL THEN  90 ELSE 0 END)
       + (CASE WHEN t.nrm_tax_schedule_id          IS NOT NULL AND t.nrm_tax_schedule_id = 'PS20'  THEN 80 ELSE 0 END)
       + (CASE WHEN t.nrm_tax_id_number            IS NOT NULL THEN  70 ELSE 0 END)
       + (CASE WHEN t.nrm_region                   IS NOT NULL THEN  60 ELSE 0 END)
       + (CASE WHEN t.nrm_postal_code              IS NOT NULL THEN  50 ELSE 0 END)
       + (CASE WHEN t.nrm_address_line_1           IS NOT NULL THEN  40 ELSE 0 END)
       + (CASE WHEN t.nrm_address_line_2           IS NOT NULL THEN  30 ELSE 0 END)
       + (CASE WHEN t.nrm_address_line_3           IS NOT NULL THEN  20 ELSE 0 END)
       + (CASE WHEN t.nrm_address_line_4           IS NOT NULL THEN  15 ELSE 0 END)
       + (CASE WHEN t.nrm_city                     IS NOT NULL THEN  10 ELSE 0 END)
       )                                                                              score
     , c.canonical_name
     , t.* EXCLUDE (rnk_score, cnt_supplier_number, score,
                    nrm_supplier_name,
                    nrm_agg_supplier_name,
                    nrm_array_address, nrm_array_bank,
                    nrm_agg_phone_1, nrm_agg_phone_2, nrm_agg_phone_3,
                    nrm_agg_fax,
                    nrm_agg_email_to_address, nrm_agg_email_cc_address, nrm_agg_email_bcc_address,
                    nrm_array_business_unit, nrm_agg_business_unit)
  FROM src_fin_supplier_id_rnk         t
       INNER JOIN
       cte_canonical_name              c
          ON c.nrm_supplier_number     = t.nrm_supplier_number
       )
     , cte_name_accounts
    AS (
-- Distinct bank accounts under each canonical name. NULLs are ignored by COUNT
-- DISTINCT, so a name with a single account plus blank rows still counts as one.
SELECT t.canonical_name
     , COUNT(DISTINCT t.nrm_bank_account_number)  cnt_distinct_accounts
  FROM cte_supplier_score              t
 GROUP BY t.canonical_name
       )
     , cte_vendor_rep
    AS (
-- Representative bank account per GP vendor_id: the account on that vendor_id's
-- highest-scoring row (ARG_MAX ignores rows with no account). Used only to decide
-- which vendor_ids belong to the same payee within a conflicting name.
SELECT t.canonical_name
     , t.nrm_vendor_id
     , ARG_MAX(t.nrm_bank_account_number, t.score)  rep_bank_account
  FROM cte_supplier_score              t
 GROUP BY t.canonical_name
        , t.nrm_vendor_id
       )
     , cte_payee_key
    AS (
-- Split a shared canonical name back into distinct payees when the merged GP
-- vendor_ids disagree on bank account. Workday does not require unique supplier
-- names (identity is the reference ID), so keeping genuinely different payees
-- separate is safe. A single vendor_id always stays together (one real supplier
-- may hold several accounts); only DIFFERENT vendor_ids sharing the name are
-- separated, and only when the name actually has >1 distinct account. Vendor_ids
-- that share the same representative account still merge (true duplicates).
--   Consequence if removed: unrelated payees (e.g. two different 'JOSH SMITH's)
--   collapse into one supplier carrying both parties' bank accounts -> payment
--   misrouting and multiple-settlement-account EIB rejections.
SELECT t.*
     , CASE
         WHEN na.cnt_distinct_accounts <= 1
         THEN t.canonical_name
         ELSE t.canonical_name || '#' || COALESCE(vr.rep_bank_account, t.nrm_vendor_id)
       END                                                                            payee_key
  FROM cte_supplier_score              t
       INNER JOIN
       cte_name_accounts               na
          ON na.canonical_name         = t.canonical_name
       LEFT OUTER JOIN
       cte_vendor_rep                  vr
          ON vr.canonical_name         = t.canonical_name
         AND vr.nrm_vendor_id          = t.nrm_vendor_id
       )
     , cte_agg
    AS (
SELECT t.payee_key
     , ARRAY_AGG(DISTINCT STRUCT_PACK(
           nrm_address_line_1    := fn_initcap(t.nrm_address_line_1)
         , nrm_address_line_2    := fn_initcap(t.nrm_address_line_2)
         , nrm_address_line_3    := fn_initcap(t.nrm_address_line_3)
         , nrm_address_line_4    := fn_initcap(t.nrm_address_line_4)
         , nrm_city              := fn_initcap(t.nrm_city)
         , nrm_city_subdivision  := fn_initcap(t.nrm_city_subdivision)
         , nrm_subregion_county  := fn_initcap(t.nrm_subregion_county)
         , nrm_region            := t.nrm_region
         , nrm_country_name      := t.nrm_country_name
         , nrm_country_code      := t.nrm_country_code
         , nrm_postal_code       := t.nrm_postal_code
         , nrm_address_code      := t.nrm_address_code
         , nrm_postal_code_valid := t.nrm_postal_code_valid
       ))                                                                              nrm_array_address
     , ARRAY_AGG(DISTINCT STRUCT_PACK(
           nrm_bank_account_type          := 'Checking'
         , nrm_bank_name                  := t.nrm_bank_name
         , nrm_bank_country_name          := t.nrm_bank_country_name
         , nrm_bank_sort_code             := t.nrm_bank_sort_code
         , nrm_bank_account_name          := t.nrm_bank_account_name
         , nrm_bank_account_nickname      := t.nrm_bank_account_name
         , nrm_bank_account_number        := t.nrm_bank_account_number
         , nrm_bank_code_routing_number   := t.nrm_bank_code_routing_number
         , nrm_bank_iban                  := t.nrm_bank_iban
         , nrm_bank_swift_code            := t.nrm_bank_swift_code
         , nrm_bank_roll_number           := t.nrm_bank_roll_number
         , nrm_bank_check_digit           := t.nrm_bank_check_digit
         , nrm_bank_branch_id             := t.nrm_bank_branch_id
         , nrm_bank_branch_name           := t.nrm_bank_branch_name
         , nrm_bank_transfer_method       := t.nrm_bank_transfer_method
       ))                                                                              nrm_array_bank
     , STRING_AGG(DISTINCT t.nrm_supplier_number, '|' ORDER BY t.nrm_supplier_number)  nrm_agg_supplier_number
     , STRING_AGG(DISTINCT t.phone_number_1     , ';' ORDER BY t.phone_number_1)       nrm_agg_phone_1
     , STRING_AGG(DISTINCT t.phone_number_2     , ';' ORDER BY t.phone_number_2)       nrm_agg_phone_2
     , STRING_AGG(DISTINCT t.phone_3            , ';' ORDER BY t.phone_3)              nrm_agg_phone_3
     , STRING_AGG(DISTINCT t.fax_number         , ';' ORDER BY t.fax_number)           nrm_agg_fax
     , STRING_AGG(DISTINCT t.email_to_address   , ';' ORDER BY t.email_to_address)     nrm_agg_email_to_address
     , STRING_AGG(DISTINCT t.email_cc_address   , ';' ORDER BY t.email_cc_address)     nrm_agg_email_cc_address
     , STRING_AGG(DISTINCT t.email_bcc_address  , ';' ORDER BY t.email_bcc_address)    nrm_agg_email_bcc_address
     , ARRAY_AGG (DISTINCT t.nrm_business_unit        ORDER BY t.nrm_business_unit)    nrm_array_business_unit
     , STRING_AGG(DISTINCT t.nrm_business_unit  , '|' ORDER BY t.nrm_business_unit)    nrm_agg_business_unit
  FROM cte_payee_key                   t
 GROUP BY t.payee_key
       )
SELECT
       ROW_NUMBER() OVER(
           PARTITION BY t.payee_key
               ORDER BY t.score DESC
                      , t.nrm_supplier_number
                      , t.nrm_business_unit
                      , t.nrm_country_code
       )                                                                      rnk_score
     , COUNT()      OVER(PARTITION BY t.payee_key)                             cnt_supplier_name
     , a.nrm_array_address
     , a.nrm_array_bank
     , a.nrm_agg_supplier_number
     , a.nrm_agg_phone_1
     , a.nrm_agg_phone_2
     , a.nrm_agg_phone_3
     , a.nrm_agg_fax
     , a.nrm_agg_email_to_address
     , a.nrm_agg_email_cc_address
     , a.nrm_agg_email_bcc_address
     , a.nrm_array_business_unit
     , a.nrm_agg_business_unit
     , t.canonical_name              nrm_supplier_name
     , t.*  EXCLUDE (canonical_name, payee_key)
  FROM cte_payee_key                  t
       INNER JOIN
       cte_agg                        a
          ON a.payee_key              = t.payee_key
;
