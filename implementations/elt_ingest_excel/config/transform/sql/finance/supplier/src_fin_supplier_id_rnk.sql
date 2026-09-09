DROP TABLE IF EXISTS src_fin_supplier_id_rnk
;
CREATE TABLE src_fin_supplier_id_rnk
    AS
  WITH cte_supplier_score
    AS (
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
       )                               score
     , t.*
  FROM src_fin_supplier_raw            t
       )
     , cte_agg
    AS (
SELECT t.nrm_supplier_number
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
     , STRING_AGG(DISTINCT t.nrm_supplier_name , '|' ORDER BY t.nrm_supplier_name )  nrm_agg_supplier_name
     , STRING_AGG(DISTINCT t.phone_number_1     , ';' ORDER BY t.phone_number_1)      nrm_agg_phone_1
     , STRING_AGG(DISTINCT t.phone_number_2     , ';' ORDER BY t.phone_number_2)      nrm_agg_phone_2
     , STRING_AGG(DISTINCT t.phone_3            , ';' ORDER BY t.phone_3)             nrm_agg_phone_3
     , STRING_AGG(DISTINCT t.fax_number         , ';' ORDER BY t.fax_number)          nrm_agg_fax
     , STRING_AGG(DISTINCT t.email_to_address   , ';' ORDER BY t.email_to_address)    nrm_agg_email_to_address
     , STRING_AGG(DISTINCT t.email_cc_address   , ';' ORDER BY t.email_cc_address)    nrm_agg_email_cc_address
     , STRING_AGG(DISTINCT t.email_bcc_address  , ';' ORDER BY t.email_bcc_address)   nrm_agg_email_bcc_address
     , ARRAY_AGG (DISTINCT t.nrm_business_unit        ORDER BY t.nrm_business_unit)   nrm_array_business_unit
     , STRING_AGG(DISTINCT t.nrm_business_unit  , '|' ORDER BY t.nrm_business_unit)   nrm_agg_business_unit
  FROM cte_supplier_score              t
 GROUP BY t.nrm_supplier_number
       )
SELECT
       ROW_NUMBER() OVER(
           PARTITION BY t.nrm_supplier_number
               ORDER BY t.score DESC
                      , t.nrm_supplier_name
                      , t.nrm_business_unit
                      , t.nrm_country_code
       )                                                                          rnk_score
     , COUNT()      OVER(PARTITION BY t.nrm_supplier_number)                        cnt_supplier_number
     , a.nrm_array_address
     , a.nrm_array_bank
     , a.nrm_agg_supplier_name
     , a.nrm_agg_phone_1
     , a.nrm_agg_phone_2
     , a.nrm_agg_phone_3
     , a.nrm_agg_fax
     , a.nrm_agg_email_to_address
     , a.nrm_agg_email_cc_address
     , a.nrm_agg_email_bcc_address
     , a.nrm_array_business_unit
     , a.nrm_agg_business_unit
     , t.*
  FROM cte_supplier_score              t
       INNER JOIN
       cte_agg                         a
          ON a.nrm_supplier_number     = t.nrm_supplier_number
;
