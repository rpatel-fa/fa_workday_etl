DROP TABLE IF EXISTS src_fin_customer_id_rnk
;
CREATE TABLE src_fin_customer_id_rnk
    AS
  WITH cte_customer_score
    AS (
SELECT
       (
         (CASE WHEN t.nrm_payment_terms_id         IS NOT NULL THEN 100 ELSE 0 END)
       + (CASE WHEN t.nrm_tax_schedule_id          IS NOT NULL AND t.nrm_tax_schedule_id = 'SS20'  THEN 90 ELSE 0 END)
       + (CASE WHEN t.nrm_tax_schedule_id          IS NOT NULL AND t.nrm_tax_schedule_id <> 'SS20' THEN 80 ELSE 0 END)
       + (CASE WHEN t.nrm_tax_registration_number  IS NOT NULL THEN  70 ELSE 0 END)
       + (CASE WHEN t.nrm_region                   IS NOT NULL THEN  60 ELSE 0 END)
       + (CASE WHEN t.nrm_postal_code              IS NOT NULL THEN  50 ELSE 0 END)
       + (CASE WHEN t.nrm_address_line_1           IS NOT NULL THEN  40 ELSE 0 END)
       + (CASE WHEN t.nrm_address_line_2           IS NOT NULL THEN  30 ELSE 0 END)
       + (CASE WHEN t.nrm_address_line_3           IS NOT NULL THEN  20 ELSE 0 END)
       + (CASE WHEN t.nrm_city                     IS NOT NULL THEN  10 ELSE 0 END)
       )                                                                              score
     , t.*
  FROM src_fin_customer_raw            t
       )
     , cte_agg
    AS (
SELECT t.nrm_customer_number
     , ARRAY_AGG(DISTINCT STRUCT_PACK(
           nrm_address_line_1    := fn_initcap(t.nrm_address_line_1)
         , nrm_address_line_2    := fn_initcap(t.nrm_address_line_2)
         , nrm_address_line_3    := fn_initcap(t.nrm_address_line_3)
         , nrm_address_line_4    := fn_initcap(t.nrm_address_line_4)
         , nrm_city_subdivision  := fn_initcap(t.nrm_city_subdivision)
         , nrm_subregion_county  := fn_initcap(t.nrm_subregion_county)
         , nrm_city              := fn_initcap(t.nrm_city)
         , nrm_region            := t.nrm_region
         , nrm_country_name      := t.nrm_country_name
         , nrm_country_code      := t.nrm_country_code
         , nrm_postal_code       := t.nrm_postal_code
         , nrm_address_code      := t.nrm_address_code
         , nrm_postal_code_valid := t.nrm_postal_code_valid
       ))                                                                              nrm_array_customer_address
     , STRING_AGG(DISTINCT t.nrm_customer_name,   '|' ORDER BY t.nrm_customer_name)    nrm_agg_customer_name
     , STRING_AGG(DISTINCT t.phone_1            , ';' ORDER BY t.phone_1)              nrm_agg_phone_1
     , STRING_AGG(DISTINCT t.phone_2            , ';' ORDER BY t.phone_2)              nrm_agg_phone_2
     , STRING_AGG(DISTINCT t.phone_3            , ';' ORDER BY t.phone_3)              nrm_agg_phone_3
     , STRING_AGG(DISTINCT t.fax                , ';' ORDER BY t.fax)                  nrm_agg_fax
     , STRING_AGG(DISTINCT t.email_to_address   , ';' ORDER BY t.email_to_address)     nrm_agg_email_to_address
     , STRING_AGG(DISTINCT t.email_cc_address   , ';' ORDER BY t.email_cc_address)     nrm_agg_email_cc_address
     , STRING_AGG(DISTINCT t.email_bcc_address  , ';' ORDER BY t.email_bcc_address)    nrm_agg_email_bcc_address
     , ARRAY_AGG (DISTINCT t.nrm_business_unit        ORDER BY t.nrm_business_unit)    nrm_array_business_unit
     , STRING_AGG(DISTINCT t.nrm_business_unit  , '|' ORDER BY t.nrm_business_unit)    nrm_agg_business_unit
  FROM cte_customer_score              t
 GROUP BY t.nrm_customer_number
       )
SELECT
       ROW_NUMBER() OVER(
           PARTITION BY t.nrm_customer_number
               ORDER BY t.score DESC
                      , t.nrm_customer_name
                      , t.nrm_business_unit
                      , t.nrm_country_code
       )                                                                          rnk_score
     , COUNT()      OVER(PARTITION BY t.nrm_customer_number)                       cnt_customer_number
     , a.nrm_array_customer_address
     , a.nrm_agg_customer_name
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
  FROM cte_customer_score              t
       INNER JOIN
       cte_agg                         a
          ON a.nrm_customer_number     = t.nrm_customer_number
;
