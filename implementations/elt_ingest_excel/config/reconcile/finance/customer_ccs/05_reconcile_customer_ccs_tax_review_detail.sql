DROP TABLE IF EXISTS reconcile_customer_ccs_tax_review_customer
;
CREATE TABLE reconcile_customer_ccs_tax_review_customer
    AS
  WITH cte_review_customer
    AS (
SELECT DISTINCT
       t.expected_customer_id
  FROM reconcile_customer_ccs_tax_analysis                                         t
 WHERE t.comparison_class IN (
         'VALUE_MISMATCH'
       , 'COUNTRY_NORMALISATION_REVIEW'
       )
       )
     , cte_source_rows
    AS (
SELECT c.customer_id
     , TRIM(x.customer_number)                                                     source_customer_number
     , b.business_unit
     , b.customer_name                                                             source_customer_name
     , b.country                                                                   source_country
     , TRIM(b.country_code)                                                        source_country_code
     , b.tax_registration_number                                                   source_tax_registration_number
     , b.tax_schedule_id                                                           source_tax_schedule_id
  FROM cte_review_customer                                                         r
       INNER JOIN
       src_fin_customer                                                            c
          ON c.customer_id                                                         = r.expected_customer_id
     , UNNEST(STRING_SPLIT(c.nrm_agg_customer_number, '|'))                        x(customer_number)
       LEFT OUTER JOIN
       src_fin_customer_base                                                       b
          ON TRIM(b.customer_number)                                               = TRIM(x.customer_number)
       )
SELECT c.customer_id                                                               expected_customer_id
     , c.nrm_customer_name                                                         expected_customer_name
     , c.nrm_agg_customer_number                                                   expected_reference_id
     , c.nrm_country_code                                                          expected_country_code
     , c.nrm_country_name                                                          expected_country_name
     , c.nrm_tax_schedule_id                                                       expected_tax_schedule_id
     , c.nrm_tax_registration_number                                               expected_tax_registration_number
     , wt.tax_status_country                                                       expected_tax_status_country
     , wt.transaction_tax_status                                                   expected_transaction_tax_status
     , w.customer_id                                                               workday_customer_id
     , w.country_for_tax_status                                                    actual_tax_status_country
     , w.transaction_tax_status                                                    actual_transaction_tax_status
     , COUNT(DISTINCT s.source_customer_number)                                    source_customer_number_count
     , STRING_AGG(DISTINCT s.source_customer_number, ';' ORDER BY s.source_customer_number)
                                                                                    source_customer_numbers
     , STRING_AGG(DISTINCT s.business_unit, ';' ORDER BY s.business_unit)          source_business_units
     , STRING_AGG(DISTINCT s.source_country, ';' ORDER BY s.source_country)        source_countries
     , STRING_AGG(DISTINCT s.source_country_code, ';' ORDER BY s.source_country_code)
                                                                                    source_country_codes
     , STRING_AGG(DISTINCT COALESCE(NULLIF(TRIM(s.source_tax_schedule_id), ''), '<NULL>'),
                  ';' ORDER BY COALESCE(NULLIF(TRIM(s.source_tax_schedule_id), ''), '<NULL>'))
                                                                                    source_tax_schedule_ids
     , STRING_AGG(DISTINCT COALESCE(NULLIF(TRIM(s.source_tax_registration_number), ''), '<NULL>'),
                  ';' ORDER BY COALESCE(NULLIF(TRIM(s.source_tax_registration_number), ''), '<NULL>'))
                                                                                    source_tax_registration_numbers
  FROM cte_review_customer                                                         r
       INNER JOIN
       src_fin_customer                                                            c
          ON c.customer_id                                                         = r.expected_customer_id
       LEFT OUTER JOIN
       workday_customer_tax                                                        wt
          ON wt.customer_id                                                        = c.customer_id
       LEFT OUTER JOIN
       src_fin_workday_extract_ccs_customer                                        w
          ON w.customer_id                                                         = c.customer_id
       LEFT OUTER JOIN
       cte_source_rows                                                             s
          ON s.customer_id                                                         = c.customer_id
 GROUP BY
       c.customer_id
     , c.nrm_customer_name
     , c.nrm_agg_customer_number
     , c.nrm_country_code
     , c.nrm_country_name
     , c.nrm_tax_schedule_id
     , c.nrm_tax_registration_number
     , wt.tax_status_country
     , wt.transaction_tax_status
     , w.customer_id
     , w.country_for_tax_status
     , w.transaction_tax_status
;

DROP TABLE IF EXISTS reconcile_customer_ccs_tax_review_source_rows
;
CREATE TABLE reconcile_customer_ccs_tax_review_source_rows
    AS
SELECT t.attribute_name
     , t.comparison_class
     , t.expected_customer_id
     , c.nrm_customer_name                                                         expected_customer_name
     , c.nrm_agg_customer_number                                                   expected_reference_id
     , s.business_unit
     , TRIM(s.customer_number)                                                     source_customer_number
     , s.customer_name                                                             source_customer_name
     , s.country                                                                   source_country
     , TRIM(s.country_code)                                                        source_country_code
     , s.tax_registration_number                                                   source_tax_registration_number
     , s.tax_schedule_id                                                           source_tax_schedule_id
     , t.expected_value
     , t.actual_value
     , t.comparison_reason
  FROM reconcile_customer_ccs_tax_analysis                                         t
       INNER JOIN
       src_fin_customer                                                            c
          ON c.customer_id                                                         = t.expected_customer_id
     , UNNEST(STRING_SPLIT(c.nrm_agg_customer_number, '|'))                        x(customer_number)
       LEFT OUTER JOIN
       src_fin_customer_base                                                       s
          ON TRIM(s.customer_number)                                               = TRIM(x.customer_number)
 WHERE t.comparison_class IN (
         'VALUE_MISMATCH'
       , 'COUNTRY_NORMALISATION_REVIEW'
       )
;

DROP TABLE IF EXISTS reconcile_customer_ccs_tax_review_pattern_summary
;
CREATE TABLE reconcile_customer_ccs_tax_review_pattern_summary
    AS
SELECT t.attribute_name
     , t.comparison_class
     , t.expected_value
     , t.actual_value
     , COUNT(*)                                                                    source_customer_number_count
     , COUNT(DISTINCT t.expected_customer_id)                                      merged_customer_count
     , STRING_AGG(DISTINCT t.expected_customer_id, ';' ORDER BY t.expected_customer_id)
                                                                                    merged_customer_ids
  FROM reconcile_customer_ccs_tax_analysis                                         t
 WHERE t.comparison_class IN (
         'VALUE_MISMATCH'
       , 'COUNTRY_NORMALISATION_REVIEW'
       )
 GROUP BY
       t.attribute_name
     , t.comparison_class
     , t.expected_value
     , t.actual_value
 ORDER BY
       t.attribute_name
     , t.comparison_class
     , t.expected_value
     , t.actual_value
;
