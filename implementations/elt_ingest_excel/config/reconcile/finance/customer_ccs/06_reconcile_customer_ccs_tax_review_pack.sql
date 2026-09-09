DROP TABLE IF EXISTS reconcile_customer_ccs_tax_review_case_summary
;
CREATE TABLE reconcile_customer_ccs_tax_review_case_summary
    AS
SELECT t.expected_customer_id
     , c.expected_customer_name
     , c.expected_reference_id
     , t.attribute_name
     , t.comparison_class
     , t.expected_value
     , t.actual_value
     , COUNT(*)                                                                    source_customer_number_count
     , STRING_AGG(DISTINCT t.customer_number, ';' ORDER BY t.customer_number)      source_customer_numbers
     , c.source_business_units
     , c.source_countries
     , c.source_country_codes
     , c.source_tax_schedule_ids
     , c.source_tax_registration_numbers
     , CASE
         WHEN t.attribute_name = 'tax_status_country'
          AND t.comparison_class = 'COUNTRY_NORMALISATION_REVIEW'
         THEN 'Accepted Workday handling: Isle of Man is reported as United Kingdom'

         WHEN t.attribute_name = 'tax_status_country'
          AND t.expected_value = 'United Kingdom'
          AND t.actual_value   = 'Singapore'
          AND POSITION('Singapore'      IN COALESCE(c.source_countries, '')) > 0
          AND POSITION('United Kingdom' IN COALESCE(c.source_countries, '')) > 0
         THEN 'Workday tax country appears to follow the Singapore primary address; if the customer should remain UK for tax, treat this as a Workday target-state issue rather than a transform defect'

         WHEN t.attribute_name = 'transaction_tax_status'
          AND t.expected_value = 'EU Company - Exempt in THIS country'
          AND t.actual_value   = 'Customer of Goods'
          AND c.expected_country_code = 'IE'
          AND c.expected_tax_schedule_id = 'SZ'
         THEN 'Potential Ireland/SZ reporting rule gap or Workday wording variation'

         ELSE 'Review manually'
       END                                                                         suggested_interpretation
     , CASE
         WHEN t.attribute_name = 'tax_status_country'
          AND t.comparison_class = 'COUNTRY_NORMALISATION_REVIEW'
         THEN 'accepted_workday_normalisation'

         WHEN t.attribute_name = 'tax_status_country'
          AND t.expected_value = 'United Kingdom'
          AND t.actual_value   = 'Singapore'
         THEN 'workday_target_state_review'

         WHEN t.attribute_name = 'transaction_tax_status'
          AND t.expected_value = 'EU Company - Exempt in THIS country'
          AND t.actual_value   = 'Customer of Goods'
          AND c.expected_country_code = 'IE'
          AND c.expected_tax_schedule_id = 'SZ'
         THEN 'ireland_sz_rule_gap'

         ELSE 'manual_review'
       END                                                                         suggested_bucket
  FROM reconcile_customer_ccs_tax_analysis                                         t
       INNER JOIN
       reconcile_customer_ccs_tax_review_customer                                  c
          ON c.expected_customer_id                                                = t.expected_customer_id
 WHERE t.comparison_class IN (
         'VALUE_MISMATCH'
       , 'COUNTRY_NORMALISATION_REVIEW'
       )
 GROUP BY
       t.expected_customer_id
     , c.expected_customer_name
     , c.expected_reference_id
     , t.attribute_name
     , t.comparison_class
     , t.expected_value
     , t.actual_value
     , c.source_business_units
     , c.source_countries
     , c.source_country_codes
     , c.source_tax_schedule_ids
     , c.source_tax_registration_numbers
     , c.expected_country_code
     , c.expected_tax_schedule_id
 ORDER BY
       t.attribute_name
     , t.expected_customer_id
;

DROP TABLE IF EXISTS reconcile_customer_ccs_tax_review_operator_pack
;
CREATE TABLE reconcile_customer_ccs_tax_review_operator_pack
    AS
SELECT ROW_NUMBER() OVER(
           ORDER BY CASE s.suggested_bucket
                      WHEN 'workday_target_state_review'  THEN 1
                      WHEN 'ireland_sz_rule_gap'          THEN 2
                      WHEN 'accepted_workday_normalisation'
                                                           THEN 3
                      ELSE 4
                    END
                  , s.attribute_name
                  , s.expected_customer_id
       )                                                                            review_order
     , CASE s.suggested_bucket
         WHEN 'workday_target_state_review'  THEN 'High'
         WHEN 'ireland_sz_rule_gap'          THEN 'Medium'
         WHEN 'accepted_workday_normalisation'
                                              THEN 'Low'
         ELSE 'Medium'
       END                                                                         review_priority
     , s.suggested_bucket
     , s.suggested_interpretation
     , s.expected_customer_id
     , s.expected_customer_name
     , s.expected_reference_id
     , s.attribute_name
     , s.comparison_class
     , s.expected_value
     , s.actual_value
     , s.source_customer_number_count
     , s.source_customer_numbers
     , s.source_business_units
     , s.source_countries
     , s.source_country_codes
     , s.source_tax_schedule_ids
     , s.source_tax_registration_numbers
  FROM reconcile_customer_ccs_tax_review_case_summary                               s
 ORDER BY
       review_order
;
