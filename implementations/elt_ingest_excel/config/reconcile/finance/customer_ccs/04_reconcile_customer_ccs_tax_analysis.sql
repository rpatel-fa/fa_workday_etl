DROP TABLE IF EXISTS reconcile_customer_ccs_tax_analysis
;
CREATE TABLE reconcile_customer_ccs_tax_analysis
    AS
SELECT t.customer_number
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , t.attribute_name
     , t.expected_value
     , t.actual_value
     , CASE
         WHEN t.is_match = 'Y'
         THEN 'EXACT_MATCH'

         WHEN t.attribute_name IN ('tax_id', 'tax_id_country', 'tax_id_type')
          AND COALESCE(TRIM(x.expected_tax_id_type), '') = 'Company Number'
          AND NULLIF(TRIM(t.expected_value), '')         IS NOT NULL
          AND NULLIF(TRIM(t.actual_value), '')           IS NULL
         THEN 'EXTRACT_OMITS_COMPANY_NUMBER'

         WHEN t.attribute_name IN ('tax_id', 'tax_id_country', 'tax_id_type')
          AND NULLIF(TRIM(t.expected_value), '') IS NOT NULL
          AND NULLIF(TRIM(t.actual_value), '')   IS NULL
         THEN 'MISSING_IN_WORKDAY'

         WHEN t.attribute_name IN ('tax_id', 'tax_id_country', 'tax_id_type')
          AND NULLIF(TRIM(t.expected_value), '') IS NULL
          AND NULLIF(TRIM(t.actual_value), '')   IS NOT NULL
         THEN 'EXTRA_IN_WORKDAY'

         WHEN t.attribute_name = 'tax_status_country'
          AND COALESCE(TRIM(t.expected_value), '') = 'Isle of Man'
          AND COALESCE(TRIM(t.actual_value), '')   = 'United Kingdom'
         THEN 'COUNTRY_NORMALISATION_REVIEW'

         WHEN t.attribute_name = 'transaction_tax_status'
          AND COALESCE(TRIM(t.actual_value), '')   = 'Company is not VAT/GST Registered in this Country'
          AND COALESCE(TRIM(t.expected_value), '') IN (
                'Non EU Company Not VAT registered in this country'
              , 'EU Company - Not Registered in this country'
              )
         THEN 'SEMANTIC_EQUIVALENT'

         ELSE 'VALUE_MISMATCH'
       END                                                                         comparison_class
     , CASE
         WHEN t.is_match = 'Y'
         THEN 'Exact string match'

         WHEN t.attribute_name IN ('tax_id', 'tax_id_country', 'tax_id_type')
          AND COALESCE(TRIM(x.expected_tax_id_type), '') = 'Company Number'
          AND NULLIF(TRIM(t.expected_value), '')         IS NOT NULL
          AND NULLIF(TRIM(t.actual_value), '')           IS NULL
         THEN 'Loaded CCS extract does not surface Company Number tax IDs; treat as an extract limitation unless a counterexample appears'

         WHEN t.attribute_name IN ('tax_id', 'tax_id_country', 'tax_id_type')
          AND NULLIF(TRIM(t.expected_value), '') IS NOT NULL
          AND NULLIF(TRIM(t.actual_value), '')   IS NULL
         THEN 'Expected transformed tax value is populated but the loaded Workday extract is blank'

         WHEN t.attribute_name IN ('tax_id', 'tax_id_country', 'tax_id_type')
          AND NULLIF(TRIM(t.expected_value), '') IS NULL
          AND NULLIF(TRIM(t.actual_value), '')   IS NOT NULL
         THEN 'Loaded Workday extract contains a tax value that is blank in the transformed target'

         WHEN t.attribute_name = 'tax_status_country'
          AND COALESCE(TRIM(t.expected_value), '') = 'Isle of Man'
          AND COALESCE(TRIM(t.actual_value), '')   = 'United Kingdom'
         THEN 'Potential country normalisation difference rather than a true target mismatch'

         WHEN t.attribute_name = 'transaction_tax_status'
          AND COALESCE(TRIM(t.actual_value), '')   = 'Company is not VAT/GST Registered in this Country'
          AND COALESCE(TRIM(t.expected_value), '') IN (
                'Non EU Company Not VAT registered in this country'
              , 'EU Company - Not Registered in this country'
              )
         THEN 'Likely semantic-equivalent wording difference for non-registered tax status'

         ELSE 'Review required - values differ materially or the equivalence rule is not yet defined'
       END                                                                         comparison_reason
  FROM reconcile_customer_ccs_attribute_detail                                     t
       LEFT OUTER JOIN
       reconcile_customer_ccs_expected_customer                                    x
          ON x.expected_customer_id                                                = t.expected_customer_id
 WHERE t.attribute_name IN (
         'tax_id'
       , 'tax_id_country'
       , 'tax_id_type'
       , 'tax_status_country'
       , 'transaction_tax_status'
       )
;

DROP TABLE IF EXISTS reconcile_customer_ccs_tax_analysis_summary
;
CREATE TABLE reconcile_customer_ccs_tax_analysis_summary
    AS
SELECT t.attribute_name
     , t.comparison_class
     , COUNT(*)                                                                    row_count
  FROM reconcile_customer_ccs_tax_analysis                                         t
 GROUP BY
       t.attribute_name
     , t.comparison_class
 ORDER BY
       t.attribute_name
     , t.comparison_class
;

INSERT INTO reconciliation_customer_ccs_issues
SELECT '3_attribute_reconciliation'                                                section
     , 'TAX_REVIEW'                                                                issue_type
     , t.customer_number
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , t.attribute_name
     , t.expected_value
     , t.actual_value
     , t.comparison_reason                                                         details
  FROM reconcile_customer_ccs_tax_analysis                                         t
 WHERE t.comparison_class IN (
         'MISSING_IN_WORKDAY'
       , 'EXTRA_IN_WORKDAY'
       , 'VALUE_MISMATCH'
       , 'COUNTRY_NORMALISATION_REVIEW'
       )
;
