DROP TABLE IF EXISTS reconcile_customer_ccs_attribute_detail
;
CREATE TABLE reconcile_customer_ccs_attribute_detail
    AS
  WITH cte_bridge
    AS (
SELECT DISTINCT
       b.customer_number
     , b.source_row_count
     , b.source_business_units
     , b.source_customer_names
     , b.source_countries
     , b.expected_customer_id
     , b.expected_reference_id
     , b.workday_customer_id
     , b.workday_reference_id
     , e.expected_customer_name
     , e.expected_customer_category
     , e.expected_payment_terms
     , e.expected_default_payment_type
     , e.expected_default_currency
     , e.expected_tax_id_country
     , e.expected_tax_id_type
     , e.expected_tax_id
     , e.expected_tax_status_country
     , e.expected_transaction_tax_status
     , w.customer                                                                  actual_customer_name
     , w.customer_category                                                         actual_customer_category
     , w.default_payment_terms                                                     actual_payment_terms
     , w.default_payment_type                                                      actual_default_payment_type
     , w.default_currency                                                          actual_default_currency
     , w.country_for_tax_id                                                        actual_tax_id_country
     , w.company_tax_id_type                                                       actual_tax_id_type
     , w.tax_id                                                                    actual_tax_id
     , w.country_for_tax_status                                                    actual_tax_status_country
     , w.transaction_tax_status                                                    actual_transaction_tax_status
  FROM reconcile_customer_ccs_bridge                                               b
       INNER JOIN
       reconcile_customer_ccs_expected_customer                                    e
          ON e.expected_customer_id                                                = b.expected_customer_id
       INNER JOIN
       src_fin_workday_extract_ccs_customer                                        w
          ON w.customer_id                                                         = b.workday_customer_id
 WHERE b.in_source_gp                                                              = 'Y'
   AND b.in_merged_customer                                                        = 'Y'
   AND b.in_workday                                                                = 'Y'
       )
SELECT t.customer_number
     , t.source_row_count
     , t.source_business_units
     , t.source_customer_names
     , t.source_countries
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , 'customer_name'                                                             attribute_name
     , t.expected_customer_name                                                    expected_value
     , t.actual_customer_name                                                      actual_value
     , CASE
         WHEN COALESCE(TRIM(t.expected_customer_name), '') = COALESCE(TRIM(t.actual_customer_name), '')
         THEN 'Y' ELSE 'N'
       END                                                                         is_match
  FROM cte_bridge                                                                  t

 UNION ALL
SELECT t.customer_number
     , t.source_row_count
     , t.source_business_units
     , t.source_customer_names
     , t.source_countries
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , 'reference_id'
     , t.expected_reference_id
     , t.workday_reference_id
     , CASE
         WHEN COALESCE(TRIM(t.expected_reference_id), '') = COALESCE(TRIM(t.workday_reference_id), '')
         THEN 'Y' ELSE 'N'
       END
  FROM cte_bridge                                                                  t

 UNION ALL
SELECT t.customer_number
     , t.source_row_count
     , t.source_business_units
     , t.source_customer_names
     , t.source_countries
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , 'customer_category'
     , t.expected_customer_category
     , t.actual_customer_category
     , CASE
         WHEN COALESCE(TRIM(t.expected_customer_category), '') = COALESCE(TRIM(t.actual_customer_category), '')
         THEN 'Y' ELSE 'N'
       END
  FROM cte_bridge                                                                  t

 UNION ALL
SELECT t.customer_number
     , t.source_row_count
     , t.source_business_units
     , t.source_customer_names
     , t.source_countries
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , 'default_payment_terms'
     , t.expected_payment_terms
     , t.actual_payment_terms
     , CASE
         WHEN COALESCE(TRIM(t.expected_payment_terms), '') = COALESCE(TRIM(t.actual_payment_terms), '')
         THEN 'Y' ELSE 'N'
       END
  FROM cte_bridge                                                                  t

 UNION ALL
SELECT t.customer_number
     , t.source_row_count
     , t.source_business_units
     , t.source_customer_names
     , t.source_countries
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , 'default_payment_type'
     , t.expected_default_payment_type
     , t.actual_default_payment_type
     , CASE
         WHEN COALESCE(TRIM(t.expected_default_payment_type), '') = COALESCE(TRIM(t.actual_default_payment_type), '')
         THEN 'Y' ELSE 'N'
       END
  FROM cte_bridge                                                                  t

 UNION ALL
SELECT t.customer_number
     , t.source_row_count
     , t.source_business_units
     , t.source_customer_names
     , t.source_countries
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , 'default_currency'
     , t.expected_default_currency
     , t.actual_default_currency
     , CASE
         WHEN COALESCE(TRIM(t.expected_default_currency), '') = COALESCE(TRIM(t.actual_default_currency), '')
         THEN 'Y' ELSE 'N'
       END
  FROM cte_bridge                                                                  t

 UNION ALL
SELECT t.customer_number
     , t.source_row_count
     , t.source_business_units
     , t.source_customer_names
     , t.source_countries
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , 'tax_id_country'
     , t.expected_tax_id_country
     , t.actual_tax_id_country
     , CASE
         WHEN COALESCE(TRIM(t.expected_tax_id_country), '') = COALESCE(TRIM(t.actual_tax_id_country), '')
         THEN 'Y' ELSE 'N'
       END
  FROM cte_bridge                                                                  t

 UNION ALL
SELECT t.customer_number
     , t.source_row_count
     , t.source_business_units
     , t.source_customer_names
     , t.source_countries
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , 'tax_id_type'
     , t.expected_tax_id_type
     , t.actual_tax_id_type
     , CASE
         WHEN COALESCE(TRIM(t.expected_tax_id_type), '') = COALESCE(TRIM(t.actual_tax_id_type), '')
         THEN 'Y' ELSE 'N'
       END
  FROM cte_bridge                                                                  t

 UNION ALL
SELECT t.customer_number
     , t.source_row_count
     , t.source_business_units
     , t.source_customer_names
     , t.source_countries
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , 'tax_id'
     , t.expected_tax_id
     , t.actual_tax_id
     , CASE
         WHEN COALESCE(TRIM(t.expected_tax_id), '') = COALESCE(TRIM(t.actual_tax_id), '')
         THEN 'Y' ELSE 'N'
       END
  FROM cte_bridge                                                                  t

 UNION ALL
SELECT t.customer_number
     , t.source_row_count
     , t.source_business_units
     , t.source_customer_names
     , t.source_countries
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , 'tax_status_country'
     , t.expected_tax_status_country
     , t.actual_tax_status_country
     , CASE
         WHEN COALESCE(TRIM(t.expected_tax_status_country), '') = COALESCE(TRIM(t.actual_tax_status_country), '')
         THEN 'Y' ELSE 'N'
       END
  FROM cte_bridge                                                                  t

 UNION ALL
SELECT t.customer_number
     , t.source_row_count
     , t.source_business_units
     , t.source_customer_names
     , t.source_countries
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , 'transaction_tax_status'
     , t.expected_transaction_tax_status
     , t.actual_transaction_tax_status
     , CASE
         WHEN COALESCE(TRIM(t.expected_transaction_tax_status), '') = COALESCE(TRIM(t.actual_transaction_tax_status), '')
         THEN 'Y' ELSE 'N'
       END
  FROM cte_bridge                                                                  t
;

DROP TABLE IF EXISTS reconcile_customer_ccs_attribute_summary
;
CREATE TABLE reconcile_customer_ccs_attribute_summary
    AS
SELECT t.attribute_name
     , COUNT(*)                                                                    source_customer_number_count
     , SUM(CASE WHEN t.is_match = 'Y' THEN 1 ELSE 0 END)                           matched_count
     , SUM(CASE WHEN t.is_match = 'N' THEN 1 ELSE 0 END)                           mismatched_count
  FROM reconcile_customer_ccs_attribute_detail                                     t
 GROUP BY
       t.attribute_name
 ORDER BY
       t.attribute_name
;

INSERT INTO reconciliation_customer_ccs_issues
SELECT '3_attribute_reconciliation'                                                section
     , 'ATTRIBUTE_MISMATCH'                                                        issue_type
     , t.customer_number
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , t.attribute_name
     , t.expected_value
     , t.actual_value
     , 'Exact string mismatch between expected transformed value and loaded Workday extract value'
                                                                                    details
  FROM reconcile_customer_ccs_attribute_detail                                     t
 WHERE t.is_match                                                                  = 'N'
   AND t.attribute_name NOT IN (
         'tax_id'
       , 'tax_id_country'
       , 'tax_id_type'
       , 'tax_status_country'
       , 'transaction_tax_status'
       )
;
