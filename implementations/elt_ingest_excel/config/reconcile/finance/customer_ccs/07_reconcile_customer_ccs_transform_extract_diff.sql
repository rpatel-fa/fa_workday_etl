DROP TABLE IF EXISTS reconcile_customer_ccs_transform_extract_tax_diff_detail
;
CREATE TABLE reconcile_customer_ccs_transform_extract_tax_diff_detail
    AS
  WITH cte_expected
    AS (
SELECT customer_id
     , customer_name
     , tax_code
     , tax_id_country
     , tax_id
     , tax_id_type
     , primary_tax_id
     , transaction_tax_id
     , tax_status_country
     , transaction_tax_status
     , withholding_tax_status
  FROM workday_customer_tax
       )
     , cte_actual
    AS (
SELECT customer_id
     , customer                                                                  actual_customer_name
     , default_tax_code                                                          actual_tax_code
     , country_for_tax_id                                                        actual_tax_id_country
     , tax_id                                                                    actual_tax_id
     , company_tax_id_type                                                       actual_tax_id_type
     , is_primary_tax_id                                                         actual_primary_tax_id
     , is_transaction_tax_id                                                     actual_transaction_tax_id
     , country_for_tax_status                                                    actual_tax_status_country
     , transaction_tax_status                                                    actual_transaction_tax_status
     , withholding_tax_status                                                    actual_withholding_tax_status
  FROM src_fin_workday_extract_ccs_customer
       )
SELECT COALESCE(e.customer_id, a.customer_id)                                    customer_id
     , CASE
         WHEN e.customer_id IS NULL THEN 'ONLY_IN_EXTRACT'
         WHEN a.customer_id IS NULL THEN 'ONLY_IN_TRANSFORM'
         ELSE 'MATCHED_CUSTOMER_ID'
       END                                                                       comparison_scope
     , e.customer_name                                                           expected_customer_name
     , a.actual_customer_name
     , e.tax_code                                                                expected_tax_code
     , a.actual_tax_code
     , e.tax_id_country                                                          expected_tax_id_country
     , a.actual_tax_id_country
     , e.tax_id                                                                  expected_tax_id
     , a.actual_tax_id
     , e.tax_id_type                                                             expected_tax_id_type
     , a.actual_tax_id_type
     , e.primary_tax_id                                                          expected_primary_tax_id
     , a.actual_primary_tax_id
     , e.transaction_tax_id                                                      expected_transaction_tax_id
     , a.actual_transaction_tax_id
     , e.tax_status_country                                                      expected_tax_status_country
     , a.actual_tax_status_country
     , e.transaction_tax_status                                                  expected_transaction_tax_status
     , a.actual_transaction_tax_status
     , e.withholding_tax_status                                                  expected_withholding_tax_status
     , a.actual_withholding_tax_status
     , CASE WHEN COALESCE(e.tax_code, '') <> COALESCE(a.actual_tax_code, '') THEN TRUE ELSE FALSE END
                                                                                has_tax_code_diff
     , CASE WHEN COALESCE(e.tax_id_country, '') <> COALESCE(a.actual_tax_id_country, '') THEN TRUE ELSE FALSE END
                                                                                has_tax_id_country_diff
     , CASE WHEN COALESCE(e.tax_id, '') <> COALESCE(a.actual_tax_id, '') THEN TRUE ELSE FALSE END
                                                                                has_tax_id_diff
     , CASE WHEN COALESCE(e.tax_id_type, '') <> COALESCE(a.actual_tax_id_type, '') THEN TRUE ELSE FALSE END
                                                                                has_tax_id_type_diff
     , CASE WHEN COALESCE(e.primary_tax_id, '') <> COALESCE(a.actual_primary_tax_id, '') THEN TRUE ELSE FALSE END
                                                                                has_primary_tax_id_diff
     , CASE WHEN COALESCE(e.transaction_tax_id, '') <> COALESCE(a.actual_transaction_tax_id, '') THEN TRUE ELSE FALSE END
                                                                                has_transaction_tax_id_diff
     , CASE WHEN COALESCE(e.tax_status_country, '') <> COALESCE(a.actual_tax_status_country, '') THEN TRUE ELSE FALSE END
                                                                                has_tax_status_country_diff
     , CASE WHEN COALESCE(e.transaction_tax_status, '') <> COALESCE(a.actual_transaction_tax_status, '') THEN TRUE ELSE FALSE END
                                                                                has_transaction_tax_status_diff
     , CASE WHEN COALESCE(e.withholding_tax_status, '') <> COALESCE(a.actual_withholding_tax_status, '') THEN TRUE ELSE FALSE END
                                                                                has_withholding_tax_status_diff
  FROM cte_expected                                                              e
       FULL OUTER JOIN
       cte_actual                                                                a
          ON a.customer_id                                                       = e.customer_id
;

DROP TABLE IF EXISTS reconcile_customer_ccs_transform_extract_tax_diff_summary
;
CREATE TABLE reconcile_customer_ccs_transform_extract_tax_diff_summary
    AS
SELECT 'matched_customer_ids'                                                    metric
     , COUNT(*)                                                                  row_count
     , 'Direct compare on shared customer_id values'                             status
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'

 UNION ALL
SELECT 'only_in_extract'
     , COUNT(*)
     , CASE
         WHEN COUNT(*) = 1
          AND MAX(customer_id) = '100001'
         THEN 'EXPECTED - dummy Workday row'
         WHEN COUNT(*) = 0
         THEN 'CONFIRMED'
         ELSE 'REVIEW'
       END
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'ONLY_IN_EXTRACT'

 UNION ALL
SELECT 'only_in_transform'
     , COUNT(*)
     , CASE WHEN COUNT(*) = 0 THEN 'CONFIRMED' ELSE 'REVIEW' END
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'ONLY_IN_TRANSFORM'

 UNION ALL
SELECT 'tax_code_diff_on_matched_customer_ids'
     , COUNT(*)
     , 'Confirmed sent-vs-extracted tax code difference'
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_tax_code_diff

 UNION ALL
SELECT 'tax_id_country_diff_on_matched_customer_ids'
     , COUNT(*)
     , 'Confirmed sent-vs-extracted tax-id country difference'
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_tax_id_country_diff

 UNION ALL
SELECT 'tax_id_diff_on_matched_customer_ids'
     , COUNT(*)
     , 'Confirmed sent-vs-extracted tax-id difference'
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_tax_id_diff

 UNION ALL
SELECT 'tax_id_type_diff_on_matched_customer_ids'
     , COUNT(*)
     , 'Confirmed sent-vs-extracted tax-id-type difference'
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_tax_id_type_diff

 UNION ALL
SELECT 'primary_tax_id_diff_on_matched_customer_ids'
     , COUNT(*)
     , 'Confirmed sent-vs-extracted primary-tax-id flag difference'
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_primary_tax_id_diff

 UNION ALL
SELECT 'transaction_tax_id_diff_on_matched_customer_ids'
     , COUNT(*)
     , 'Confirmed sent-vs-extracted transaction-tax-id flag difference'
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_transaction_tax_id_diff

 UNION ALL
SELECT 'tax_status_country_diff_on_matched_customer_ids'
     , COUNT(*)
     , 'Confirmed sent-vs-extracted tax-status-country difference'
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_tax_status_country_diff

 UNION ALL
SELECT 'transaction_tax_status_diff_on_matched_customer_ids'
     , COUNT(*)
     , 'Confirmed sent-vs-extracted transaction-tax-status difference'
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_transaction_tax_status_diff

 UNION ALL
SELECT 'withholding_tax_status_diff_on_matched_customer_ids'
     , COUNT(*)
     , 'Confirmed sent-vs-extracted withholding-tax-status difference'
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_withholding_tax_status_diff
;

DROP TABLE IF EXISTS reconcile_customer_ccs_transform_extract_tax_diff_pattern_summary
;
CREATE TABLE reconcile_customer_ccs_transform_extract_tax_diff_pattern_summary
    AS
SELECT 'tax_code'                                                                 attribute_name
     , expected_tax_code                                                          expected_value
     , actual_tax_code                                                            actual_value
     , COUNT(*)                                                                   row_count
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_tax_code_diff
 GROUP BY
       expected_tax_code
     , actual_tax_code

 UNION ALL
SELECT 'tax_id_country'
     , expected_tax_id_country
     , actual_tax_id_country
     , COUNT(*)
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_tax_id_country_diff
 GROUP BY
       expected_tax_id_country
     , actual_tax_id_country

 UNION ALL
SELECT 'tax_id'
     , expected_tax_id
     , actual_tax_id
     , COUNT(*)
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_tax_id_diff
 GROUP BY
       expected_tax_id
     , actual_tax_id

 UNION ALL
SELECT 'tax_id_type'
     , expected_tax_id_type
     , actual_tax_id_type
     , COUNT(*)
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_tax_id_type_diff
 GROUP BY
       expected_tax_id_type
     , actual_tax_id_type

 UNION ALL
SELECT 'primary_tax_id'
     , expected_primary_tax_id
     , actual_primary_tax_id
     , COUNT(*)
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_primary_tax_id_diff
 GROUP BY
       expected_primary_tax_id
     , actual_primary_tax_id

 UNION ALL
SELECT 'transaction_tax_id'
     , expected_transaction_tax_id
     , actual_transaction_tax_id
     , COUNT(*)
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_transaction_tax_id_diff
 GROUP BY
       expected_transaction_tax_id
     , actual_transaction_tax_id

 UNION ALL
SELECT 'tax_status_country'
     , expected_tax_status_country
     , actual_tax_status_country
     , COUNT(*)
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_tax_status_country_diff
 GROUP BY
       expected_tax_status_country
     , actual_tax_status_country

 UNION ALL
SELECT 'transaction_tax_status'
     , expected_transaction_tax_status
     , actual_transaction_tax_status
     , COUNT(*)
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_transaction_tax_status_diff
 GROUP BY
       expected_transaction_tax_status
     , actual_transaction_tax_status

 UNION ALL
SELECT 'withholding_tax_status'
     , expected_withholding_tax_status
     , actual_withholding_tax_status
     , COUNT(*)
  FROM reconcile_customer_ccs_transform_extract_tax_diff_detail
 WHERE comparison_scope                                                          = 'MATCHED_CUSTOMER_ID'
   AND has_withholding_tax_status_diff
 GROUP BY
       expected_withholding_tax_status
     , actual_withholding_tax_status

 ORDER BY
       attribute_name
     , row_count DESC
     , expected_value
     , actual_value
;
