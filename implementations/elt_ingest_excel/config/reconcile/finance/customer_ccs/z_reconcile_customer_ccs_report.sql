DROP TABLE IF EXISTS reconcile_customer_ccs_report
;
CREATE TABLE reconcile_customer_ccs_report
    AS
SELECT '0_reconciliation_contract'                                                 section
     , '0a_primary_lens'                                                           metric
     , 'sent_workday_payload_vs_loaded_extract'                                    metric_value
     , 'Primary reconciliation lens - prove whether the loaded extract preserves what was sent to Workday'
                                                                                    status

 UNION ALL
SELECT '0_reconciliation_contract'
     , '0b_secondary_lens'
     , 'gp_source_vs_loaded_extract'                                               metric_value
     , 'Secondary diagnostic lens - explain safe cleansing and other upstream transformations before load'

 UNION ALL
SELECT '0_reconciliation_contract'
     , '0c_gold_definition'
     , 'workday_* transform output tables'                                         metric_value
     , 'Use sent Workday transform output as the gold expected value'

 UNION ALL
SELECT '1_population_funnel'                                                       section
     , '1a_src_fin_customer_base_row_count'                                        metric
     , COUNT(*)::VARCHAR                                                            metric_value
     , CAST(NULL AS VARCHAR)                                                        status
  FROM src_fin_customer_base

 UNION ALL
SELECT '1_population_funnel'
     , '1b_distinct_gp_customer_number'
     , COUNT(DISTINCT TRIM(customer_number))::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM src_fin_customer_base
 WHERE NULLIF(TRIM(customer_number), '')                                            IS NOT NULL

 UNION ALL
SELECT '1_population_funnel'
     , '1c_src_fin_customer_row_count'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM src_fin_customer

 UNION ALL
SELECT '1_population_funnel'
     , '1d_exploded_src_fin_customer_number'
     , COUNT(DISTINCT TRIM(customer_number))::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconcile_customer_ccs_source_merge_bridge

 UNION ALL
SELECT '1_population_funnel'
     , '1e_src_fin_workday_extract_ccs_customer_row_count'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM src_fin_workday_extract_ccs_customer

 UNION ALL
SELECT '1_population_funnel'
     , '1f_exploded_workday_reference_id'
     , COUNT(DISTINCT TRIM(customer_number))::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconcile_customer_ccs_workday_reference_bridge

 UNION ALL
SELECT '2_population_exceptions'
     , '2a_source_customer_numbers_missing_in_merged_customer'
     , COUNT(*)::VARCHAR
     , CASE WHEN COUNT(*) = 0 THEN 'CONFIRMED' ELSE 'REVIEW' END
  FROM (
        SELECT s.customer_number
          FROM reconcile_customer_ccs_source_profile     s
               LEFT OUTER JOIN
               reconcile_customer_ccs_source_merge_bridge m
                  ON m.customer_number                   = s.customer_number
         WHERE m.customer_number                         IS NULL
       )                                                t

 UNION ALL
SELECT '2_population_exceptions'
     , '2b_source_customer_numbers_missing_in_workday'
     , COUNT(*)::VARCHAR
     , CASE WHEN COUNT(*) = 0 THEN 'CONFIRMED' ELSE 'REVIEW' END
  FROM reconcile_customer_ccs_missing_in_workday

 UNION ALL
SELECT '2_population_exceptions'
     , '2c_extra_workday_reference_ids_not_in_source'
     , COUNT(*)::VARCHAR
     , CASE
         WHEN COUNT(*) = 1
          AND (SELECT COUNT(*) FROM reconcile_customer_ccs_workday_dummy_row) = 1
         THEN 'EXPECTED - one dummy Workday test row'
         WHEN COUNT(*) = 0
         THEN 'CONFIRMED'
         ELSE 'REVIEW'
       END
  FROM reconcile_customer_ccs_extra_in_workday

 UNION ALL
SELECT '2_population_exceptions'
     , '2d_workday_dummy_rows'
     , COUNT(*)::VARCHAR
     , CASE WHEN COUNT(*) = 1 THEN 'EXPECTED - customer_id 100001 / reference_id CUSTOMER-9-1' ELSE 'REVIEW' END
  FROM reconcile_customer_ccs_workday_dummy_row

 UNION ALL
SELECT '2_population_exceptions'
     , '2e_source_customer_numbers_linked_to_multiple_workday_rows'
     , COUNT(*)::VARCHAR
     , CASE WHEN COUNT(*) = 0 THEN 'CONFIRMED' ELSE 'REVIEW' END
  FROM reconcile_customer_ccs_source_multi_workday

 UNION ALL
SELECT '2_population_exceptions'
     , '2f_merged_customers_split_across_workday_rows'
     , COUNT(*)::VARCHAR
     , CASE WHEN COUNT(*) = 0 THEN 'CONFIRMED' ELSE 'REVIEW' END
  FROM reconcile_customer_ccs_merged_split_workday

 UNION ALL
SELECT '2_population_exceptions'
     , '2g_workday_rows_linked_to_multiple_merged_customers'
     , COUNT(*)::VARCHAR
     , CASE WHEN COUNT(*) = 0 THEN 'CONFIRMED' ELSE 'REVIEW' END
  FROM reconcile_customer_ccs_workday_overlap_merged

 UNION ALL
SELECT '3_attribute_reconciliation'
     , '3_' || t.attribute_name
     , t.matched_count::VARCHAR || ' matched / ' || t.mismatched_count::VARCHAR || ' mismatched'
     , 'Primary lens - sent Workday business attributes vs extract, aligned to source customer_number grain'
  FROM reconcile_customer_ccs_attribute_summary                                     t

 UNION ALL
SELECT '3_tax_classification'
     , '3_' || t.attribute_name || '_' || LOWER(t.comparison_class)
     , t.row_count::VARCHAR
     , CASE
         WHEN t.comparison_class = 'SEMANTIC_EQUIVALENT'
         THEN 'REVIEW - extract wording differs semantically from the sent Workday value'
         WHEN t.comparison_class = 'EXTRACT_OMITS_COMPANY_NUMBER'
         THEN 'REVIEW - extract omits Company Number payload sent to Workday'
         WHEN t.comparison_class = 'EXACT_MATCH'
         THEN 'CONFIRMED'
         ELSE 'REVIEW - extract value differs from the sent Workday value'
       END
  FROM reconcile_customer_ccs_tax_analysis_summary                                  t

 UNION ALL
SELECT '3_tax_review_focus'
     , '3a_tax_review_merged_customers'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconcile_customer_ccs_tax_review_customer

 UNION ALL
SELECT '3_tax_review_focus'
     , '3b_tax_review_source_rows'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconcile_customer_ccs_tax_review_source_rows

 UNION ALL
SELECT '3_tax_review_focus'
     , '3c_tax_review_patterns'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconcile_customer_ccs_tax_review_pattern_summary

 UNION ALL
SELECT '3_tax_review_focus'
     , '3d_tax_review_case_summaries'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconcile_customer_ccs_tax_review_case_summary

 UNION ALL
SELECT '3_tax_review_focus'
     , '3e_tax_review_operator_pack_rows'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconcile_customer_ccs_tax_review_operator_pack

 UNION ALL
SELECT '4_transform_vs_extract'
     , '4a_' || metric
     , row_count::VARCHAR
     , 'Primary lens - ' || status
  FROM reconcile_customer_ccs_transform_extract_tax_diff_summary

 UNION ALL
SELECT '5_provenance'
     , '5a_provenance_case_rows'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconcile_customer_ccs_provenance_case_summary

 UNION ALL
SELECT '5_provenance'
     , '5b_' || provenance_bucket
     , case_count::VARCHAR
     , issue_scope
  FROM reconcile_customer_ccs_provenance_bucket_summary

 UNION ALL
SELECT '6_contact_reconciliation'
     , '6_' || t.attribute_name || '_' || LOWER(t.comparison_class)
     , t.row_count::VARCHAR
     , CASE
         WHEN t.comparison_class = 'EXACT_MATCH'
         THEN 'CONFIRMED'
         WHEN t.comparison_class = 'FORMAT_VARIATION'
         THEN 'REVIEW - extract formatting differs from the sent Workday value'
         WHEN t.comparison_class = 'BOTH_BLANK'
         THEN 'INFO - blank in both transform and extract'
         ELSE 'REVIEW - extract value differs from the sent Workday value'
       END
  FROM reconcile_customer_ccs_contact_summary                                      t

 UNION ALL
SELECT '6_contact_reconciliation'
     , '6_support_' || metric
     , row_count::VARCHAR
     , status
  FROM reconcile_customer_ccs_contact_extract_support

 UNION ALL
SELECT '6_contact_reconciliation'
     , '6_review_case_rows'
     , COUNT(*)::VARCHAR
     , 'Residual address review rows after normalization/equivalence passes'
  FROM reconcile_customer_ccs_contact_review_case_summary

 UNION ALL
SELECT '6_contact_reconciliation'
     , '6_review_patterns'
     , COUNT(*)::VARCHAR
     , 'Grouped residual address-review buckets'
  FROM reconcile_customer_ccs_contact_review_pattern_summary

 UNION ALL
SELECT '6_contact_reconciliation'
     , '6_review_bucket_' || review_bucket
     , row_count::VARCHAR || ' rows / ' || merged_customer_count::VARCHAR || ' merged customers'
     , 'Grouped residual address-review bucket'
  FROM reconcile_customer_ccs_contact_review_pattern_summary

 UNION ALL
SELECT '6_issue_inventory'
     , '6a_total_issue_rows'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconciliation_customer_ccs_issues

 UNION ALL
SELECT '6_issue_inventory'
     , '6b_population_exception_rows'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconciliation_customer_ccs_issues
 WHERE section                                                                      = '2_population_exceptions'

 UNION ALL
SELECT '6_issue_inventory'
     , '6c_exact_attribute_mismatch_rows'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconciliation_customer_ccs_issues
 WHERE issue_type                                                                    = 'ATTRIBUTE_MISMATCH'

 UNION ALL
SELECT '6_issue_inventory'
     , '6d_tax_review_rows'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconciliation_customer_ccs_issues
 WHERE issue_type                                                                    = 'TAX_REVIEW'

 UNION ALL
SELECT '6_issue_inventory'
     , '6e_contact_review_rows'
     , COUNT(*)::VARCHAR
     , CAST(NULL AS VARCHAR)
  FROM reconciliation_customer_ccs_issues
 WHERE issue_type                                                                    = 'CONTACT_REVIEW'

 UNION ALL
SELECT '7_closure'
     , '7a_closure_rows'
     , COUNT(*)::VARCHAR
     , 'One row per mismatch/exception bucket with closure action'
  FROM reconcile_customer_ccs_closure_summary

 UNION ALL
SELECT '7_closure'
     , '7b_operator_review_rows'
     , COUNT(*)::VARCHAR
     , 'Buckets that still require an explicit operator decision'
  FROM reconcile_customer_ccs_closure_summary
 WHERE operator_action                                                              = 'DECIDE_OPERATOR_REVIEW'

 UNION ALL
SELECT '7_closure'
     , '7c_operator_review_buckets'
     , COALESCE(
           STRING_AGG(issue_bucket, ';' ORDER BY closure_order)
         , ''
       )                                                                           metric_value
     , 'Use closure_summary + workbook tabs to decide which buckets become accepted equivalence vs true defects'
  FROM reconcile_customer_ccs_closure_summary
 WHERE operator_action                                                              = 'DECIDE_OPERATOR_REVIEW'

 UNION ALL
SELECT '7_closure'
     , '7d_not_supported_items'
     , COUNT(*)::VARCHAR
     , 'Fields not supported by the current extract (cannot be reconciled)'
  FROM reconcile_customer_ccs_closure_summary
 WHERE sent_vs_extract                                                              = 'NOT_SUPPORTED'

 UNION ALL
SELECT '7_closure'
     , '7e_decision_pack_rows'
     , COUNT(*)::VARCHAR
     , 'Operator decision checklist rows to close the remaining reconcile scope'
  FROM reconcile_customer_ccs_closure_decision_pack

 ORDER BY
       1
     , 2
;
