DROP TABLE IF EXISTS reconcile_customer_ccs_closure_summary
;
CREATE TABLE reconcile_customer_ccs_closure_summary
    AS
SELECT 1                                                                          closure_order
     , 'population'                                                               issue_area
     , 'workday_dummy_row'                                                        issue_bucket
     , 'exception'                                                                issue_scope
     , 'CONFIRMED_SENT_VS_EXTRACT_DIFF'                                           sent_vs_extract
     , CAST(NULL AS BIGINT)                                                       source_customer_number_count
     , CAST(NULL AS BIGINT)                                                       merged_customer_count
     , COUNT(*)                                                                   matched_customer_id_count
     , 'reconcile_customer_ccs_workday_dummy_row'                                 evidence_table
     , 'ACCEPT_EXCEPTION'                                                         operator_action
     , 'Workday test record present in extract only; keep isolated from defect counts'
                                                                                    operator_notes
  FROM reconcile_customer_ccs_workday_dummy_row

 UNION ALL
SELECT 10 + review_order                                                          closure_order
     , 'tax'                                                                      issue_area
     , provenance_bucket                                                          issue_bucket
     , issue_scope
     , 'CONFIRMED_SENT_VS_EXTRACT_DIFF'                                           sent_vs_extract
     , source_customer_number_count
     , merged_customer_count
     , matched_customer_id_count
     , 'reconcile_customer_ccs_provenance_case_summary'                           evidence_table
     , CASE
         WHEN issue_scope = 'systematic'
          AND provenance_bucket = 'semantic_wording_variation'
         THEN 'ACCEPT_EQUIVALENCE'
         WHEN provenance_bucket = 'accepted_workday_normalisation'
         THEN 'ACCEPT_WORKDAY_BEHAVIOUR'
         WHEN issue_scope = 'systematic'
         THEN 'ACCEPT_EXTRACT_BEHAVIOUR'
         ELSE 'DECIDE_OPERATOR_REVIEW'
       END                                                                        operator_action
     , CASE
         WHEN provenance_bucket = 'accepted_workday_normalisation'
         THEN 'Workday reports Isle of Man as United Kingdom; treat this as accepted Workday handling rather than a transform defect'
         WHEN provenance_bucket = 'workday_target_state_review'
         THEN 'Workday tax country appears to follow the Singapore primary address; if the business expects UK, raise this as a Workday issue rather than a transform defect'
         ELSE provenance_explanation || ' | ' || suggested_operator_action
       END                                                                        operator_notes
  FROM reconcile_customer_ccs_provenance_case_summary

 UNION ALL
SELECT 100                                                                         closure_order
     , 'contact'                                                                  issue_area
     , 'default_address_format_variation'                                         issue_bucket
     , 'systematic'                                                               issue_scope
     , 'CONFIRMED_SENT_VS_EXTRACT_DIFF'                                           sent_vs_extract
     , CAST(NULL AS BIGINT)                                                       source_customer_number_count
     , CAST(NULL AS BIGINT)                                                       merged_customer_count
     , row_count                                                                   matched_customer_id_count
     , 'reconcile_customer_ccs_contact_summary'                                   evidence_table
     , 'ACCEPT_EQUIVALENCE'                                                       operator_action
     , 'Extract formatting/layout differs from sent address; captured as FORMAT_VARIATION'
                                                                                    operator_notes
  FROM reconcile_customer_ccs_contact_summary
 WHERE attribute_name                                                              = 'default_address'
   AND comparison_class                                                            = 'FORMAT_VARIATION'

 UNION ALL
SELECT 110                                                                         closure_order
     , 'contact'
     , 'default_address_value_mismatch'
     , 'operator_review'
     , 'CONFIRMED_SENT_VS_EXTRACT_DIFF'
     , CAST(NULL AS BIGINT)
     , CAST(NULL AS BIGINT)
     , row_count
     , 'reconcile_customer_ccs_contact_review_case_summary'
     , 'DECIDE_OPERATOR_REVIEW'
     , 'Residual address mismatches after all normalisation/equivalence passes; review by bucket'
  FROM reconcile_customer_ccs_contact_summary
 WHERE attribute_name                                                              = 'default_address'
   AND comparison_class                                                            = 'VALUE_MISMATCH'

 UNION ALL
SELECT 120 + ROW_NUMBER() OVER(ORDER BY review_bucket)                             closure_order
     , 'contact'
     , 'address_review_bucket:' || review_bucket                                  issue_bucket
     , 'operator_review'
     , 'CONFIRMED_SENT_VS_EXTRACT_DIFF'
     , row_count                                                                   source_customer_number_count
     , merged_customer_count
     , CAST(NULL AS BIGINT)                                                       matched_customer_id_count
     , 'reconcile_customer_ccs_contact_review_pattern_summary'
     , 'DECIDE_OPERATOR_REVIEW'
     , 'Bucketed residual address review queue'
  FROM reconcile_customer_ccs_contact_review_pattern_summary

 UNION ALL
SELECT 200                                                                         closure_order
     , 'contact'
     , 'default_phone_format_variation'
     , 'systematic'
     , 'CONFIRMED_SENT_VS_EXTRACT_DIFF'
     , CAST(NULL AS BIGINT)
     , CAST(NULL AS BIGINT)
     , row_count
     , 'reconcile_customer_ccs_contact_summary'
     , 'ACCEPT_EQUIVALENCE'
     , 'Extract phone display formatting differs from sent primary phone; no remaining phone review rows'
  FROM reconcile_customer_ccs_contact_summary
 WHERE attribute_name                                                              = 'default_phone_number'
   AND comparison_class                                                            = 'FORMAT_VARIATION'

 UNION ALL
SELECT 300                                                                         closure_order
     , 'contact'
     , 'email_support_in_extract'
     , 'info'
     , CASE
         WHEN MAX(CASE WHEN metric = 'customer_contacts_rows_in_extract' THEN status END)
              = 'EMAIL_NOT_SUPPORTED_IN_CURRENT_EXTRACT'
         THEN 'NOT_SUPPORTED'
         ELSE 'REVIEW'
       END                                                                        sent_vs_extract
     , CAST(NULL AS BIGINT)
     , CAST(NULL AS BIGINT)
     , MAX(CASE WHEN metric = 'sent_email_customer_ids_in_transform' THEN row_count END)
     , 'reconcile_customer_ccs_contact_extract_support'
     , CASE
         WHEN MAX(CASE WHEN metric = 'customer_contacts_rows_in_extract' THEN status END)
              = 'EMAIL_NOT_SUPPORTED_IN_CURRENT_EXTRACT'
         THEN 'ACCEPT_NOT_SUPPORTED'
         ELSE 'DECIDE_OPERATOR_REVIEW'
       END
     , MAX(CASE WHEN metric = 'sent_email_customer_ids_in_transform' THEN status END)
  FROM reconcile_customer_ccs_contact_extract_support
;
