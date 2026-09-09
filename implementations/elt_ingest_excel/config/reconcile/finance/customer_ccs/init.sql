DROP TABLE IF EXISTS reconciliation_customer_ccs_issues
;
CREATE TABLE reconciliation_customer_ccs_issues
    (
      section                   VARCHAR
    , issue_type                VARCHAR
    , customer_number           VARCHAR
    , expected_customer_id      VARCHAR
    , workday_customer_id       VARCHAR
    , expected_reference_id     VARCHAR
    , workday_reference_id      VARCHAR
    , attribute_name            VARCHAR
    , expected_value            VARCHAR
    , actual_value              VARCHAR
    , details                   VARCHAR
    )
;

DROP TABLE IF EXISTS reconcile_customer_ccs_source_profile
;
DROP TABLE IF EXISTS reconcile_customer_ccs_source_merge_bridge
;
DROP TABLE IF EXISTS reconcile_customer_ccs_workday_reference_bridge
;
DROP TABLE IF EXISTS reconcile_customer_ccs_bridge
;
DROP TABLE IF EXISTS reconcile_customer_ccs_expected_customer
;
DROP TABLE IF EXISTS reconcile_customer_ccs_missing_in_workday
;
DROP TABLE IF EXISTS reconcile_customer_ccs_extra_in_workday
;
DROP TABLE IF EXISTS reconcile_customer_ccs_workday_dummy_row
;
DROP TABLE IF EXISTS reconcile_customer_ccs_source_multi_workday
;
DROP TABLE IF EXISTS reconcile_customer_ccs_merged_split_workday
;
DROP TABLE IF EXISTS reconcile_customer_ccs_workday_overlap_merged
;
DROP TABLE IF EXISTS reconcile_customer_ccs_attribute_detail
;
DROP TABLE IF EXISTS reconcile_customer_ccs_attribute_summary
;
DROP TABLE IF EXISTS reconcile_customer_ccs_tax_analysis
;
DROP TABLE IF EXISTS reconcile_customer_ccs_tax_analysis_summary
;
DROP TABLE IF EXISTS reconcile_customer_ccs_tax_review_customer
;
DROP TABLE IF EXISTS reconcile_customer_ccs_tax_review_source_rows
;
DROP TABLE IF EXISTS reconcile_customer_ccs_tax_review_pattern_summary
;
DROP TABLE IF EXISTS reconcile_customer_ccs_tax_review_case_summary
;
DROP TABLE IF EXISTS reconcile_customer_ccs_tax_review_operator_pack
;
DROP TABLE IF EXISTS reconcile_customer_ccs_transform_extract_tax_diff_detail
;
DROP TABLE IF EXISTS reconcile_customer_ccs_transform_extract_tax_diff_summary
;
DROP TABLE IF EXISTS reconcile_customer_ccs_transform_extract_tax_diff_pattern_summary
;
DROP TABLE IF EXISTS reconcile_customer_ccs_provenance_case_summary
;
DROP TABLE IF EXISTS reconcile_customer_ccs_provenance_bucket_summary
;
DROP TABLE IF EXISTS reconcile_customer_ccs_contact_detail
;
DROP TABLE IF EXISTS reconcile_customer_ccs_contact_summary
;
DROP TABLE IF EXISTS reconcile_customer_ccs_contact_extract_support
;
DROP TABLE IF EXISTS reconcile_customer_ccs_contact_review_case_summary
;
DROP TABLE IF EXISTS reconcile_customer_ccs_contact_review_pattern_summary
;
DROP TABLE IF EXISTS reconcile_customer_ccs_closure_summary
;
DROP TABLE IF EXISTS reconcile_customer_ccs_closure_decision_pack
;
DROP TABLE IF EXISTS reconcile_customer_ccs_report
;
DROP TABLE IF EXISTS reconcile_customer_ccs_board_summary
;
