-- Flat metric report (section / metric / metric_value / status), mirroring the
-- customer_ccs report shape so the runner can print a summary and gate on it.

DROP TABLE IF EXISTS reconcile_supplier_ccs_report
;
CREATE TABLE reconcile_supplier_ccs_report
    (
      section        VARCHAR
    , metric         VARCHAR
    , metric_value   BIGINT
    , status         VARCHAR
    )
;

INSERT INTO reconcile_supplier_ccs_report VALUES
  ('1_population', '1a_distinct_gp_vendor_id',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_source_profile), NULL)
, ('1_population', '1b_src_fin_supplier_winning_rows',
   (SELECT COUNT(*) FROM src_fin_supplier WHERE rnk_score = 1), NULL)
, ('1_population', '1c_fin_workday_extract_supplier_rows',
   (SELECT COUNT(*) FROM fin_workday_extract_supplier), NULL)
, ('1_population', '1d_distinct_workday_reference_id',
   (SELECT COUNT(DISTINCT reference_id) FROM fin_workday_extract_supplier
     WHERE NULLIF(TRIM(reference_id), '') IS NOT NULL), NULL)
, ('2_population_exceptions', '2a_gp_missing_in_workday',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_missing_in_workday),
   (SELECT CASE WHEN COUNT(*) = 0 THEN 'PASS' ELSE 'ATTENTION' END FROM reconcile_supplier_ccs_missing_in_workday))
, ('2_population_exceptions', '2b_extra_workday_reference_ids_not_in_gp',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_extra_in_workday),
   (SELECT CASE WHEN COUNT(*) = 0 THEN 'PASS' ELSE 'ATTENTION' END FROM reconcile_supplier_ccs_extra_in_workday))
, ('2_population_exceptions', '2c_workday_dummy_rows',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_workday_dummy_row),
   (SELECT CASE WHEN COUNT(*) = 0 THEN 'PASS' ELSE 'REVIEW' END FROM reconcile_supplier_ccs_workday_dummy_row))
, ('2_population_exceptions', '2d_suppliers_multiple_extract_rows',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_multi_row_supplier), 'INFO')
, ('2_population_exceptions', '2e_merge_group_diff',
   (SELECT COUNT(*) FROM reconciliation_supplier_ccs_issues WHERE issue_type = 'MERGE_GROUP_DIFF'), 'INFO')
, ('3_known_exceptions', '3a_manual_update_total',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update), NULL)
, ('3_known_exceptions', '3b_manual_update_address',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update WHERE exception_type = 'ADDRESS'), NULL)
, ('3_known_exceptions', '3c_manual_update_bank',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update WHERE exception_type = 'BANK'), NULL)
, ('3_known_exceptions', '3d_manual_update_payment',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update WHERE exception_type = 'PAYMENT'), NULL)
, ('3_known_exceptions', '3e_manual_update_unmatched',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update_unmatched),
   (SELECT CASE WHEN COUNT(*) = 0 THEN 'PASS' ELSE 'REVIEW' END FROM reconcile_supplier_ccs_manual_update_unmatched))
, ('3_known_exceptions', '3f_manual_update_still_open_in_workday',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update_status WHERE still_open_in_workday = 'Yes'), 'INFO')
, ('5_finance_actions', '5a_finance_action_list_total',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_finance_action_list), 'ACTION')
, ('4_attribute_gaps', '4a_address_gap_expected',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'ADDRESS' AND gap_status = 'EXPECTED'), 'INFO')
, ('4_attribute_gaps', '4b_address_gap_unexplained',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'ADDRESS' AND gap_status = 'UNEXPLAINED'),
   (SELECT CASE WHEN COUNT(*) = 0 THEN 'PASS' ELSE 'ATTENTION' END FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'ADDRESS' AND gap_status = 'UNEXPLAINED'))
, ('4_attribute_gaps', '4c_bank_gap_expected',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'BANK' AND gap_status = 'EXPECTED'), 'INFO')
, ('4_attribute_gaps', '4d_bank_gap_unexplained',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'BANK' AND gap_status = 'UNEXPLAINED'),
   (SELECT CASE WHEN COUNT(*) = 0 THEN 'PASS' ELSE 'ATTENTION' END FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'BANK' AND gap_status = 'UNEXPLAINED'))
, ('4_attribute_gaps', '4e_payment_gap_expected',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'PAYMENT' AND gap_status = 'EXPECTED'), 'INFO')
, ('4_attribute_gaps', '4f_payment_gap_unexplained',
   (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'PAYMENT' AND gap_status = 'UNEXPLAINED'),
   (SELECT CASE WHEN COUNT(*) = 0 THEN 'PASS' ELSE 'ATTENTION' END FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'PAYMENT' AND gap_status = 'UNEXPLAINED'))
;
