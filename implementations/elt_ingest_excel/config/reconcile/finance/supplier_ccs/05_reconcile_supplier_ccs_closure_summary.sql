-- Closure summary: one line per reconcile category with its count and a
-- pass/attention status, so the pack can be read top-down.

DROP TABLE IF EXISTS reconcile_supplier_ccs_closure_summary
;
CREATE TABLE reconcile_supplier_ccs_closure_summary
    AS
SELECT * FROM (
  SELECT 1                                                                            sort_order
       , 'Population'                                                                 category
       , 'GP suppliers missing in Workday'                                            measure
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_missing_in_workday)             item_count
       , CASE WHEN (SELECT COUNT(*) FROM reconcile_supplier_ccs_missing_in_workday) = 0
              THEN 'PASS' ELSE 'ATTENTION' END                                        status
  UNION ALL
  SELECT 2, 'Population', 'Workday reference ids not in GP source'
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_extra_in_workday)
       , CASE WHEN (SELECT COUNT(*) FROM reconcile_supplier_ccs_extra_in_workday) = 0
              THEN 'PASS' ELSE 'ATTENTION' END
  UNION ALL
  SELECT 3, 'Population', 'Workday dummy / non-GP rows'
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_workday_dummy_row)
       , CASE WHEN (SELECT COUNT(*) FROM reconcile_supplier_ccs_workday_dummy_row) = 0
              THEN 'PASS' ELSE 'REVIEW' END
  UNION ALL
  SELECT 4, 'Population', 'Suppliers spanning multiple extract rows'
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_multi_row_supplier)
       , 'INFO'
  UNION ALL
  SELECT 5, 'Population', 'GP vendor ids merged differently in Workday'
       , (SELECT COUNT(*) FROM reconciliation_supplier_ccs_issues WHERE issue_type = 'MERGE_GROUP_DIFF')
       , 'INFO'
  UNION ALL
  SELECT 6, 'Known exceptions', 'Manual-update suppliers (address removed)'
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update WHERE exception_type = 'ADDRESS')
       , 'INFO'
  UNION ALL
  SELECT 7, 'Known exceptions', 'Manual-update suppliers (bank removed)'
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update WHERE exception_type = 'BANK')
       , 'INFO'
  UNION ALL
  SELECT 8, 'Known exceptions', 'Manual-update suppliers (payment set to manual)'
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update WHERE exception_type = 'PAYMENT')
       , 'INFO'
  UNION ALL
  SELECT 9, 'Known exceptions', 'Manual-update names not matched to extract'
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update_unmatched)
       , CASE WHEN (SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update_unmatched) = 0
              THEN 'PASS' ELSE 'REVIEW' END
  UNION ALL
  SELECT 10, 'Attribute gaps', 'Address gaps — expected (manual list)'
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'ADDRESS' AND gap_status = 'EXPECTED')
       , 'INFO'
  UNION ALL
  SELECT 11, 'Attribute gaps', 'Address gaps — UNEXPLAINED'
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'ADDRESS' AND gap_status = 'UNEXPLAINED')
       , CASE WHEN (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'ADDRESS' AND gap_status = 'UNEXPLAINED') = 0
              THEN 'PASS' ELSE 'ATTENTION' END
  UNION ALL
  SELECT 12, 'Attribute gaps', 'Bank gaps — expected (manual list)'
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'BANK' AND gap_status = 'EXPECTED')
       , 'INFO'
  UNION ALL
  SELECT 13, 'Attribute gaps', 'Bank gaps — UNEXPLAINED'
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'BANK' AND gap_status = 'UNEXPLAINED')
       , CASE WHEN (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'BANK' AND gap_status = 'UNEXPLAINED') = 0
              THEN 'PASS' ELSE 'ATTENTION' END
  UNION ALL
  SELECT 14, 'Attribute gaps', 'Payment set to manual — expected (manual list)'
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'PAYMENT' AND gap_status = 'EXPECTED')
       , 'INFO'
  UNION ALL
  SELECT 15, 'Attribute gaps', 'Payment set to manual — UNEXPLAINED'
       , (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'PAYMENT' AND gap_status = 'UNEXPLAINED')
       , CASE WHEN (SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'PAYMENT' AND gap_status = 'UNEXPLAINED') = 0
              THEN 'PASS' ELSE 'ATTENTION' END
) ORDER BY sort_order
;
