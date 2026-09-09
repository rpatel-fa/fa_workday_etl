-- Board summary: the one-glance headline. Population match, known exceptions
-- accounted for, and the count of anything left UNEXPLAINED.

DROP TABLE IF EXISTS reconcile_supplier_ccs_board_summary
;
CREATE TABLE reconcile_supplier_ccs_board_summary
    AS
SELECT * FROM (
  SELECT 1                                                                            sort_order
       , 'GP suppliers (distinct vendor ids)'                                         headline
       , CAST((SELECT COUNT(*) FROM reconcile_supplier_ccs_source_profile) AS VARCHAR) headline_value
  UNION ALL
  SELECT 2, 'Workday extract rows (2026-07-28)'
       , CAST((SELECT COUNT(*) FROM fin_workday_extract_supplier) AS VARCHAR)
  UNION ALL
  SELECT 3, 'Population reconciled (missing + extra)'
       , CASE WHEN (SELECT COUNT(*) FROM reconcile_supplier_ccs_missing_in_workday)
                 + (SELECT COUNT(*) FROM reconcile_supplier_ccs_extra_in_workday) = 0
              THEN 'CLEAN — 0 missing, 0 extra' ELSE 'SEE population sheets' END
  UNION ALL
  SELECT 4, 'Workday dummy / non-GP rows'
       , CAST((SELECT COUNT(*) FROM reconcile_supplier_ccs_workday_dummy_row) AS VARCHAR)
  UNION ALL
  SELECT 5, 'Known manual-update suppliers (address/bank/payment)'
       , CAST((SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update) AS VARCHAR)
       || ' (' || CAST((SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update WHERE exception_type = 'ADDRESS') AS VARCHAR)
       || ' addr / ' || CAST((SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update WHERE exception_type = 'BANK') AS VARCHAR)
       || ' bank / ' || CAST((SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update WHERE exception_type = 'PAYMENT') AS VARCHAR)
       || ' pay)'
  UNION ALL
  SELECT 6, 'Address gaps UNEXPLAINED (not on manual list)'
       , CAST((SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'ADDRESS' AND gap_status = 'UNEXPLAINED') AS VARCHAR)
  UNION ALL
  SELECT 7, 'Bank gaps UNEXPLAINED (not on manual list)'
       , CAST((SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'BANK' AND gap_status = 'UNEXPLAINED') AS VARCHAR)
  UNION ALL
  SELECT 8, 'Payment gaps UNEXPLAINED (not on manual list)'
       , CAST((SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_gap WHERE gap_type = 'PAYMENT' AND gap_status = 'UNEXPLAINED') AS VARCHAR)
) ORDER BY sort_order
;
