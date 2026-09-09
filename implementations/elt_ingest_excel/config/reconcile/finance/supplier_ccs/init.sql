-- Supplier GP -> Workday reconciliation (CCS-style review pack).
-- Bridges GP source (src_fin_supplier_base) <-> transform output (workday_supplier_*)
-- <-> live Workday extract (fin_workday_extract_supplier), and classifies the
-- known go-live manual-update exceptions (fin_supplier_manual_update).
--
-- Join key is ALWAYS the GP Vendor ID / reference_id. The Workday supplier_id in
-- the extract is a fresh tenant allocation and is NOT a stable cross-run key.

CREATE OR REPLACE MACRO nrm_name(x) AS
   TRIM(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(COALESCE(x, '')), '[^A-Z0-9]+', ' ', 'g'), ' +', ' ', 'g'))
;

DROP TABLE IF EXISTS reconciliation_supplier_ccs_issues
;
CREATE TABLE reconciliation_supplier_ccs_issues
    (
      section                   VARCHAR
    , issue_type                VARCHAR
    , supplier_number           VARCHAR
    , expected_supplier_id      VARCHAR
    , workday_supplier_id       VARCHAR
    , expected_reference_id     VARCHAR
    , workday_reference_id      VARCHAR
    , attribute_name            VARCHAR
    , expected_value            VARCHAR
    , actual_value              VARCHAR
    , details                   VARCHAR
    )
;

DROP TABLE IF EXISTS reconcile_supplier_ccs_source_profile
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_source_merge_bridge
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_workday_reference_bridge
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_bridge
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_missing_in_workday
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_extra_in_workday
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_workday_dummy_row
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_multi_row_supplier
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_source_multi_workday
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_manual_update
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_manual_update_unmatched
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_manual_update_status
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_finance_action_list
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_expected_supplier
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_actual_supplier
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_attribute_detail
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_attribute_gap
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_attribute_summary
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_merge_group_diff
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_closure_summary
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_report
;
DROP TABLE IF EXISTS reconcile_supplier_ccs_board_summary
;
