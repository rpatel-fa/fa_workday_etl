-- ============================================================================
-- REVIEW PACK: unmapped suppliers + workday-cleanup actions finance view.
-- Consolidates the 3 permitted-drop reports into a single finance-friendly
-- action list with total-value weighting so the largest items are first.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_finance_actions
;
CREATE TABLE reconcile_cs_finance_actions
    AS
SELECT 'UNMAPPED_SUPPLIER'                                                    action_type
     , company || ' | ' || supplier_code                                     recon_key
     , supplier_name
     , row_count
     , total_commitment_gbp                                                   value_gbp
     , 'Add supplier reference to Workday map or confirm out of scope.'    suggested_action
  FROM reconcile_cs_unmapped_suppliers

 UNION ALL
SELECT 'ZERO_COMMITMENT_PO'
     , company || ' | ' || order_number
     , supplier_name
     , row_count
     , total_net_value_gbp
     , 'No remaining commitment (soft + hard = 0) — confirm PO is fully closed.'
  FROM reconcile_cs_zero_commitments

 UNION ALL
SELECT 'OLD_SOFT_CUTOFF_PO'
     , company || ' | ' || order_number
     , supplier_name
     , row_count
     , total_value_gbp
     , 'SOFT commitment, commitment_date < 2025-08-01 — excluded by cutoff. Confirm no Workday load required.'
  FROM reconcile_cs_old_soft_cutoff

 ORDER BY action_type, value_gbp DESC, recon_key
;
