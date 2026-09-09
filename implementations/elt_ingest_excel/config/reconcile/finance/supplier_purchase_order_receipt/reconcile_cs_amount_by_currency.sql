-- ============================================================================
-- RECONCILE: Commitment Stack PO VALUE by currency
--   source nrm (base) → staged (2 business drops applied) → workday EIB
--
-- Amount basis MIRRORS the pipeline exactly:
--   base_amount   = soft_commitment_value_local + hard_commitment_value_local (nrm, pre-gates)
--   staged_amount = same sum from src_fin_supplier_po_commitment_stack (post 2 gates + 0.005)
--   workday_amount = column_gf (same sum as staged — 1:1 because unmapped
--                    suppliers are NOT dropped from the EIB; the wd_supplier_id
--                    filter is currently commented in the PO EIB output SQL).
-- Two ACTUAL row-drop gates (applied at staged):
--   DROP_ZERO_COMMITMENT - soft + hard <= 0
--   DROP_OLD_SOFT        - SOFT class + commitment_date < 2025-08-01
-- A third informational bucket tracks unmapped-supplier exposure (they flow
-- through to the EIB — not dropped — and appear in validation_issues instead).
-- Summed as DECIMAL(18,2) both sides so a fully-migrated set ties to the penny.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_amount_by_currency
;
CREATE TABLE reconcile_cs_amount_by_currency
    AS
  WITH cte_base
    AS (
SELECT nrm_company_id                                                     company
     , currency
     , SUM(
         CAST(
           COALESCE(soft_commitment_value_local, 0)
           + COALESCE(hard_commitment_value_local, 0)
           AS DECIMAL(18,2)
         )
       )                                                                  base_amount
  FROM src_fin_supplier_po_commitment_stack_nrm
 GROUP BY nrm_company_id, currency
       )
     , cte_drops_base
    AS (
SELECT nrm_company_id                                                     company
     , currency
     , SUM(
         CASE
           WHEN COALESCE(soft_commitment_value_gbp, 0)
              + COALESCE(hard_commitment_value_gbp, 0) <= 0
           THEN CAST(
                  COALESCE(soft_commitment_value_local, 0)
                + COALESCE(hard_commitment_value_local, 0)
                AS DECIMAL(18,2)
                )
           ELSE CAST(0 AS DECIMAL(18,2))
         END
       )                                                                  drop_zero_commitment_amount
     , SUM(
         CASE
           WHEN COALESCE(soft_commitment_value_gbp, 0)
              + COALESCE(hard_commitment_value_gbp, 0) > 0
            AND UPPER(TRIM(COALESCE(commitment_class, ''))) = 'SOFT'
            AND commitment_date < DATE '2025-08-01'
           THEN CAST(
                  COALESCE(soft_commitment_value_local, 0)
                + COALESCE(hard_commitment_value_local, 0)
                AS DECIMAL(18,2)
                )
           ELSE CAST(0 AS DECIMAL(18,2))
         END
       )                                                                  drop_old_soft_amount
  FROM src_fin_supplier_po_commitment_stack_nrm
 GROUP BY nrm_company_id, currency
       )
     , cte_staged
    AS (
SELECT nrm_company_id                                                     company
     , currency
     , SUM(
         CAST(
           COALESCE(soft_commitment_value_local, 0)
           + COALESCE(hard_commitment_value_local, 0)
           AS DECIMAL(18,2)
         )
       )                                                                  staged_amount
  FROM src_fin_supplier_po_commitment_stack
 GROUP BY nrm_company_id, currency
       )
     , cte_drops_staged
    AS (
SELECT nrm_company_id                                                     company
     , currency
     , SUM(
         CASE
           WHEN wd_supplier_id IS NULL
           THEN CAST(
                  COALESCE(soft_commitment_value_local, 0)
                + COALESCE(hard_commitment_value_local, 0)
                AS DECIMAL(18,2)
                )
           ELSE CAST(0 AS DECIMAL(18,2))
         END
       )                                                                  unmapped_supplier_amount
  FROM src_fin_supplier_po_commitment_stack
 GROUP BY nrm_company_id, currency
       )
     , cte_workday
    AS (
SELECT column_p                                                           company
     , column_ah                                                          currency
     , SUM(TRY_CAST(column_gf AS DECIMAL(18,2)))                          workday_amount
  FROM workday_supplier_po_commitment_stack
 GROUP BY column_p, column_ah
       )
SELECT COALESCE(b.company, d.company, s.company, u.company, w.company)    company
     , COALESCE(b.currency, d.currency, s.currency, u.currency, w.currency) currency
     , COALESCE(b.base_amount, 0)                                         base_amount
     , COALESCE(d.drop_zero_commitment_amount, 0)                         drop_zero_commitment_amount
     , COALESCE(d.drop_old_soft_amount, 0)                                drop_old_soft_amount
     , COALESCE(b.base_amount, 0)
       - COALESCE(d.drop_zero_commitment_amount, 0)
       - COALESCE(d.drop_old_soft_amount, 0)                              expected_to_staged
     , COALESCE(s.staged_amount, 0)                                       staged_amount
     , COALESCE(b.base_amount, 0)
       - COALESCE(d.drop_zero_commitment_amount, 0)
       - COALESCE(d.drop_old_soft_amount, 0)
       - COALESCE(s.staged_amount, 0)                                     unexplained_staged_amount
     , COALESCE(u.unmapped_supplier_amount, 0)                            unmapped_supplier_amount
     , COALESCE(s.staged_amount, 0)                                       expected_to_workday
     , COALESCE(w.workday_amount, 0)                                      workday_amount
     , COALESCE(s.staged_amount, 0)
       - COALESCE(w.workday_amount, 0)                                    unexplained_workday_amount
  FROM cte_base                                                           b
       FULL OUTER JOIN
       cte_drops_base                                                     d
          ON d.company  = b.company
         AND d.currency = b.currency
       FULL OUTER JOIN
       cte_staged                                                         s
          ON s.company  = COALESCE(b.company, d.company)
         AND s.currency = COALESCE(b.currency, d.currency)
       FULL OUTER JOIN
       cte_drops_staged                                                   u
          ON u.company  = COALESCE(b.company, d.company, s.company)
         AND u.currency = COALESCE(b.currency, d.currency, s.currency)
       FULL OUTER JOIN
       cte_workday                                                        w
          ON w.company  = COALESCE(b.company, d.company, s.company, u.company)
         AND w.currency = COALESCE(b.currency, d.currency, s.currency, u.currency)
 ORDER BY company, currency
;

INSERT INTO reconciliation_supplier_po_receipt_issues
SELECT 'reconcile_cs_amount_by_currency'                                   table_name
     , 'UNEXPLAINED_STAGED_AMOUNT'                                        rule_type
     , currency                                                           col
     , company                                                            recon_key
     , unexplained_staged_amount::VARCHAR                                 detail
     , 'Base minus (zero-commitment + old-soft) amount does not tie to staged.' message
  FROM reconcile_cs_amount_by_currency
 WHERE ABS(unexplained_staged_amount) > 0.005
 UNION ALL
SELECT 'reconcile_cs_amount_by_currency'                                   table_name
     , 'UNEXPLAINED_WORKDAY_AMOUNT'                                       rule_type
     , currency                                                           col
     , company                                                            recon_key
     , unexplained_workday_amount::VARCHAR                                detail
     , 'Staged amount does not tie to Workday extended_amount (expected 1:1 — unmapped suppliers no longer dropped).' message
  FROM reconcile_cs_amount_by_currency
 WHERE ABS(unexplained_workday_amount) > 0.005
;
