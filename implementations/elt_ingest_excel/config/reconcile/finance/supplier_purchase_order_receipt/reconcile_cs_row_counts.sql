-- ============================================================================
-- RECONCILE: Commitment Stack PO ROW COUNTS
--   source ingest (base) → nrm (supplier lookup) → staged (filters) → workday EIB
-- Row lifecycle is base → staged (2 business drops) → workday EIB (1:1 — unmapped
-- suppliers are NOT dropped from the EIB; the wd_supplier_id filter is currently
-- commented in workday_supplier_po_commitment_stack.sql, line 81, and unmapped
-- suppliers flow to the EIB with NULL wd_supplier_id, surfacing in validation
-- issues for finance review. The two ACTUAL row-drop gates applied at staged are:
--   DROP_ZERO_COMMITMENT - soft + hard <= 0 (no remaining commitment value)
--   DROP_OLD_SOFT        - SOFT class + commitment_date < 2025-08-01 (cutoff rule)
--   unmapped_supplier_rows - informational only (NOT dropped — they reach Workday)
-- unexplained_staged = base_rows - zero-commitment drops - old-soft - staged_rows.
-- unexplained_workday = staged_rows - workday_rows (must be 0 — 1:1 right now).
-- unexplained <> 0 is a reconciliation break.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_row_counts
;
CREATE TABLE reconcile_cs_row_counts
    AS
  WITH cte_base
    AS (
SELECT nrm_company_id                                                     company
     , COUNT(*)                                                           base_rows
  FROM src_fin_supplier_po_commitment_stack_nrm
 GROUP BY nrm_company_id
       )
     , cte_staged
    AS (
SELECT nrm_company_id                                                     company
     , COUNT(*)                                                           staged_rows
  FROM src_fin_supplier_po_commitment_stack
 GROUP BY nrm_company_id
       )
     , cte_drops_staged
    AS (
SELECT nrm_company_id                                                     company
     , COUNT(*) FILTER (WHERE wd_supplier_id IS NULL)                     unmapped_supplier_rows
  FROM src_fin_supplier_po_commitment_stack
 GROUP BY nrm_company_id
       )
     , cte_drops
    AS (
SELECT n.nrm_company_id                                                   company
     , COUNT(*) FILTER (
         WHERE COALESCE(n.soft_commitment_value_gbp, 0)
             + COALESCE(n.hard_commitment_value_gbp, 0) <= 0
       )                                                                  drop_zero_commitment
     , COUNT(*) FILTER (
         WHERE COALESCE(n.soft_commitment_value_gbp, 0)
             + COALESCE(n.hard_commitment_value_gbp, 0) > 0
           AND UPPER(TRIM(COALESCE(n.commitment_class, ''))) = 'SOFT'
           AND n.commitment_date < DATE '2025-08-01'
       )                                                                  drop_old_soft
     , COALESCE(s.unmapped_supplier_rows, 0)                              unmapped_supplier_rows
  FROM src_fin_supplier_po_commitment_stack_nrm                           n
       LEFT OUTER JOIN
       cte_drops_staged                                                   s
          ON s.company = n.nrm_company_id
 GROUP BY n.nrm_company_id, s.unmapped_supplier_rows
       )
     , cte_workday
    AS (
SELECT column_p                                                           company
     , COUNT(*)                                                           workday_rows
  FROM workday_supplier_po_commitment_stack
 GROUP BY column_p
       )
SELECT COALESCE(b.company, d.company, s.company, w.company)               company
     , COALESCE(b.base_rows, 0)                                           base_rows
     , COALESCE(d.drop_zero_commitment, 0)                                drop_zero_commitment
     , COALESCE(d.drop_old_soft, 0)                                       drop_old_soft
     , COALESCE(b.base_rows, 0)
       - COALESCE(d.drop_zero_commitment, 0)
       - COALESCE(d.drop_old_soft, 0)                                     expected_to_staged
     , COALESCE(s.staged_rows, 0)                                         staged_rows
     , COALESCE(b.base_rows, 0)
       - COALESCE(d.drop_zero_commitment, 0)
       - COALESCE(d.drop_old_soft, 0)
       - COALESCE(s.staged_rows, 0)                                       unexplained_staged
     , COALESCE(d.unmapped_supplier_rows, 0)                              unmapped_supplier_rows
     , COALESCE(s.staged_rows, 0)                                         expected_to_workday
     , COALESCE(w.workday_rows, 0)                                        workday_rows
     , COALESCE(s.staged_rows, 0)
       - COALESCE(w.workday_rows, 0)                                      unexplained_workday
  FROM cte_base                                                           b
       FULL OUTER JOIN
       cte_drops                                                          d
          ON d.company = b.company
       FULL OUTER JOIN
       cte_staged                                                         s
          ON s.company = COALESCE(b.company, d.company)
       FULL OUTER JOIN
       cte_workday                                                        w
          ON w.company = COALESCE(b.company, d.company, s.company)
 ORDER BY company
;

INSERT INTO reconciliation_supplier_po_receipt_issues
SELECT 'reconcile_cs_row_counts'                                          table_name
     , 'UNEXPLAINED_STAGED_ROWS'                                         rule_type
     , 'row'                                                              col
     , company                                                            recon_key
     , unexplained_staged::VARCHAR                                        detail
     , 'Base rows minus (zero-commitment + old-soft) do not tie to staged.' message
  FROM reconcile_cs_row_counts
 WHERE unexplained_staged <> 0
 UNION ALL
SELECT 'reconcile_cs_row_counts'                                          table_name
     , 'UNEXPLAINED_WORKDAY_ROWS'                                        rule_type
     , 'row'                                                              col
     , company                                                            recon_key
     , unexplained_workday::VARCHAR                                       detail
     , 'Staged rows do not tie to Workday EIB rows (expected 1:1 — unmapped suppliers no longer dropped).' message
  FROM reconcile_cs_row_counts
 WHERE unexplained_workday <> 0
;
