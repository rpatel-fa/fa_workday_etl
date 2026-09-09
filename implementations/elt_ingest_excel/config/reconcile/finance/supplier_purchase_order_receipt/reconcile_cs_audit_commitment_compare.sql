-- ============================================================================
-- OFF-PIPELINE — DO NOT ADD TO order.txt WITHOUT INGEST CONFIG CHANGE FIRST.
-- RECONCILE: Soft/Hard commitment comparison vs Audit on matched POs (BOTH set).
-- Per-PO granularity (the audit has no position grain). Commitment totals are
-- summed per PO on BOTH sides, then compared with a £2 tolerance bucket.
-- Soft/Hard tie-out proven on GBP subset in the prior pipeline (95.4% within £2);
-- the reconcile surfaces the DIFF tail for finance review.
--
-- DEPENDENCY NOT IN THIS MODULE: references fin_supplier_po_audit_commitment,
-- which is (as at 08.08.2026) NEITHER declared in this module's ingest config
-- (elt_ingest_excel/config/ingest/finance/supplier_purchase_order_receipt.json)
-- NOR produced by the transform SQL folder. The audit extract must be loaded
-- externally before this file can run. Run ad-hoc only.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_audit_commitment_compare
;
CREATE TABLE reconcile_cs_audit_commitment_compare
    AS
  WITH cte_scope
    AS (
SELECT 'FA'   company UNION ALL SELECT 'NFC' UNION ALL SELECT 'WNSL'
       )
     , cte_cs
    AS (
SELECT n.company_shortcode                                                  company
     , n.order_number                                                       pono
     , STRING_AGG(DISTINCT n.currency, ', ' ORDER BY n.currency)           currencies
     , ROUND(SUM(COALESCE(n.soft_commitment_value_gbp, 0)), 2)              cs_soft
     , ROUND(SUM(COALESCE(n.hard_commitment_value_gbp, 0)), 2)              cs_hard
     , ROUND(SUM(
         COALESCE(n.soft_commitment_value_gbp, 0)
       + COALESCE(n.hard_commitment_value_gbp, 0)
       ), 2)                                                                cs_total
  FROM src_fin_supplier_po_commitment_stack_nrm                             n
       INNER JOIN cte_scope e ON e.company = n.company_shortcode
 GROUP BY n.company_shortcode, n.order_number
       )
     , cte_audit
    AS (
SELECT a.company
     , a.orderno                                                             pono
     , ROUND(SUM(TRY_CAST(NULLIF(TRIM(a.soft_commitment), '') AS DOUBLE)), 2) audit_soft
     , ROUND(SUM(TRY_CAST(NULLIF(TRIM(a.hard_commitment), '') AS DOUBLE)), 2) audit_hard
     , ROUND(SUM(
         COALESCE(TRY_CAST(NULLIF(TRIM(a.soft_commitment), '') AS DOUBLE), 0)
       + COALESCE(TRY_CAST(NULLIF(TRIM(a.hard_commitment), '') AS DOUBLE), 0)
       ), 2)                                                                audit_total
  FROM fin_supplier_po_audit_commitment                                      a
       INNER JOIN cte_scope e ON e.company = a.company
 WHERE NULLIF(TRIM(a.orderno), '') IS NOT NULL
 GROUP BY a.company, a.orderno
       )
SELECT s.company
     , s.pono
     , s.currencies
     , s.cs_soft
     , a.audit_soft
     , ROUND(s.cs_soft - COALESCE(a.audit_soft, 0), 2)                      soft_delta
     , CASE WHEN ABS(s.cs_soft - COALESCE(a.audit_soft, 0)) <= 2.00 THEN 'MATCH_2GBP' ELSE 'DIFF' END soft_match
     , s.cs_hard
     , a.audit_hard
     , ROUND(s.cs_hard - COALESCE(a.audit_hard, 0), 2)                     hard_delta
     , CASE WHEN ABS(s.cs_hard - COALESCE(a.audit_hard, 0)) <= 2.00 THEN 'MATCH_2GBP' ELSE 'DIFF' END hard_match
     , s.cs_total
     , a.audit_total
     , ROUND(s.cs_total - COALESCE(a.audit_total, 0), 2)                   total_delta
     , CASE WHEN ABS(s.cs_total - COALESCE(a.audit_total, 0)) <= 2.00 THEN 'MATCH_2GBP' ELSE 'DIFF' END total_match
  FROM cte_cs                                                                s
       INNER JOIN
       cte_audit                                                             a
          ON a.company = s.company
         AND a.pono    = s.pono
 ORDER BY total_match, s.company, s.pono
;
