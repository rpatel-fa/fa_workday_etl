-- ============================================================================
-- OFF-PIPELINE — DO NOT ADD TO order.txt WITHOUT INGEST CONFIG CHANGE FIRST.
-- RECONCILE: PO coverage — Commitment Stack vs Audit commitments extract.
-- Scope limited to FA/NFC/WNSL (WPG excluded).
--
-- DEPENDENCY NOT IN THIS MODULE: references fin_supplier_po_audit_commitment,
-- which is (as at 08.08.2026) NOT in this module's ingest config and NOT
-- produced by the transform SQL folder. The audit extract must be loaded
-- externally before this file can run. Run ad-hoc only.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_audit_coverage
;
CREATE TABLE reconcile_cs_audit_coverage
    AS
  WITH cte_scope_entities
    AS (
SELECT 'FA'   company
UNION ALL
SELECT 'NFC'
UNION ALL
SELECT 'WNSL'
       )
     , cte_cs_po
    AS (
SELECT DISTINCT
       b.company_shortcode                                                    company
     , b.order_number                                                         pono
  FROM src_fin_supplier_po_commitment_stack_base                             b
       INNER JOIN
       cte_scope_entities                                                   e
          ON e.company = b.company_shortcode
 WHERE NULLIF(TRIM(b.order_number), '') IS NOT NULL
       )
     , cte_audit_po
    AS (
SELECT DISTINCT
       a.company
     , a.orderno                                                             pono
  FROM fin_supplier_po_audit_commitment                                      a
       INNER JOIN
       cte_scope_entities                                                   e
          ON e.company = a.company
 WHERE NULLIF(TRIM(a.orderno), '') IS NOT NULL
       )
     , cte_by_company
    AS (
SELECT COALESCE(s.company, a.company)                                       company
     , SUM(CASE WHEN s.pono IS NOT NULL THEN 1 ELSE 0 END)                  cs_po_count
     , SUM(CASE WHEN a.pono IS NOT NULL THEN 1 ELSE 0 END)                  audit_po_count
     , SUM(CASE WHEN s.pono IS NOT NULL AND a.pono IS NOT NULL THEN 1 ELSE 0 END)  in_both_po_count
     , SUM(CASE WHEN s.pono IS NOT NULL AND a.pono IS NULL THEN 1 ELSE 0 END)      cs_only_po_count
     , SUM(CASE WHEN s.pono IS NULL AND a.pono IS NOT NULL THEN 1 ELSE 0 END)      audit_only_po_count
  FROM cte_cs_po                                                             s
       FULL OUTER JOIN
       cte_audit_po                                                          a
          ON a.company = s.company
         AND a.pono    = s.pono
 GROUP BY COALESCE(s.company, a.company)
       )
SELECT company
     , cs_po_count
     , audit_po_count
     , in_both_po_count
     , cs_only_po_count
     , audit_only_po_count
  FROM cte_by_company
UNION ALL
SELECT 'TOTAL'                                                              company
     , SUM(cs_po_count)                                                     cs_po_count
     , SUM(audit_po_count)                                                  audit_po_count
     , SUM(in_both_po_count)                                                in_both_po_count
     , SUM(cs_only_po_count)                                                cs_only_po_count
     , SUM(audit_only_po_count)                                             audit_only_po_count
  FROM cte_by_company
;
