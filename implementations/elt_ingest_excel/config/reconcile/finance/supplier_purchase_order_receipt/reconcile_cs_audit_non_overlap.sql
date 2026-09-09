-- ============================================================================
-- OFF-PIPELINE — DO NOT ADD TO order.txt WITHOUT INGEST CONFIG CHANGE FIRST.
-- RECONCILE: Non-overlapping POs — Commitment Stack ONLY or Audit ONLY.
-- Audit-only rows are expected large (audit is a commitments superset of closed/
-- out-of-window POs). Shows count + value so finance can triage.
--
-- DEPENDENCY NOT IN THIS MODULE: references fin_supplier_po_audit_commitment,
-- which is (as at 08.08.2026) NEITHER declared in this module's ingest config
-- (elt_ingest_excel/config/ingest/finance/supplier_purchase_order_receipt.json)
-- NOR produced by the transform SQL folder. The audit extract must be loaded
-- externally before this file can run. Run ad-hoc only.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_audit_non_overlap
;
CREATE TABLE reconcile_cs_audit_non_overlap
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
SELECT b.company_shortcode                                                  company
     , b.order_number                                                       pono
     , COUNT(*)                                                             cs_row_count
     , STRING_AGG(DISTINCT b.currency, ', ' ORDER BY b.currency)           cs_currencies
     , ROUND(MAX(p.cs_po_total), 2)                                         cs_po_total
  FROM src_fin_supplier_po_commitment_stack_base                            b
       INNER JOIN (
         SELECT company_shortcode, order_number,
                ROUND(SUM(COALESCE(net_value_gbp, 0)), 2) cs_po_total
           FROM src_fin_supplier_po_commitment_stack_base
          GROUP BY company_shortcode, order_number
       )                                                                   p
          ON p.company_shortcode = b.company_shortcode
         AND p.order_number    = b.order_number
       INNER JOIN
       cte_scope_entities                                                   e
          ON e.company = b.company_shortcode
 WHERE NULLIF(TRIM(b.order_number), '') IS NOT NULL
 GROUP BY b.company_shortcode, b.order_number
       )
     , cte_audit_po
    AS (
SELECT a.company
     , a.orderno                                                             pono
     , COUNT(*)                                                             audit_row_count
     , MAX(a.companyname)                                                   audit_companyname
     , ROUND(MAX(TRY_CAST(NULLIF(TRIM(a.ordernetvalue), '') AS DOUBLE)), 2) audit_po_total_max
     , ROUND(SUM(TRY_CAST(NULLIF(TRIM(a.soft_commitment), '') AS DOUBLE)), 2)  audit_soft_total
     , ROUND(SUM(TRY_CAST(NULLIF(TRIM(a.hard_commitment), '') AS DOUBLE)), 2)  audit_hard_total
  FROM fin_supplier_po_audit_commitment                                      a
       INNER JOIN
       cte_scope_entities                                                   e
          ON e.company = a.company
 WHERE NULLIF(TRIM(a.orderno), '') IS NOT NULL
 GROUP BY a.company, a.orderno
       )
SELECT 'CS_ONLY'                                                            present_in
     , s.company
     , s.pono
     , s.cs_row_count
     , s.cs_currencies
     , s.cs_po_total
     , NULL::BIGINT                                                         audit_row_count
     , NULL::VARCHAR                                                        audit_companyname
     , NULL::DECIMAL(18,2)                                                  audit_po_total_max
     , NULL::DECIMAL(18,2)                                                  audit_soft_total
     , NULL::DECIMAL(18,2)                                                  audit_hard_total
  FROM cte_cs_po                                                             s
       LEFT OUTER JOIN
       cte_audit_po                                                          a
          ON a.company = s.company
         AND a.pono    = s.pono
 WHERE a.pono IS NULL
UNION ALL
SELECT 'AUDIT_ONLY'
     , a.company
     , a.pono
     , NULL::BIGINT
     , NULL::VARCHAR
     , NULL::DECIMAL(18,2)
     , a.audit_row_count
     , a.audit_companyname
     , a.audit_po_total_max
     , a.audit_soft_total
     , a.audit_hard_total
  FROM cte_audit_po                                                          a
       LEFT OUTER JOIN
       cte_cs_po                                                             s
          ON s.company = a.company
         AND s.pono    = a.pono
 WHERE s.pono IS NULL
 ORDER BY present_in, company, pono
;
