-- ============================================================================
-- OFF-PIPELINE — DO NOT ADD TO order.txt WITHOUT INGEST CONFIG CHANGE FIRST.
-- RECONCILE: PO reconcile (ALL scenarios) — Commitment Stack vs Audit.
-- One row per PO across both datasets, tagged present_in = BOTH / CS_ONLY /
-- AUDIT_ONLY, with PO total tie-out + commitment_class comparison BOTH sides.
--
-- DEPENDENCY NOT IN THIS MODULE: references fin_supplier_po_audit_commitment,
-- which is (as at 08.08.2026) NEITHER declared in this module's ingest config
-- (elt_ingest_excel/config/ingest/finance/supplier_purchase_order_receipt.json)
-- NOR produced by the transform SQL folder. The audit extract must be loaded
-- externally before this file can run. Run ad-hoc only.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_audit_reconcile
;
CREATE TABLE reconcile_cs_audit_reconcile
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
     , COUNT(DISTINCT TRY_CAST(NULLIF(TRIM(a.ordernetvalue), '') AS DOUBLE))  audit_distinct_total_count
     , ROUND(MIN(TRY_CAST(NULLIF(TRIM(a.ordernetvalue), '') AS DOUBLE)), 2) audit_po_total_min
     , ROUND(MAX(TRY_CAST(NULLIF(TRIM(a.ordernetvalue), '') AS DOUBLE)), 2) audit_po_total_max
     , ROUND(SUM(TRY_CAST(NULLIF(TRIM(a.soft_commitment), '') AS DOUBLE)), 2)  audit_soft_total
     , ROUND(SUM(TRY_CAST(NULLIF(TRIM(a.hard_commitment), '') AS DOUBLE)), 2)  audit_hard_total
       -- commitment class from audit (TRIM trailing space on 'Soft ')
     , CASE
           WHEN MAX(CASE WHEN TRIM(a.commitmenttype) = 'Soft' THEN 1 ELSE 0 END) = 1
            AND MAX(CASE WHEN TRIM(a.commitmenttype) = 'Hard' THEN 1 ELSE 0 END) = 1
           THEN 'MIXED'
           WHEN MAX(CASE WHEN TRIM(a.commitmenttype) = 'Soft' THEN 1 ELSE 0 END) = 1
           THEN 'SOFT'
           WHEN MAX(CASE WHEN TRIM(a.commitmenttype) = 'Hard' THEN 1 ELSE 0 END) = 1
           THEN 'HARD'
           ELSE 'NONE'
       END                                                                  audit_commitment_class
       -- inferred currency for AUDIT_ONLY rows
     , CASE
           WHEN SUM(TRY_CAST(NULLIF(TRIM(a.invoicednominalamount), '') AS DOUBLE)) IS NULL
             OR SUM(TRY_CAST(NULLIF(TRIM(a.invoicednominalamount), '') AS DOUBLE)) = 0
           THEN 'UNKNOWN'
           WHEN ABS(
                    SUM(TRY_CAST(NULLIF(TRIM(a.invoicednominalamount), '') AS DOUBLE))
                  - COALESCE(SUM(TRY_CAST(NULLIF(TRIM(a.invoicednominalamountgbp), '') AS DOUBLE)), 0)
                ) > 0.01 * ABS(SUM(TRY_CAST(NULLIF(TRIM(a.invoicednominalamount), '') AS DOUBLE)))
           THEN 'NON_GBP'
           ELSE 'GBP'
       END                                                                  audit_currency
  FROM fin_supplier_po_audit_commitment                                      a
       INNER JOIN
       cte_scope_entities                                                   e
          ON e.company = a.company
 WHERE NULLIF(TRIM(a.orderno), '') IS NOT NULL
 GROUP BY a.company, a.orderno
       )
     , cte_cs_commitment
    AS (
SELECT n.company_shortcode                                                  company
     , n.order_number                                                       pono
     , ROUND(SUM(COALESCE(n.soft_commitment_value_gbp, 0)), 2)              cs_soft_total
     , ROUND(SUM(COALESCE(n.hard_commitment_value_gbp, 0)), 2)              cs_hard_total
     , CASE
           WHEN ROUND(SUM(COALESCE(n.soft_commitment_value_gbp, 0)), 2) > 0
            AND ROUND(SUM(COALESCE(n.hard_commitment_value_gbp, 0)), 2) > 0
           THEN 'MIXED'
           WHEN ROUND(SUM(COALESCE(n.soft_commitment_value_gbp, 0)), 2) > 0
           THEN 'SOFT'
           WHEN ROUND(SUM(COALESCE(n.hard_commitment_value_gbp, 0)), 2) > 0
           THEN 'HARD'
           ELSE 'NONE'
       END                                                                  cs_commitment_class
  FROM src_fin_supplier_po_commitment_stack_nrm                             n
 GROUP BY n.company_shortcode, n.order_number
       )
SELECT COALESCE(s.company, a.company)                                       company
     , COALESCE(s.pono, a.pono)                                             pono
     , CASE
           WHEN s.pono IS NOT NULL AND a.pono IS NOT NULL THEN 'BOTH'
           WHEN s.pono IS NOT NULL                            THEN 'CS_ONLY'
           ELSE                                                        'AUDIT_ONLY'
       END                                                                  present_in
     , COALESCE(s.cs_currencies, a.audit_currency)                         po_currency
     , s.cs_row_count
     , s.cs_currencies
     , s.cs_po_total
     , a.audit_row_count
     , a.audit_companyname
     , a.audit_distinct_total_count
     , a.audit_po_total_min
     , a.audit_po_total_max
     , a.audit_commitment_class
     , a.audit_soft_total
     , a.audit_hard_total
     , cc.cs_commitment_class
     , cc.cs_soft_total
     , cc.cs_hard_total
     , CASE
           WHEN s.pono IS NULL OR a.pono IS NULL                    THEN NULL
           WHEN cc.cs_commitment_class IS NULL                      THEN 'CS_UNCLASSIFIED'
           WHEN cc.cs_commitment_class = a.audit_commitment_class   THEN 'MATCH'
           ELSE                                                           'DIFF'
       END                                                                  commitment_class_agreement
     , a.audit_currency
     , CASE
           WHEN s.pono IS NULL OR a.pono IS NULL
           THEN NULL
           ELSE ROUND(s.cs_po_total - a.audit_po_total_max, 2)
       END                                                                  delta_to_selected_total
     , CASE
           WHEN a.pono IS NULL                                         THEN NULL
           WHEN a.audit_distinct_total_count > 1                       THEN 'AUDIT_VALUE_VARIABLE'
           ELSE                                                              'AUDIT_VALUE_STABLE'
       END                                                                  audit_total_stability_status
     , CASE
           WHEN s.pono IS NOT NULL AND a.pono IS NULL                    THEN 'CS_ONLY'
           WHEN s.pono IS NULL AND a.pono IS NOT NULL                    THEN 'AUDIT_ONLY'
           WHEN ABS(COALESCE(s.cs_po_total, 0) - COALESCE(a.audit_po_total_max, 0)) <= 0.01
                                                                       THEN 'MATCH_EXACT'
           WHEN ABS(COALESCE(s.cs_po_total, 0) - COALESCE(a.audit_po_total_max, 0)) <= 2.00
                                                                       THEN 'MATCH_ROUNDING'
           ELSE                                                              'VARIANCE_REVIEW'
       END                                                                  reconcile_status
  FROM cte_cs_po                                                             s
       FULL OUTER JOIN
       cte_audit_po                                                          a
          ON a.company = s.company
         AND a.pono    = s.pono
       LEFT OUTER JOIN
       cte_cs_commitment                                                     cc
          ON cc.company = s.company
         AND cc.pono    = s.pono
 ORDER BY present_in, company, pono
;
