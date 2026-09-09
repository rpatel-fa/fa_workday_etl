-- ============================================================================
-- OFF-PIPELINE — DO NOT ADD TO order.txt WITHOUT INGEST CONFIG CHANGE FIRST.
-- RECONCILE: AUDIT_ONLY triage — residual POs in the audit but not in the
-- Commitment Stack. The audit is a commitments superset by design, so most
-- AUDIT_ONLY POs are out-of-scope (closed/fully-invoiced/old-soft cutoffs).
-- Classifies each PO into a suggested triage_bucket using invoice/delivery
-- status + commitment class from the audit.
--
-- DEPENDENCY NOT IN THIS MODULE: references fin_supplier_po_audit_commitment,
-- which is (as at 08.08.2026) NEITHER declared in this module's ingest config
-- (elt_ingest_excel/config/ingest/finance/supplier_purchase_order_receipt.json)
-- NOR produced by the transform SQL folder. The audit extract must be loaded
-- externally before this file can run. Run ad-hoc only.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_audit_only_triage
;
CREATE TABLE reconcile_cs_audit_only_triage
    AS
  WITH cte_scope_entities
    AS (
SELECT 'FA'   company
UNION ALL
SELECT 'NFC'
UNION ALL
SELECT 'WNSL'
       )
     , cte_audit_po
    AS (
SELECT a.company
     , a.orderno                                                             pono
     , COUNT(*)                                                             audit_row_count
     , MAX(a.companyname)                                                   audit_companyname
     , MAX(NULLIF(TRIM(a.invoicestatusdescription), ''))                   invoice_status
     , MAX(NULLIF(TRIM(a.deliverystatusdescription), ''))                  delivery_status
     , ROUND(MAX(TRY_CAST(NULLIF(TRIM(a.ordernetvalue), '') AS DOUBLE)), 2) audit_po_total
     , ROUND(SUM(TRY_CAST(NULLIF(TRIM(a.soft_commitment), '') AS DOUBLE)), 2)  audit_soft_total
     , ROUND(SUM(TRY_CAST(NULLIF(TRIM(a.hard_commitment), '') AS DOUBLE)), 2)  audit_hard_total
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
     , MAX(NULLIF(TRIM(a.commitmentdate), ''))                             max_commitmentdate
     , MAX(NULLIF(TRIM(a.orderdate), ''))                                  orderdate
     , MAX(NULLIF(TRIM(a.invoicenumber), ''))                              invoicenumber
  FROM fin_supplier_po_audit_commitment                                      a
       INNER JOIN
       cte_scope_entities                                                   e
          ON e.company = a.company
 WHERE NULLIF(TRIM(a.orderno), '') IS NOT NULL
 GROUP BY a.company, a.orderno
       )
     , cte_cs_po
    AS (
SELECT DISTINCT b.company_shortcode company, b.order_number pono
  FROM src_fin_supplier_po_commitment_stack_base                            b
 WHERE NULLIF(TRIM(b.order_number), '') IS NOT NULL
       )
SELECT a.company
     , a.pono
     , a.audit_companyname
     , a.audit_commitment_class
     , a.invoice_status
     , a.delivery_status
     , a.audit_po_total
     , a.audit_soft_total
     , a.audit_hard_total
     , a.orderdate
     , a.max_commitmentdate
     , a.invoicenumber
     , a.audit_row_count
     , CASE
           WHEN a.audit_commitment_class = 'HARD'
            AND a.delivery_status = 'Fully Delivered'
            AND UPPER(COALESCE(a.invoice_status, '')) LIKE '%PART%INVOICED%'
           THEN 'HARD_ACCRUAL_REVIEW'
           WHEN a.audit_commitment_class = 'HARD'
            AND UPPER(COALESCE(a.invoice_status, '')) IN ('FULLY INVOICED', 'FULLY INVOICED ')
           THEN 'LIKELY_CLOSED_INVOICED'
           WHEN a.audit_commitment_class = 'SOFT'
           THEN 'EXPECTED_OLD_SOFT'
           WHEN a.audit_commitment_class = 'MIXED'
            AND UPPER(COALESCE(a.invoice_status, '')) LIKE '%INVOICED%'
           THEN 'MIXED_PARTIAL_INVOICE_REVIEW'
           ELSE
                'OTHER_AUDIT_ONLY_REVIEW'
       END                                                                  triage_bucket
  FROM cte_audit_po                                                          a
       LEFT OUTER JOIN
       cte_cs_po                                                             s
          ON s.company = a.company
         AND s.pono    = a.pono
 WHERE s.pono IS NULL
 ORDER BY a.audit_commitment_class
        , a.company
        , a.pono
;
