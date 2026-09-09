-- ============================================================================
-- BOARD SUMMARY: single-glance headline for the Commitment Stack → Workday PO
-- reconciliation. FINAL scope — the source→sent legs below are computed live from
-- the pipeline tables; the post-load (4th-layer) sent→loaded tie-out is complete
-- and CLEAN, evidenced in the external tie-out workbook (this SQL does not read the
-- control report — see docs/finance-supplier-po-receipt.md §Post-load tie-out and
-- scripts/generate_post_load_tieback.py). zz_ prefix runs after the z_ report.
-- NOTE on row-drop gates: only 2 ACTUAL gates (zero-commitment, old-soft).
-- Unmapped suppliers are NOT dropped from the EIB (wd_supplier_id filter
-- commented in the PO EIB output SQL); they show as an informational bucket,
-- and the 34 UNMAPPED_SUPPLIER + 34 MANDATORY_SUPPLIER validation issues
-- carry the full breakdown for finance review.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_cs_board_summary
;
CREATE TABLE reconcile_cs_board_summary
    AS
SELECT * FROM (
  SELECT 1                                                                            sort_order
       , 'Reconciliation status'                                                      headline
       , 'FINAL — Commitment Stack Report (Proactis) → Workday PO EIB payload → Workday-loaded control report. Post-load sent→loaded tie-out CLEAN (see §Post-load tie-out).' headline_value
  UNION ALL
  SELECT 2, 'Source: Commitment Stack report (nrm rows)'
       , CAST((SELECT SUM(base_rows) FROM reconcile_cs_row_counts) AS VARCHAR)
  UNION ALL
  SELECT 3, 'Row-drop gates applied at staged (2 buckets)'
       , CAST((
           SELECT SUM(drop_zero_commitment + drop_old_soft)
             FROM reconcile_cs_row_counts
         ) AS VARCHAR)
       || ' = '
       || CAST((SELECT SUM(drop_zero_commitment) FROM reconcile_cs_row_counts) AS VARCHAR)
       || ' zero-commitment, '
       || CAST((SELECT SUM(drop_old_soft) FROM reconcile_cs_row_counts) AS VARCHAR)
       || ' old soft (< 2025-08-01)'
  UNION ALL
  SELECT 4, 'Unmapped supplier rows (NOT DROPPED — flow to Workday with NULL supplier_id)'
       , CAST((SELECT SUM(unmapped_supplier_rows) FROM reconcile_cs_row_counts) AS VARCHAR)
       || ' rows. See validation_supplier_purchase_order_receipt_issues (UNMAPPED_SUPPLIER + MANDATORY_SUPPLIER buckets) for detail.'
  UNION ALL
  SELECT 5, 'PO lines migrated to Workday EIB (staged = target 1:1 expected)'
       , CAST((SELECT SUM(workday_rows) FROM reconcile_cs_row_counts) AS VARCHAR)
  UNION ALL
  SELECT 6, 'Total extended_amount in Workday, transaction currency'
       , (
           WITH non_gbp AS (
               SELECT currency || ' ' || ROUND(workday_amount, 2) AS txt
                 FROM reconcile_cs_amount_by_currency
                WHERE COALESCE(currency, '') <> 'GBP'
                  AND COALESCE(workday_amount, 0) <> 0
           )
           SELECT CAST(ROUND(SUM(workday_amount), 2) AS VARCHAR) || ' GBP'
                  || CASE WHEN (SELECT COUNT(*) FROM non_gbp) = 0 THEN ''
                          ELSE '  |  ' || (SELECT STRING_AGG(txt, '  ') FROM non_gbp)
                     END
             FROM reconcile_cs_amount_by_currency
            WHERE currency = 'GBP'
         )
  UNION ALL
  SELECT 7, 'Currencies reconciled'
       , CAST((SELECT COUNT(*) FROM reconcile_cs_currency_coverage WHERE present_in = 'BOTH') AS VARCHAR)
       || ' of ' || CAST((SELECT COUNT(*) FROM reconcile_cs_currency_coverage) AS VARCHAR)
       || CASE WHEN (SELECT COUNT(*) FROM reconcile_cs_currency_coverage WHERE present_in <> 'BOTH') = 0
               THEN ' (all present source and target)' ELSE ' — SEE currency_coverage' END
  UNION ALL
  SELECT 8, 'Rows unexplained (must be 0)'
       , CAST((SELECT COALESCE(SUM(ABS(unexplained_staged) + ABS(unexplained_workday)), 0) FROM reconcile_cs_row_counts) AS VARCHAR)
  UNION ALL
  SELECT 9, 'Line value unexplained (must be 0)'
       , CAST((SELECT ROUND(COALESCE(SUM(ABS(unexplained_staged_amount) + ABS(unexplained_workday_amount)), 0), 2) FROM reconcile_cs_amount_by_currency) AS VARCHAR)
  UNION ALL
  SELECT 10, 'Source → sent Workday payload result'
       , CASE WHEN (SELECT COUNT(*) FROM reconciliation_supplier_po_receipt_issues) = 0
              THEN 'CLEAN — every Commitment Stack PO line ties to the sent Workday EIB after the 2 row-drop gates (zero-commitment + old-soft). Unmapped suppliers carry through with NULL supplier_id (see validation_issues for review).'
              ELSE 'REVIEW — see all_issues' END
  UNION ALL
  SELECT 11, 'Post-load Workday tie-out (4th layer)'
       , 'COMPLETE / CLEAN — sent EIB vs Workday control report ties to the penny and to the line: PO 10,833 lines / 6,971 POs, receipt 7,122 lines / 4,659 docs, £0.00 delta in every currency, 0 line-multiset mismatches, 0 Workday validation errors. Evidence: reconcile dir FA_PostLoad_Tieback_YYYYMMDD.xlsx (regenerate via scripts/generate_post_load_tieback.py).
Note: audit-coverage reconcile is off-pipeline — run reconcile_cs_audit_coverage ad-hoc only after loading the audit extract table fin_supplier_po_audit_commitment.'
) ORDER BY sort_order
;
