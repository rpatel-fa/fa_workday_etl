-- ============================================================================
-- RECONCILE: Supplier tax proof (row-level)
-- workday_supplier_tax.sql is a direct 1-row-per-supplier_id passthrough
-- from src_fin_supplier (no UNNEST, no filter), so coverage is guaranteed by
-- construction. This table proves it explicitly: one row per supplier_id in
-- src_fin_supplier, confirming exactly one corresponding row exists in
-- workday_supplier_tax. Whether the derived tax_id_type/transaction_tax_status
-- VALUES are correct is a data-quality concern handled by
-- validate_supplier_tax.sql, not a reconciliation/coverage concern.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_supplier_tax_proof
;
CREATE TABLE reconcile_supplier_tax_proof
    AS
SELECT s.supplier_id                                                                             supplier_id
     , CASE WHEN w.supplier_id IS NOT NULL THEN 'Y' ELSE 'N' END                                 found_in_workday_supplier_tax
     , w.tax_id                                                                                  tax_id
  FROM src_fin_supplier                                                                          s
  LEFT JOIN workday_supplier_tax                                                                 w
         ON w.supplier_id                                                                        = s.supplier_id
 ORDER BY found_in_workday_supplier_tax
     , supplier_id
;
