-- ============================================================================
-- RECONCILE: Supplier payment proof (row-level)
-- workday_supplier_payment.sql is a 1-row-per-supplier_id passthrough from
-- src_fin_supplier. It has a `WHERE t.nrm_payment_terms_id IS NOT NULL`
-- clause, but src_fin_supplier itself is already filtered to only contain
-- suppliers with a non-null nrm_payment_terms_id (see src_fin_supplier.sql),
-- so that WHERE clause is redundant in practice - coverage is guaranteed by
-- construction. This table proves it explicitly: one row per supplier_id in
-- src_fin_supplier, confirming exactly one corresponding row exists in
-- workday_supplier_payment. Whether the mapped payment_terms VALUE is
-- correct is a data-quality concern handled by validate_supplier_payment.sql,
-- not a reconciliation/coverage concern.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_supplier_payment_proof
;
CREATE TABLE reconcile_supplier_payment_proof
    AS
SELECT s.supplier_id                                                                             supplier_id
     , CASE WHEN w.supplier_id IS NOT NULL THEN 'Y' ELSE 'N' END                                 found_in_workday_supplier_payment
     , w.payment_terms                                                                          payment_terms
  FROM src_fin_supplier                                                                          s
  LEFT JOIN workday_supplier_payment                                                             w
         ON w.supplier_id                                                                        = s.supplier_id
 ORDER BY found_in_workday_supplier_payment
     , supplier_id
;
