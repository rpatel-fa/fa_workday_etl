-- ============================================================================
-- RECONCILE: Customer payment proof (row-level)
-- workday_customer_payment.sql is a direct 1-row-per-customer_id passthrough
-- from src_fin_customer (no UNNEST, no filter), so coverage is guaranteed by
-- construction. This table proves it explicitly: one row per customer_id in
-- src_fin_customer, confirming exactly one corresponding row exists in
-- workday_customer_payment. Whether the mapped payment_terms VALUE is
-- correct (e.g. ref_source_payment_terms produced a real mapping rather than
-- NULL) is a data-quality concern handled by validate_customer_payment.sql,
-- not a reconciliation/coverage concern.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_customer_payment_proof
;
CREATE TABLE reconcile_customer_payment_proof
    AS
SELECT c.customer_id                                                                            customer_id
     , CASE WHEN w.customer_id IS NOT NULL THEN 'Y' ELSE 'N' END                                found_in_workday_customer_payment
     , w.payment_terms                                                                          payment_terms
  FROM src_fin_customer                                                                         c
  LEFT JOIN workday_customer_payment                                                             w
         ON w.customer_id                                                                       = c.customer_id
 ORDER BY found_in_workday_customer_payment
     , customer_id
;
