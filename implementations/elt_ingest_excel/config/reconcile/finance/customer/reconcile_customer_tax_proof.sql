-- ============================================================================
-- RECONCILE: Customer tax proof (row-level)
-- workday_customer_tax.sql is a direct 1-row-per-customer_id passthrough
-- from src_fin_customer (no UNNEST, no filter), so coverage is guaranteed by
-- construction. This table proves it explicitly: one row per customer_id in
-- src_fin_customer, confirming exactly one corresponding row exists in
-- workday_customer_tax. Whether the derived tax_id_type/transaction_tax_status
-- VALUES are correct is a data-quality concern handled by
-- validate_customer_tax.sql, not a reconciliation/coverage concern.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_customer_tax_proof
;
CREATE TABLE reconcile_customer_tax_proof
    AS
SELECT c.customer_id                                                                            customer_id
     , CASE WHEN w.customer_id IS NOT NULL THEN 'Y' ELSE 'N' END                                found_in_workday_customer_tax
     , w.tax_id                                                                                 tax_id
  FROM src_fin_customer                                                                         c
  LEFT JOIN workday_customer_tax                                                                 w
         ON w.customer_id                                                                       = c.customer_id
 ORDER BY found_in_workday_customer_tax
     , customer_id
;
