-- ============================================================================
-- RECONCILE: Customer phone/fax coverage confirmation
-- workday_customer_phone.sql reads nrm_agg_phone_1/2/3 and nrm_agg_fax (the
-- aggregated columns) directly - unlike supplier, which has a known gap here
-- (config/reconcile/finance/supplier/reconcile_supplier_phone_gap.sql).
-- This check confirms parity holds all the way to the final published table:
-- every distinct phone/fax number captured in the aggregate should have a
-- corresponding row in workday_customer_phone. A shortfall here would point
-- to the phone-parsing/validation step (length >= 7, unparseable format,
-- country-code guessing) silently dropping a number, not the aggregation.
-- ============================================================================

INSERT INTO reconciliation_customer_issues
  WITH cte_agg_count
    AS (
SELECT c.customer_id                                                                            customer_id
     , COUNT(DISTINCT u.phone)                                                                  agg_phone_count
  FROM src_fin_customer                                                                         c
     , UNNEST(STRING_SPLIT(
           COALESCE(c.nrm_agg_phone_1, '') || ';' ||
           COALESCE(c.nrm_agg_phone_2, '') || ';' ||
           COALESCE(c.nrm_agg_phone_3, '') || ';' ||
           COALESCE(c.nrm_agg_fax,     '')
       , ';'))                                                                                  u(phone)
 WHERE NULLIF(TRIM(u.phone), '')                                                                IS NOT NULL
 GROUP BY c.customer_id
       )
     , cte_final_count
    AS (
SELECT customer_id
     , COUNT(*)                                                                                 final_phone_count
  FROM workday_customer_phone
 GROUP BY customer_id
       )
SELECT
       'workday_customer_phone'                                                                 table_name
     , 'COVERAGE_COUNT'                                                                         rule_type
     , 'phone_number'                                                                           col
     , a.customer_id                                                                            customer_id
     , a.agg_phone_count || ' aggregated in src_fin_customer.nrm_agg_phone/fax vs ' ||
       COALESCE(f.final_phone_count, 0) || ' rows in workday_customer_phone'                    detail
     , 'Distinct phone/fax numbers aggregated from merged duplicate records do not all appear in workday_customer_phone - likely dropped by phone parsing/validation (e.g. length < 7, unparseable format), not by the aggregation step.' message
  FROM cte_agg_count                                                                            a
  LEFT JOIN cte_final_count                                                                     f
         ON f.customer_id                                                                       = a.customer_id
 WHERE a.agg_phone_count > COALESCE(f.final_phone_count, 0)
;
