-- ============================================================================
-- RECONCILE: Supplier phone/fax coverage confirmation
-- workday_supplier_phone.sql now reads nrm_agg_phone_1/2/3 and nrm_agg_fax
-- (the aggregated columns) directly, matching workday_customer_phone.sql -
-- previously it read the scalar phone_number_1/2/3/fax_number columns
-- (single winning row only), which meant phone/fax numbers captured from
-- merged duplicate records never reached the published table (see git
-- history for workday_supplier_phone.sql for the prior behaviour).
--
-- This check confirms parity holds all the way to the final published
-- table: every distinct phone/fax number captured in the aggregate should
-- have a corresponding row in workday_supplier_phone. A shortfall here would
-- point to the phone-parsing/validation step (length >= 7, unparseable
-- format, country-code guessing) silently dropping a number, not the
-- aggregation or column-choice.
-- ============================================================================

INSERT INTO reconciliation_supplier_issues
  WITH cte_agg_count
    AS (
SELECT s.supplier_id                                                                            supplier_id
     , COUNT(DISTINCT u.phone)                                                                  agg_phone_count
  FROM src_fin_supplier                                                                         s
     , UNNEST(STRING_SPLIT(
           COALESCE(s.nrm_agg_phone_1, '') || ';' ||
           COALESCE(s.nrm_agg_phone_2, '') || ';' ||
           COALESCE(s.nrm_agg_phone_3, '') || ';' ||
           COALESCE(s.nrm_agg_fax,     '')
       , ';'))                                                                                  u(phone)
 WHERE NULLIF(TRIM(u.phone), '')                                                                IS NOT NULL
 GROUP BY s.supplier_id
       )
     , cte_final_count
    AS (
SELECT supplier_id
     , COUNT(*)                                                                                 final_phone_count
  FROM workday_supplier_phone
 GROUP BY supplier_id
       )
SELECT
       'workday_supplier_phone'                                                                 table_name
     , 'COVERAGE_COUNT'                                                                         rule_type
     , 'phone_number'                                                                           col
     , a.supplier_id                                                                            supplier_id
     , a.agg_phone_count || ' aggregated in src_fin_supplier.nrm_agg_phone/fax vs ' ||
       COALESCE(f.final_phone_count, 0) || ' rows in workday_supplier_phone'                    detail
     , 'Distinct phone/fax numbers aggregated from merged duplicate records do not all appear in workday_supplier_phone - likely dropped by phone parsing/validation (e.g. length < 7, unparseable format), not by the aggregation step.' message
  FROM cte_agg_count                                                                            a
  LEFT JOIN cte_final_count                                                                     f
         ON f.supplier_id                                                                       = a.supplier_id
 WHERE a.agg_phone_count > COALESCE(f.final_phone_count, 0)
;
