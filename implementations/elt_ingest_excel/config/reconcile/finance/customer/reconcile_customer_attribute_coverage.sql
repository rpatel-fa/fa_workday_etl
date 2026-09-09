-- ============================================================================
-- RECONCILE: Customer attribute coverage
-- Confirms that address and email details aggregated across merged duplicate
-- records actually made it through to the workday_customer_* tables.
-- Note: workday_customer_settlement_account.sql is currently disabled
-- (hardcoded `AND 1 = 2`), so there is no bank coverage check for customers.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- COVERAGE_COUNT: distinct non-blank addresses aggregated per customer must
-- all appear in workday_customer_address_base (pre country-specific filter,
-- so this isolates aggregation loss from the deliberate HU exclusion).
-- ----------------------------------------------------------------------------
INSERT INTO reconciliation_customer_issues
  WITH cte_source_address_count
    AS (
SELECT c.customer_id                                                                            customer_id
     , COUNT(DISTINCT u.address)                                                                source_address_count
  FROM src_fin_customer                                                                         c
     , UNNEST(c.nrm_array_customer_address)                                                     u(address)
 WHERE COALESCE(
           NULLIF(TRIM(u.address.nrm_address_line_1), '')
         , NULLIF(TRIM(u.address.nrm_address_line_2), '')
         , NULLIF(TRIM(u.address.nrm_address_line_3), '')
         , NULLIF(TRIM(u.address.nrm_address_line_4), '')
       )                                                                                        IS NOT NULL
 GROUP BY c.customer_id
       )
     , cte_final_address_count
    AS (
SELECT customer_id
     , COUNT(*)                                                                                 final_address_count
  FROM workday_customer_address_base
 GROUP BY customer_id
       )
SELECT
       'workday_customer_address_base'                                                          table_name
     , 'COVERAGE_COUNT'                                                                         rule_type
     , 'address'                                                                                col
     , c.customer_id                                                                            customer_id
     , c.source_address_count || ' source vs ' || COALESCE(f.final_address_count, 0) || ' final' detail
     , 'Distinct addresses aggregated from merged duplicate records do not all appear in workday_customer_address_base.' message
  FROM cte_source_address_count                                                                 c
  LEFT JOIN cte_final_address_count                                                             f
         ON f.customer_id                                                                       = c.customer_id
 WHERE c.source_address_count > COALESCE(f.final_address_count, 0)
;

-- ----------------------------------------------------------------------------
-- CROSS_ID_DEDUP: workday_customer_email keeps each email address under only
-- one customer_id (RANK() PARTITION BY email_address in workday_customer_email.sql).
-- If two customers share an email, the losing customer's copy disappears from
-- the final table entirely. This flags every customer where that happened.
-- ----------------------------------------------------------------------------
INSERT INTO reconciliation_customer_issues
  WITH cte_source_emails
    AS (
SELECT DISTINCT
       c.customer_id                                                                            customer_id
     , TRIM(LOWER(u.email))                                                                     email_address
  FROM src_fin_customer                                                                         c
     , UNNEST(STRING_SPLIT(
           COALESCE(c.nrm_agg_email_to_address, '')  || ';' ||
           COALESCE(c.nrm_agg_email_cc_address, '')  || ';' ||
           COALESCE(c.nrm_agg_email_bcc_address, '')
       , ';'))                                                                                  u(email)
 WHERE NULLIF(TRIM(u.email), '')                                                                IS NOT NULL
   AND TRIM(u.email)                                                                            ~ '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'
       )
SELECT
       'workday_customer_email'                                                                 table_name
     , 'CROSS_ID_DEDUP'                                                                         rule_type
     , 'email_address'                                                                          col
     , se.customer_id                                                                           customer_id
     , se.email_address                                                                         detail
     , 'Valid source email is not present under this customer_id in workday_customer_email - likely lost to the cross-customer dedup by email_address, or dropped outright.' message
  FROM cte_source_emails                                                                        se
  LEFT JOIN workday_customer_email                                                               w
         ON w.customer_id                                                                       = se.customer_id
        AND w.email_address                                                                     = se.email_address
 WHERE w.email_address IS NULL
;
