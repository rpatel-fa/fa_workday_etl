-- ============================================================================
-- RECONCILE: Customer email proof (row-level)
-- One row per distinct valid (customer_id, email_address) pair aggregated
-- from nrm_agg_email_to/cc/bcc_address (src_fin_customer), showing whether
-- that email is present under the SAME customer_id in workday_customer_email.
-- A found='N' row here is expected for emails lost to the cross-customer
-- dedup in workday_customer_email.sql (RANK() PARTITION BY email_address) -
-- see reconciliation_customer_issues (CROSS_ID_DEDUP) for which other
-- customer_id retained it.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_customer_email_proof
;
CREATE TABLE reconcile_customer_email_proof
    AS
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
SELECT se.customer_id                                                                           customer_id
     , se.email_address                                                                         email_address
     , CASE WHEN w.email_address IS NOT NULL THEN 'Y' ELSE 'N' END                              found_in_workday_customer_email
  FROM cte_source_emails                                                                        se
  LEFT JOIN workday_customer_email                                                               w
         ON w.customer_id                                                                       = se.customer_id
        AND w.email_address                                                                     = se.email_address
 ORDER BY found_in_workday_customer_email
     , customer_id
;
