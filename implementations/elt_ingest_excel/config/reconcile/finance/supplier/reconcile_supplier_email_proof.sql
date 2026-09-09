-- ============================================================================
-- RECONCILE: Supplier email proof (row-level)
-- One row per distinct valid (supplier_id, email_address) pair aggregated
-- from nrm_agg_email_to/cc/bcc_address (src_fin_supplier), showing whether
-- that email is present under the SAME supplier_id in workday_supplier_email.
-- A found='N' row here is expected for emails lost to the cross-supplier
-- dedup in workday_supplier_email.sql (RANK() PARTITION BY email_address) -
-- see reconciliation_supplier_issues (CROSS_ID_DEDUP) for which other
-- supplier_id retained it.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_supplier_email_proof
;
CREATE TABLE reconcile_supplier_email_proof
    AS
  WITH cte_source_emails
    AS (
SELECT DISTINCT
       s.supplier_id                                                                            supplier_id
     , TRIM(LOWER(u.email))                                                                     email_address
  FROM src_fin_supplier                                                                         s
     , UNNEST(STRING_SPLIT(
           COALESCE(s.nrm_agg_email_to_address, '')  || ';' ||
           COALESCE(s.nrm_agg_email_cc_address, '')  || ';' ||
           COALESCE(s.nrm_agg_email_bcc_address, '')
       , ';'))                                                                                  u(email)
 WHERE NULLIF(TRIM(u.email), '')                                                                IS NOT NULL
   AND TRIM(u.email)                                                                            ~ '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'
       )
SELECT se.supplier_id                                                                           supplier_id
     , se.email_address                                                                         email_address
     , CASE WHEN w.email_address IS NOT NULL THEN 'Y' ELSE 'N' END                              found_in_workday_supplier_email
  FROM cte_source_emails                                                                        se
  LEFT JOIN workday_supplier_email                                                               w
         ON w.supplier_id                                                                       = se.supplier_id
        AND w.email_address                                                                     = se.email_address
 ORDER BY found_in_workday_supplier_email
     , supplier_id
;
