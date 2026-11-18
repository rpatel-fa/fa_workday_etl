DROP TABLE IF EXISTS workday_crm_customer_mapped_wd_ignored
;
CREATE TABLE workday_crm_customer_mapped_wd_ignored
    AS
-- One row per Workday customer (fin_workday_extract_customer) that gets no CRM id, i.e. has no row
-- in workday_crm_customer_mapped_wd, with the reason why. Matched on licence number (the rule):
--   DUPLICATE                       - _clean sent its CRM record to an earlier Workday customer
--                                     (kept_wd_customer_id), which now holds that CRM id
--   CRM_LOST_TO_HIGHER_RANK         - _clean sent its CRM record to kept_wd_customer_id, but that
--                                     customer took a higher-ranked CRM record (kept_crm_*) in _wd,
--                                     so the matched CRM record is set on no Workday customer
-- Not matched on licence number (Workday reference_id should hold the CRM licence number):
--   WD_REF_BLANK                    - Workday reference_id is blank
--   WD_REF_IS_CRM_ID                - reference_id holds a CRM FAN / Account ID instead of a licence
--                                     number, and that CRM record has a licence number (Workday fix)
--   WD_REF_IS_CRM_ID_NO_CRM_LICENCE - as above, but the CRM record has no licence number either
--                                     (Workday and CRM fix)
--   NOT_IN_CRM                      - reference_id matches no CRM licence number, FAN or Account ID
-- matched_crm_customer_type / _id are the CRM records matched by licence number, or for the
-- WD_REF_IS_CRM_ID* reasons, the CRM records the reference_id points at (FAN / Account ID).
-- matched_crm_licenceref_no (WD_REF_IS_CRM_ID* only) is each matched CRM record's licence number,
-- i.e. the value Workday's reference_id should hold; blank where the record has none (the 3d fix).
-- Several matches are '|'-separated, in the same order in every matched_ column.
WITH cte_ignored
  AS (
SELECT w.customer_id                                               wd_customer_id
     , w.reference_id                                              wd_reference_id
  FROM fin_workday_extract_customer         w
 WHERE w.customer_id NOT IN (SELECT wd_customer_id FROM workday_crm_customer_mapped_wd)
     )
-- Matched on licence number, but dropped by _clean / _wd
     , cte_lic
  AS (
SELECT DISTINCT
       m.wd_customer_id
     , m.crm_customer_type
     , m.crm_customer_id
  FROM workday_crm_customer_mapped          m
 WHERE m.wd_customer_id IN (SELECT wd_customer_id FROM cte_ignored)
     )
     , cte_lic_kept
  AS (
SELECT l.*
     , c.wd_customer_id                                            kept_wd_customer_id
     , w.crm_customer_type                                         kept_crm_customer_type
     , w.crm_customer_id                                           kept_crm_customer_id
     , (    w.crm_customer_type = l.crm_customer_type
        AND w.crm_customer_id   = l.crm_customer_id)               crm_kept
  FROM cte_lic                              l
       INNER JOIN
       workday_crm_customer_mapped_clean    c
         ON  c.crm_customer_type = l.crm_customer_type
         AND c.crm_customer_id   = l.crm_customer_id
       INNER JOIN
       workday_crm_customer_mapped_wd       w
         ON  w.wd_customer_id    = c.wd_customer_id
     )
     , cte_lic_agg
  AS (
SELECT k.wd_customer_id
     , STRING_AGG(k.crm_customer_type, '|' ORDER BY k.crm_customer_type, k.crm_customer_id) matched_crm_customer_type
     , STRING_AGG(k.crm_customer_id,   '|' ORDER BY k.crm_customer_type, k.crm_customer_id) matched_crm_customer_id
     , STRING_AGG(DISTINCT k.kept_wd_customer_id, '|' ORDER BY k.kept_wd_customer_id) kept_wd_customer_id
     , STRING_AGG(DISTINCT k.kept_crm_customer_type, '|' ORDER BY k.kept_crm_customer_type) kept_crm_customer_type
     , STRING_AGG(DISTINCT k.kept_crm_customer_id, '|' ORDER BY k.kept_crm_customer_id) kept_crm_customer_id
     , CASE WHEN BOOL_OR(k.crm_kept) THEN 'DUPLICATE'
            ELSE 'CRM_LOST_TO_HIGHER_RANK'
       END                                                         reason
  FROM cte_lic_kept                         k
 GROUP BY k.wd_customer_id
     )
-- Not matched on licence number: does reference_id hold a CRM FAN / Account ID instead?
     , cte_crm_id
  AS (
SELECT 'ACCOUNT'                                                   crm_customer_type
     , UPPER(TRIM(CAST(ca.account_number_id AS VARCHAR)))          crm_customer_id
     , NULLIF(NULLIF(TRIM(ca.licencerefno), ''), 'NULL')           crm_licenceref_no
  FROM crm_account                          ca
UNION ALL
SELECT 'INDIVIDUAL'
     , UPPER(TRIM(CAST(cc.fan AS VARCHAR)))
     , NULLIF(NULLIF(TRIM(cc.licencesref), ''), 'NULL')
  FROM crm_contact                          cc
     )
     , cte_ref
  AS (
SELECT i.wd_customer_id
     , UPPER(TRIM(UNNEST(STRING_SPLIT(i.wd_reference_id, '|'))))  wd_unnest_reference_id
  FROM cte_ignored                          i
 WHERE i.wd_customer_id NOT IN (SELECT wd_customer_id FROM cte_lic_agg)
     )
     , cte_ref_match
  AS (
SELECT DISTINCT
       r.wd_customer_id
     , x.crm_customer_type
     , x.crm_customer_id
     , x.crm_licenceref_no
  FROM cte_ref                              r
       INNER JOIN
       cte_crm_id                           x
         ON  x.crm_customer_id = r.wd_unnest_reference_id
 WHERE r.wd_unnest_reference_id <> ''
     )
     , cte_ref_agg
  AS (
SELECT x.wd_customer_id
     , STRING_AGG(x.crm_customer_type, '|' ORDER BY x.crm_customer_type, x.crm_customer_id) matched_crm_customer_type
     , STRING_AGG(x.crm_customer_id,   '|' ORDER BY x.crm_customer_type, x.crm_customer_id) matched_crm_customer_id
     , STRING_AGG(COALESCE(x.crm_licenceref_no, ''), '|'
                  ORDER BY x.crm_customer_type, x.crm_customer_id) matched_crm_licenceref_no
     , CASE WHEN BOOL_OR(x.crm_licenceref_no IS NOT NULL) THEN 'WD_REF_IS_CRM_ID'
            ELSE 'WD_REF_IS_CRM_ID_NO_CRM_LICENCE'
       END                                                         reason
  FROM cte_ref_match                        x
 GROUP BY x.wd_customer_id
     )
SELECT i.wd_customer_id
     , i.wd_reference_id
     , COALESCE(l.matched_crm_customer_type, r.matched_crm_customer_type) matched_crm_customer_type
     , COALESCE(l.matched_crm_customer_id,   r.matched_crm_customer_id)   matched_crm_customer_id
     , r.matched_crm_licenceref_no
     , l.kept_wd_customer_id
     , l.kept_crm_customer_type
     , l.kept_crm_customer_id
     , CASE WHEN l.reason IS NOT NULL                                         THEN l.reason
            WHEN NULLIF(NULLIF(TRIM(i.wd_reference_id), ''), 'NULL') IS NULL THEN 'WD_REF_BLANK'
            WHEN r.reason IS NOT NULL                                         THEN r.reason
            ELSE 'NOT_IN_CRM'
       END                                                         reason
  FROM cte_ignored                          i
       LEFT OUTER JOIN
       cte_lic_agg                          l
         ON l.wd_customer_id = i.wd_customer_id
       LEFT OUTER JOIN
       cte_ref_agg                          r
         ON r.wd_customer_id = i.wd_customer_id
 ORDER BY i.wd_customer_id
;
