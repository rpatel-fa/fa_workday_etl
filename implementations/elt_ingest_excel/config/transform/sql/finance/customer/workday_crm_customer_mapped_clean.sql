DROP TABLE IF EXISTS workday_crm_customer_mapped_clean
;
CREATE TABLE workday_crm_customer_mapped_clean
    AS
-- One row per CRM customer. Where a CRM record maps to several Workday customers (created
-- against different shared CRM licence numbers), the earliest Workday customer survives and
-- its reference id becomes the combined list of every licence number held across the matches.
WITH cte_grp
  AS (
SELECT m.crm_customer_type
     , m.crm_customer_id
     , MIN(m.wd_customer_id)                                       wd_customer_id
     , COUNT(DISTINCT m.wd_customer_id)                            wd_customer_count
     , STRING_AGG(DISTINCT m.wd_customer_id, '|' ORDER BY m.wd_customer_id) wd_customer_ids
  FROM workday_crm_customer_mapped      m
 GROUP BY m.crm_customer_type
        , m.crm_customer_id
     )
     , cte_ref
  AS (
SELECT m.crm_customer_type
     , m.crm_customer_id
     , MIN(m.wd_customer_id)                                       ref_order
     , TRIM(UNNEST(STRING_SPLIT(m.wd_reference_id, '|')))          wd_unnest_reference_id
  FROM workday_crm_customer_mapped      m
 GROUP BY m.crm_customer_type
        , m.crm_customer_id
        , m.wd_reference_id
     )
     , cte_ref_agg
  AS (
SELECT r.crm_customer_type
     , r.crm_customer_id
     , STRING_AGG(r.wd_unnest_reference_id, '|' ORDER BY r.ref_order, r.wd_unnest_reference_id) wd_reference_id_combined
  FROM (
SELECT crm_customer_type, crm_customer_id, wd_unnest_reference_id, MIN(ref_order) ref_order
  FROM cte_ref
 GROUP BY crm_customer_type, crm_customer_id, wd_unnest_reference_id
       )                                r
 GROUP BY r.crm_customer_type
        , r.crm_customer_id
     )
SELECT DISTINCT
       m.crm_customer_type
     , m.crm_customer_id
     , m.crm_licenceref_no
     , m.crm_fan
     , g.wd_customer_id
     , m.wd_reference_id                                           wd_reference_id_original
     , a.wd_reference_id_combined                                  wd_reference_id
     , g.wd_customer_count
     , g.wd_customer_ids
  FROM workday_crm_customer_mapped      m
       INNER JOIN
       cte_grp                          g
         ON  g.crm_customer_type = m.crm_customer_type
         AND g.crm_customer_id   = m.crm_customer_id
         AND g.wd_customer_id    = m.wd_customer_id
       INNER JOIN
       cte_ref_agg                      a
         ON  a.crm_customer_type = m.crm_customer_type
         AND a.crm_customer_id   = m.crm_customer_id
;
