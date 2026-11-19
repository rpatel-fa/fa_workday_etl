DROP TABLE IF EXISTS workday_crm_customer_mapped_wd
;
CREATE TABLE workday_crm_customer_mapped_wd
    AS
-- One row per Workday customer. Where several CRM records map to the same Workday customer,
-- the primary is chosen by: 1) Active CRM record before Inactive, 2) ACCOUNT before INDIVIDUAL,
-- 3) latest CRM created_on, 4) latest CRM modified_on. crm_customer_id is only a fallback so a full
-- tie still gives the same result every run.
-- The other CRM records stay listed in crm_secondary_customers.
WITH cte_rnk
  AS (
SELECT c.*
     , LEN(LIST_INTERSECT(
              LIST_TRANSFORM(STRING_SPLIT(c.crm_licenceref_no, '|'), x -> TRIM(x))
            , LIST_TRANSFORM(STRING_SPLIT(c.wd_reference_id,   '|'), x -> TRIM(x))
       ))                                                          crm_licence_overlap
  FROM workday_crm_customer_mapped_clean    c
     )
     , cte_ord
  AS (
SELECT r.*
     , ROW_NUMBER() OVER (
         PARTITION BY r.wd_customer_id
             ORDER BY r.crm_is_active DESC
                    , (r.crm_customer_type = 'ACCOUNT') DESC
                    , r.crm_created_on DESC NULLS LAST
                    , r.crm_modified_on DESC NULLS LAST
                    , r.crm_customer_id
       )                                                           crm_rank
     , COUNT(*) OVER (PARTITION BY r.wd_customer_id)               crm_customer_count
  FROM cte_rnk                              r
     )
SELECT p.wd_customer_id
     , p.wd_reference_id
     , p.crm_customer_type
     , p.crm_customer_id
     , p.crm_licenceref_no
     , p.crm_fan
     , p.crm_licence_overlap
     , p.crm_customer_count
     , (SELECT STRING_AGG(s.crm_customer_type || ':' || s.crm_customer_id, '|'
                          ORDER BY s.crm_rank)
          FROM cte_ord s
         WHERE s.wd_customer_id = p.wd_customer_id
           AND s.crm_rank > 1)                                     crm_secondary_customers
  FROM cte_ord                              p
 WHERE p.crm_rank = 1
;
