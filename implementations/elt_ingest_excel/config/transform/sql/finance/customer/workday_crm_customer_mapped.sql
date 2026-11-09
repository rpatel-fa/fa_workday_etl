DROP TABLE IF EXISTS workday_crm_customer_mapped
;
CREATE TABLE workday_crm_customer_mapped
    AS
WITH cte_wd 
  AS (
SELECT w.customer_id                                    wd_customer_id
     , w.reference_id                                   wd_reference_id
     , TRIM(UNNEST(STRING_SPLIT(w.reference_id, '|')))  wd_unnest_reference_id
  FROM fin_workday_extract_customer     w
     )
     , cte_crm_wd
  AS (
SELECT
       crm.* EXCLUDE (crm_unnest_reference_id)
     , w.wd_customer_id
     , w.wd_reference_id
  FROM src_crm_customer_base            crm
       INNER JOIN
       cte_wd                           w
         ON crm.crm_unnest_reference_id = w.wd_unnest_reference_id
     )
SELECT DISTINCT *
  FROM cte_crm_wd
;
