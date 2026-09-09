DROP TABLE IF EXISTS src_hcm_contingent_worker_base
;
CREATE TABLE src_hcm_contingent_worker_base
    AS
  WITH cte_contingent_worker_base
    AS (
SELECT DISTINCT
       t.* 
  FROM hcm_okta_cw                     t
 WHERE 
       1 = 1
   AND UPPER(TRIM(t.new_status)) <> UPPER(TRIM('Deprovisioned'))
   AND COALESCE(
         NULLIF(TRIM(t.user_id)        , '')
       , NULLIF(TRIM(t.username)       , '')
       , NULLIF(TRIM(t.primary_email)  , '')
       , NULLIF(TRIM(t.secondary_email), '')
       ) IS NOT NULL
   AND NOT
       (
       COALESCE(NULLIF(UPPER(TRIM(t.location))    , ''), 'NULL') IN ('COUNTY', 'SERVICE_ACCOUNT', 'SERVICE ACCOUNT')
    OR COALESCE(NULLIF(UPPER(TRIM(t.department))  , ''), 'NULL') IN ('COUNTY', 'SERVICE_ACCOUNT', 'SERVICE ACCOUNT')
    OR COALESCE(NULLIF(UPPER(TRIM(t.department_1)), ''), 'NULL') IN ('COUNTY', 'SERVICE_ACCOUNT', 'SERVICE ACCOUNT')
    OR COALESCE(NULLIF(UPPER(TRIM(t.user_type))   , ''), 'NULL') IN ('COUNTY', 'SERVICE_ACCOUNT', 'SERVICE ACCOUNT')
       )
   AND NOT (
       UPPER(TRIM(t.first_name))                                   LIKE '%ROOM%'
    OR UPPER(TRIM(t.last_name ))                                   LIKE '%ROOM%'
       )
   AND NOT (
       UPPER(TRIM(t.first_name))                                   LIKE '%APP%'
   AND UPPER(TRIM(t.last_name ))                                   LIKE '%DEV%'
       )
   AND NOT (
       UPPER(TRIM(t.first_name))                                   LIKE '%COACHING'
   AND UPPER(TRIM(t.last_name ))                                   LIKE 'OFFICE'
       )
   AND NOT (
       UPPER(TRIM(t.first_name))                                      = 'CREDIT'
   AND UPPER(TRIM(t.last_name ))                                   LIKE 'CONTROL'
       )
   AND NOT (
       UPPER(TRIM(t.first_name))                                      = 'GRASSROOT'
   AND UPPER(TRIM(t.last_name ))                                   LIKE 'DATA PRODUCT'
       )
   AND NOT (
       UPPER(TRIM(t.first_name))                                      = 'CLOUDFLARE'
   AND UPPER(TRIM(t.last_name ))                                   LIKE 'BREAKGLASS%'
       )
   AND NOT (
       UPPER(TRIM(t.first_name))                                      = 'CMA'
       )
   AND NOT (
       UPPER(TRIM(t.first_name))                                      = 'CET'
       )
   AND NOT (
       UPPER(TRIM(t.last_name))                                       = 'GATEHOUSE'
       )
   AND NOT (
       UPPER(TRIM(t.first_name))                                      = 'RECEPTION'
    OR UPPER(TRIM(t.last_name ))                                      = 'RECPETION'
       )
   AND NOT (
       UPPER(TRIM(t.first_name))                                      = 'C'
   AND UPPER(TRIM(t.last_name ))                                   LIKE '%RELATIONS%'
       )
       )
     , cte_contingent_worker_rnk
    AS (
SELECT t.*
     , ROW_NUMBER() OVER
       (PARTITION BY username
            ORDER BY created_date      DESC NULLS LAST
       )                               data_rnk
  FROM cte_contingent_worker_base      t
       )
SELECT t.*
  FROM cte_contingent_worker_rnk       t
 WHERE t.data_rnk                      = 1
;
