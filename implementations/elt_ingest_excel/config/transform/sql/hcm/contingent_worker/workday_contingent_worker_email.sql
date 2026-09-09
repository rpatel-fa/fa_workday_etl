DROP TABLE IF EXISTS workday_contingent_worker_email
;
CREATE TABLE workday_contingent_worker_email
    AS
  WITH cte_worker
    AS (
SELECT
       t.contingent_worker_id                            worker_id
     , t.nrm_first_name                                  first_name
     , t.nrm_last_name                                   last_name
     , 'Contingent Worker'                               worker_type
     , 'Work'                                            email_type
     , t.nrm_primary_email                               email_address
     , 'Y'                                               is_public
     , 'Y'                                               is_primary
  FROM src_hcm_contingent_worker t
 WHERE t.nrm_primary_email             IS NOT NULL
UNION ALL
SELECT
       t.contingent_worker_id                            worker_id
     , t.nrm_first_name                                  first_name
     , t.nrm_last_name                                   last_name
     , 'Contingent Worker'                               worker_type
     , 'Other'                                           email_type
     , t.nrm_secondary_email                             email_address
     , 'Y'                                               is_public
     , 'N'                                               is_primary
  FROM src_hcm_contingent_worker t
 WHERE t.nrm_secondary_email           IS NOT NULL
       )
SELECT t.*
  FROM cte_worker                      t
 WHERE t.email_address                 IS NOT NULL
 ORDER BY
       t.worker_id
     , t.is_primary DESC
;
