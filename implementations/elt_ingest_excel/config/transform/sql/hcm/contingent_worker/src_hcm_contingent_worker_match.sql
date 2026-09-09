DROP TABLE IF EXISTS src_hcm_contingent_worker_match
;
CREATE TABLE src_hcm_contingent_worker_match
    AS
  WITH cte_workday_worker
    AS (
SELECT
       NULLIF(LOWER(TRIM(t.employee_id)), '')               employee_id
     , NULLIF(LOWER(TRIM(t.worker_type)), '')               worker_type
     , NULLIF(LOWER(TRIM(t.employee_type)), '')             employee_type
     , NULLIF(LOWER(TRIM(t.worker)), '')                    worker
     , NULLIF(LOWER(TRIM(t.legal_name_first_name)), '')     legal_first_name
     , NULLIF(LOWER(TRIM(t.legal_name_last_name)), '')      legal_last_name
     , NULLIF(LOWER(TRIM(t.preferred_name_first_name)), '') preferred_first_name
     , NULLIF(LOWER(TRIM(t.preferred_name_last_name)), '')  preferred_last_name
     , NULLIF(LOWER(TRIM(t.email_primary_work)), '')        email_primary_work
     , NULLIF(LOWER(TRIM(t.email_primary_home)), '')        email_primary_home
  FROM hcm_workday_worker_all          t
 WHERE worker_type                     = 'Employee'
UNION ALL
SELECT
       NULLIF(LOWER(TRIM(t.employee_id)), '')               employee_id
     , NULLIF(LOWER(TRIM(t.worker_type)), '')               worker_type
     , NULLIF(LOWER(TRIM(t.employee_type)), '')             employee_type
     , NULLIF(LOWER(TRIM(t.worker)), '')                    worker
     , NULLIF(LOWER(TRIM(t.legal_name_first_name)), '')     legal_first_name
     , NULLIF(LOWER(TRIM(t.legal_name_last_name)), '')      legal_last_name
     , NULLIF(LOWER(TRIM(t.preferred_name_first_name)), '') preferred_first_name
     , NULLIF(LOWER(TRIM(t.preferred_name_last_name)), '')  preferred_last_name
     , NULLIF(LOWER(TRIM(t.email_primary_work)), '')        email_primary_work
     , NULLIF(LOWER(TRIM(t.email_primary_home)), '')        email_primary_home
  FROM hcm_workday_worker_all             t
       INNER JOIN
       hcm_workday_contingent_worker      wcw
          ON wcw.contingent_worker        = t.worker
 WHERE t.worker_type                     <> 'Employee'
       )
     , cte_src
    AS (
SELECT
       NULLIF(LOWER(TRIM(cw.nrm_user_id)), '')                                         nrm_user_id
     , NULLIF(LOWER(TRIM(cw.nrm_username)), '')                                        nrm_username
     , NULLIF(LOWER(TRIM(cw.nrm_first_name)), '')                                      nrm_first_name
     , NULLIF(LOWER(TRIM(cw.nrm_last_name)), '')                                       nrm_last_name
     , NULLIF(LOWER(TRIM(cw.nrm_primary_email)), '')                                   nrm_primary_email
     , NULLIF(LOWER(TRIM(cw.nrm_secondary_email)), '')                                 nrm_secondary_email
     , NULLIF(LOWER(TRIM(cw.nrm_first_name || ' ' || cw.nrm_last_name)), '')           nrm_worker_key_employee
     , NULLIF(LOWER(TRIM(cw.nrm_first_name || ' ' || cw.nrm_last_name || ' [C]')), '') nrm_worker_key_cw
  FROM src_hcm_contingent_worker_nrm cw
 WHERE NULLIF(TRIM(cw.nrm_user_id) , '') IS NOT NULL
       )
     , cte_emp_match
    AS (
SELECT
       s.*
     , w.*
     , CASE
          WHEN w.worker = s.nrm_worker_key_employee
            OR w.worker = s.nrm_worker_key_cw 
          THEN 1
          ELSE 0
       END                                                        flg_wd_match_worker
     , CASE
          WHEN w.legal_first_name = s.nrm_first_name
           AND w.legal_last_name  = s.nrm_last_name
          THEN 1
          ELSE 0
       END                                                        flg_wd_match_legal_name
     , CASE
          WHEN w.preferred_first_name = s.nrm_first_name
           AND w.preferred_last_name  = s.nrm_last_name
          THEN 1
          ELSE 0
       END                                                        flg_wd_match_preferred_name
     , CASE
          WHEN s.nrm_primary_email IS NOT NULL
           AND w.email_primary_work = s.nrm_primary_email
          THEN 1
          ELSE 0
       END                                                        flg_wd_match_email_primary_work
     , CASE
          WHEN s.nrm_secondary_email IS NOT NULL
           AND w.email_primary_home = s.nrm_secondary_email
          THEN 1
          ELSE 0
       END                                                        flg_wd_match_email_primary_home
  FROM cte_src                         s
       INNER JOIN 
       cte_workday_worker              w
          ON (
                   w.worker                        = s.nrm_worker_key_employee
                OR w.worker                        = s.nrm_worker_key_cw
                OR (w.legal_first_name             = s.nrm_first_name AND w.legal_last_name          = s.nrm_last_name)
                OR (w.preferred_first_name         = s.nrm_first_name AND w.preferred_last_name      = s.nrm_last_name)
                OR (s.nrm_primary_email   IS NOT NULL AND w.email_primary_work = s.nrm_primary_email)
                OR (s.nrm_secondary_email IS NOT NULL AND w.email_primary_home = s.nrm_secondary_email)
             )
       )
SELECT t.*       
     , CASE
          WHEN (
                    t.flg_wd_match_worker
                  + t.flg_wd_match_legal_name
                  + t.flg_wd_match_preferred_name
                  + t.flg_wd_match_email_primary_work
                  + t.flg_wd_match_email_primary_home
               ) > 0
          THEN 1
          ELSE 0
       END flg_wd_match_employee
  FROM cte_emp_match                   t
;
