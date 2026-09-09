DROP TABLE IF EXISTS src_hcm_contingent_worker_mgr
;
CREATE TABLE src_hcm_contingent_worker_mgr
    AS
  WITH cte_cw
    AS (
SELECT
       cw.contingent_worker_id
     , cw.nrm_first_name
     , cw.nrm_last_name
     , cw.nrm_managerid
     , cw.nrm_manager_name
     , LOWER(TRIM(COALESCE(NULLIF(TRIM(mu.suggested_manager_email), ''), cw.nrm_manager_email_address))) nrm_manager_email_address
     , cw.nrm_manager_email
  FROM src_hcm_contingent_worker       cw
       LEFT OUTER JOIN
       hcm_okta_cw_line_manager_update mu
          ON NULLIF(LOWER(TRIM(mu.user_id)), '') = cw.nrm_user_id
       )
     , cte_mgr_workday
    AS (
SELECT t.*
  FROM hcm_workday_worker_all t
 WHERE t.worker_type          = 'Employee'
UNION ALL 
SELECT wdw.*
  FROM hcm_workday_contingent_worker wdcw
       INNER JOIN 
       hcm_workday_worker_all        wdw
          ON wdw.worker              = wdcw.contingent_worker
         AND wdw.worker_type        <> 'Employee'
       )
     , cte_mgr_employee
    AS (
SELECT
       NULLIF(TRIM(t.employee_id), '')                                                                   mgr_emp_employee_id
     , NULLIF(LOWER(TRIM(t.worker)), '')                                                                 mgr_emp_worker
     , t.worker_type                                                                                     mgr_emp_worker_type
     , NULLIF(LOWER(TRIM(t.email_primary_work)), '')                                                     mgr_emp_email_primary_work
     , NULLIF(LOWER(TRIM(t.email_primary_home)), '')                                                     mgr_emp_email_primary_home
     , TRIM(t.legal_name_first_name)                                                                     mgr_emp_first_name
     , TRIM(t.legal_name_last_name )                                                                     mgr_emp_last_name
     , NULLIF(LOWER(TRIM(t.legal_name_first_name)     || ' ' || TRIM(t.legal_name_last_name)), '')       mgr_emp_legal_name
     , NULLIF(LOWER(TRIM(t.preferred_name_first_name) || ' ' || TRIM(t.preferred_name_last_name)), '')   mgr_emp_preferred_name
  FROM cte_mgr_workday        t
       )
     , cte_mgr_employee_match
    AS (
SELECT
       cw.contingent_worker_id
     , e.mgr_emp_employee_id
     , e.mgr_emp_worker
     , e.mgr_emp_worker_type
     , e.mgr_emp_email_primary_work
     , e.mgr_emp_email_primary_home
     , e.mgr_emp_first_name
     , e.mgr_emp_last_name
     , e.mgr_emp_legal_name
     , e.mgr_emp_preferred_name
     , CASE
          WHEN cw.nrm_managerid IS NOT NULL
           AND LOWER(e.mgr_emp_employee_id) = LOWER(cw.nrm_managerid)
          THEN 1
          WHEN cw.nrm_manager_email_address        IS NOT NULL
           AND LOWER(e.mgr_emp_email_primary_work) = LOWER(cw.nrm_manager_email_address)
           AND cw.nrm_manager_name                 IS NOT NULL
           AND LOWER(e.mgr_emp_legal_name)         = LOWER(cw.nrm_manager_name)
          THEN 2
          WHEN cw.nrm_manager_email                IS NOT NULL
           AND LOWER(e.mgr_emp_email_primary_work) = LOWER(cw.nrm_manager_email)
           AND cw.nrm_manager_name                 IS NOT NULL
           AND LOWER(e.mgr_emp_legal_name)         = LOWER(cw.nrm_manager_name)
          THEN 3
          WHEN cw.nrm_manager_email_address        IS NOT NULL
           AND LOWER(e.mgr_emp_email_primary_work) = LOWER(cw.nrm_manager_email_address)
          THEN 4
          WHEN cw.nrm_manager_email                IS NOT NULL
           AND LOWER(e.mgr_emp_email_primary_work) = LOWER(cw.nrm_manager_email)
          THEN 5
          WHEN cw.nrm_manager_name                 IS NOT NULL
           AND LOWER(e.mgr_emp_legal_name)         = LOWER(cw.nrm_manager_name)
          THEN 6
          WHEN cw.nrm_manager_name                 IS NOT NULL
           AND LOWER(e.mgr_emp_preferred_name)     = LOWER(cw.nrm_manager_name)
          THEN 7
          WHEN cw.nrm_manager_email_address        IS NOT NULL
           AND LOWER(e.mgr_emp_email_primary_home) = LOWER(cw.nrm_manager_email_address)
          THEN 8
          WHEN cw.nrm_manager_email                IS NOT NULL
           AND LOWER(e.mgr_emp_email_primary_home) = LOWER(cw.nrm_manager_email)
          THEN 9
          ELSE 999
       END                                                                                           mgr_emp_match_priority
  FROM cte_cw             cw
       LEFT OUTER JOIN
       cte_mgr_employee   e
         ON (
            (cw.nrm_managerid             IS NOT NULL AND LOWER(e.mgr_emp_employee_id       ) = LOWER(cw.nrm_managerid))
         OR (cw.nrm_manager_name          IS NOT NULL AND LOWER(e.mgr_emp_legal_name        ) = LOWER(cw.nrm_manager_name))
         OR (cw.nrm_manager_name          IS NOT NULL AND LOWER(e.mgr_emp_preferred_name    ) = LOWER(cw.nrm_manager_name))
         OR (cw.nrm_manager_email_address IS NOT NULL AND LOWER(e.mgr_emp_email_primary_work) = LOWER(cw.nrm_manager_email_address))
         OR (cw.nrm_manager_email         IS NOT NULL AND LOWER(e.mgr_emp_email_primary_work) = LOWER(cw.nrm_manager_email))
         OR (cw.nrm_manager_email_address IS NOT NULL AND LOWER(e.mgr_emp_email_primary_home) = LOWER(cw.nrm_manager_email_address))
         OR (cw.nrm_manager_email         IS NOT NULL AND LOWER(e.mgr_emp_email_primary_home) = LOWER(cw.nrm_manager_email))
            )
       )
     , cte_mgr_employee_best
    AS (
SELECT t.*
  FROM (
       SELECT
              t.*
            , ROW_NUMBER() OVER (PARTITION BY t.contingent_worker_id ORDER BY t.mgr_emp_match_priority) rnk
         FROM cte_mgr_employee_match t
        WHERE t.mgr_emp_employee_id IS NOT NULL
       ) t
 WHERE t.rnk = 1
       )
     , cte_mgr_src_contingent_worker_match
    AS (
SELECT
       cw.contingent_worker_id                                                         contingent_worker_id
     , cwm.contingent_worker_id                                                        mgr_src_cw_contingent_worker_id
     , cwm.nrm_first_name                                                              mgr_src_cw_first_name
     , cwm.nrm_last_name                                                               mgr_src_cw_last_name
     , cwm.nrm_primary_email                                                           mgr_src_cw_primary_email
     , cwm.nrm_secondary_email                                                         mgr_src_cw_secondary_email
     , CASE
          WHEN cw.nrm_manager_email_address     IS NOT NULL
           AND LOWER(cwm.nrm_primary_email)     = LOWER(cw.nrm_manager_email_address)
           AND cw.nrm_manager_name              IS NOT NULL
           AND LOWER(cwm.nrm_first_name || ' ' || cwm.nrm_last_name) = LOWER(cw.nrm_manager_name)
          THEN 1
          WHEN cw.nrm_manager_email             IS NOT NULL
           AND LOWER(cwm.nrm_primary_email)     = LOWER(cw.nrm_manager_email)
           AND cw.nrm_manager_name              IS NOT NULL
           AND LOWER(cwm.nrm_first_name || ' ' || cwm.nrm_last_name) = LOWER(cw.nrm_manager_name)
          THEN 2
          WHEN cw.nrm_manager_email_address     IS NOT NULL
           AND LOWER(cwm.nrm_primary_email)     = LOWER(cw.nrm_manager_email_address)
          THEN 3
          WHEN cw.nrm_manager_email             IS NOT NULL
           AND LOWER(cwm.nrm_primary_email)     = LOWER(cw.nrm_manager_email)
          THEN 4
          WHEN cw.nrm_manager_name              IS NOT NULL
           AND LOWER(cwm.nrm_first_name || ' ' || cwm.nrm_last_name) = LOWER(cw.nrm_manager_name)
          THEN 5
          WHEN cw.nrm_manager_email_address     IS NOT NULL
           AND LOWER(cwm.nrm_secondary_email)   = LOWER(cw.nrm_manager_email_address)
          THEN 6
          WHEN cw.nrm_manager_email             IS NOT NULL
           AND LOWER(cwm.nrm_secondary_email)   = LOWER(cw.nrm_manager_email)
          THEN 7
          ELSE 999
       END                                                                            mgr_src_cw_match_priority
  FROM cte_cw                         cw
       LEFT OUTER JOIN
       src_hcm_contingent_worker      cwm
         ON (
            cw.contingent_worker_id <> cwm.contingent_worker_id
        AND NOT 
            (
               (LOWER(TRIM(cwm.nrm_first_name)) = LOWER('Michaela') AND LOWER(TRIM(cwm.nrm_last_name)) = LOWER('Garrard'))
            OR (LOWER(TRIM(cwm.nrm_first_name)) = LOWER('Pancras')  AND LOWER(TRIM(cwm.nrm_last_name)) = LOWER('Valentine'))
            )
        AND (
               (cw.nrm_manager_name          IS NOT NULL AND LOWER(cwm.nrm_first_name || ' ' || cwm.nrm_last_name) = LOWER(cw.nrm_manager_name))
            OR (cw.nrm_manager_email_address IS NOT NULL AND LOWER(cwm.nrm_primary_email)    = LOWER(cw.nrm_manager_email_address))
            OR (cw.nrm_manager_email         IS NOT NULL AND LOWER(cwm.nrm_primary_email)    = LOWER(cw.nrm_manager_email))
            OR (cw.nrm_manager_email_address IS NOT NULL AND LOWER(cwm.nrm_secondary_email)  = LOWER(cw.nrm_manager_email_address))
            OR (cw.nrm_manager_email         IS NOT NULL AND LOWER(cwm.nrm_secondary_email)  = LOWER(cw.nrm_manager_email))
            )
            )
       )
     , cte_mgr_src_contingent_worker_best
    AS (
SELECT t.*
  FROM (
       SELECT
              t.*
            , ROW_NUMBER() OVER (PARTITION BY t.contingent_worker_id ORDER BY t.mgr_src_cw_match_priority) rnk
         FROM cte_mgr_src_contingent_worker_match t
        WHERE t.mgr_src_cw_contingent_worker_id IS NOT NULL
       ) t
 WHERE t.rnk = 1
       )
     , cte_cw_mgr
    AS (
SELECT
       cw.contingent_worker_id
     , cw.nrm_first_name
     , cw.nrm_last_name
     , cw.nrm_managerid
     , cw.nrm_manager_name
     , cw.nrm_manager_email_address
     , cw.nrm_manager_email
     , me.mgr_emp_employee_id
     , me.mgr_emp_worker
     , me.mgr_emp_worker_type
     , me.mgr_emp_email_primary_work
     , me.mgr_emp_email_primary_home
     , me.mgr_emp_legal_name
     , me.mgr_emp_preferred_name
     , me.mgr_emp_match_priority
     , ms.mgr_src_cw_contingent_worker_id
     , ms.mgr_src_cw_first_name
     , ms.mgr_src_cw_last_name
     , ms.mgr_src_cw_primary_email
     , ms.mgr_src_cw_secondary_email
     , ms.mgr_src_cw_match_priority
     , COALESCE(me.mgr_emp_employee_id, ms.mgr_src_cw_contingent_worker_id) resolved_manager_id
     , CASE
          WHEN me.mgr_emp_employee_id IS NOT NULL
          THEN 'Employee'
          WHEN ms.mgr_src_cw_contingent_worker_id IS NOT NULL
          THEN 'Contingent Worker (src)'
          ELSE NULL
       END                                                                                         resolved_manager_source
     , COALESCE(me.mgr_emp_first_name, ms.mgr_src_cw_first_name)                                   resolved_manager_first_name
     , COALESCE(me.mgr_emp_last_name , ms.mgr_src_cw_last_name )                                   resolved_manager_last_name
     , COALESCE(me.mgr_emp_legal_name, ms.mgr_src_cw_first_name || ' ' || ms.mgr_src_cw_last_name) resolved_manager_name
     , COALESCE(me.mgr_emp_email_primary_work, ms.mgr_src_cw_primary_email)                        resolved_manager_email
  FROM cte_cw                            cw
       LEFT OUTER JOIN cte_mgr_employee_best                     me ON me.contingent_worker_id = cw.contingent_worker_id
       LEFT OUTER JOIN cte_mgr_src_contingent_worker_best        ms ON ms.contingent_worker_id = cw.contingent_worker_id
       )
SELECT *
  FROM cte_cw_mgr
 WHERE resolved_manager_source IS NOT NULL
;
