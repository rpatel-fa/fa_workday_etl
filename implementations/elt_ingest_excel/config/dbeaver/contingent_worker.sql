SELECT * FROM hcm_okta_all;
SELECT * FROM hcm_okta_cw; 
SELECT * FROM hcm_okta_cw_line_manager_update;
SELECT * FROM hcm_workday_sup_org;
SELECT * FROM hcm_workday_worker_all;
SELECT * FROM hcm_workday_contingent_worker;
SELECT * FROM hcm_workday_account_contingent_worker ORDER BY 1;
SELECT * FROM src_hcm_contingent_worker_base;
SELECT * FROM src_hcm_contingent_worker_nrm;
SELECT * FROM src_hcm_contingent_worker_match;
SELECT * FROM src_hcm_contingent_worker;
SELECT * FROM src_hcm_contingent_worker_mgr;
SELECT * FROM validation_contingent_worker_issues;
SELECT * FROM validation_contingent_worker_result;
SELECT * FROM hcm_workday_extract_cw ORDER BY 1;


SELECT count(*) FROM hcm_workday_extract_cw;
SELECT count(*) FROM src_hcm_contingent_worker;
SELECT * FROM workday_contingent_worker_name;

SELECT * 
  FROM src_hcm_contingent_worker t
       LEFT OUTER JOIN 
       hcm_workday_extract_cw    w
         ON w.employee_id = t.contingent_worker_id
 WHERE w.employee_id IS NULL;

 
SELECT * 
  FROM src_hcm_contingent_worker t
       LEFT OUTER JOIN 
       hcm_workday_extract_cw    w
         ON w.legal_name_first_name = t.nrm_first_name
        AND w.legal_name_last_name  = t.nrm_last_name
 WHERE w.employee_id IS NULL;
 
 SELECT * 
  FROM src_hcm_contingent_worker t
       LEFT OUTER JOIN 
       hcm_workday_extract_cw    w
         ON w.email_primary_work = t.nrm_primary_email
 WHERE w.employee_id IS NULL;

 
SELECT * FROM hcm_workday_extract_cw WHERE worker LIKE '%Rael%';

SELECT * FROM hcm_workday_extract_cw WHERE employee_id = 'CW-100104';

SELECT * FROM src_hcm_contingent_worker WHERE contingent_worker_id NOT LIKE 'CW%';

SELECT * FROM hcm_workday_worker_all WHERE TRIM(upper(worker)) LIKE upper('%Emma%') ORDER BY 2;

SELECT * FROM hcm_workday_worker_all WHERE employee_id LIKE 'CW-%'

SELECT * FROM hcm_workday_worker_all WHERE worker_type = 'Contingent Worker'

SELECT * FROM hcm_okta_all WHERE employee_number LIKE 'CW%';

SELECT * FROM hcm_okta_cw WHERE employee_number NOT LIKE 'CW%';


SELECT employee_number, * FROM src_hcm_contingent_worker WHERE employee_number  NOT LIKE '%CW%' ORDER BY 3;



Surendran Ravichandran surendran.ravichandran@thefa.com

SELECT * FROM src_hcm_contingent_worker_match ORDER BY 3;

SELECT user_id, username from src_hcm_contingent_worker WHERE nrm_first_name LIKE '%Surendran%';

00u5kbjcelflOC1JH4x7
00u5kbjcelfloc1jh4x7

SELECT* FROM hcm_okta_cw_line_manager_update;

SELECT lower(user_id), count(*) FROM hcm_okta_cw GROUP BY lower(user_id) HAVING count(*)>1;

SELECT * 