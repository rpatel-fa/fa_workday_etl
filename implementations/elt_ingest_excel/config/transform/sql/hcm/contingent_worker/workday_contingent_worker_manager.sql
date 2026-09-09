DROP TABLE IF EXISTS workday_contingent_worker_manager
;
CREATE TABLE workday_contingent_worker_manager
    AS
  WITH cte_default_manager
    AS (
SELECT
       employee_id                     default_manager_id
     , legal_name_first_name           default_manager_first_name
     , legal_name_last_name            default_manager_last_name
  FROM hcm_workday_worker_all
 WHERE LOWER(TRIM(worker)) = LOWER(TRIM('Sam Vickers'))
 LIMIT 1
       )
SELECT
       t.worker_id
     , t.position_id
     , t.worker_id_of_direct_manager
     , t.manager_first_name
     , t.manager_last_name
     , t.first_name
     , t.last_name
     , t.worker_type
     , t.manager_position_id
     , t.supervisory_org_code
     , t.supervisory_org_name
     , t.supervisory_id
     , t.staffing_model
  FROM (
       SELECT
              cw.contingent_worker_id                                            worker_id
            , NULL                                                               position_id
            , COALESCE(mgr.resolved_manager_id, dm.default_manager_id)           worker_id_of_direct_manager
            , CASE
                 WHEN mgr.resolved_manager_id IS NULL
                 THEN dm.default_manager_first_name
                 ELSE mgr.resolved_manager_first_name
              END                                                                manager_first_name
            , CASE
                 WHEN mgr.resolved_manager_id IS NULL
                 THEN dm.default_manager_last_name
                 ELSE mgr.resolved_manager_last_name
              END                                                                manager_last_name
            , cw.nrm_first_name                                                  first_name
            , cw.nrm_last_name                                                   last_name
            , 'Contingent Worker'                                                worker_type
            , NULL                                                               manager_position_id
            , NULL                                                               supervisory_org_code
            , NULL                                                               supervisory_org_name
            , NULL                                                               supervisory_id
            , NULL                                                               staffing_model
         FROM src_hcm_contingent_worker         cw
              CROSS JOIN
              cte_default_manager               dm
              LEFT OUTER JOIN 
              src_hcm_contingent_worker_mgr     mgr
                ON mgr.contingent_worker_id     = cw.contingent_worker_id
       ) t
 WHERE t.worker_id_of_direct_manager            IS NOT NULL
 ORDER BY 
       t.worker_id
;
