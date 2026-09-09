INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker'                            table_name
     , 'MISSING'                                              rule_type
     , 'contingent_worker_id'                                 col
     , CAST(NULL AS VARCHAR)                                  contingent_worker_id
     , CAST(NULL AS VARCHAR)                                  detail
     , 'contingent_worker_id is NULL or blank.'               message
  FROM src_hcm_contingent_worker t
 WHERE NULLIF(TRIM(t.contingent_worker_id), '') IS NULL
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker'                                                table_name
     , 'DUPLICATE'                                                                rule_type
     , 'contingent_worker_id'                                                     col
     , CAST(NULL AS VARCHAR)                                                      contingent_worker_id
     , t.contingent_worker_id                                                     detail
     , 'contingent_worker_id appears more than once in src_hcm_contingent_worker.' message
  FROM src_hcm_contingent_worker t
 WHERE NULLIF(TRIM(t.contingent_worker_id), '') IS NOT NULL
 GROUP BY t.contingent_worker_id
HAVING COUNT(*) > 1
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker'                                                                        table_name
     , 'CROSS_TAB'                                                                                        rule_type
     , 'user_id'                                                                                          col
     , cw.contingent_worker_id                                                                            contingent_worker_id
     , cw.nrm_user_id                                                                                     detail
     , 'src_hcm_contingent_worker contains a user_id flagged as a Workday match (employee or contingent).' message
  FROM src_hcm_contingent_worker cw
       INNER JOIN
       src_hcm_contingent_worker_match m
          ON NULLIF(TRIM(m.nrm_user_id), '') = NULLIF(TRIM(cw.nrm_user_id), '')
 WHERE COALESCE(m.flg_wd_match_employee, 0) = 1
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker'                                                         table_name
     , 'EXCLUDE'                                                                           rule_type
     , 'worker|legal_name|preferred_name|email_primary_work|email_primary_home'            col
     , cw.contingent_worker_id                                                             contingent_worker_id
     , CAST(
          (CASE WHEN COALESCE(m.flg_wd_match_worker, 0) = 1 THEN 'worker|' ELSE '' END) ||
          (CASE WHEN COALESCE(m.flg_wd_match_legal_name, 0) = 1 THEN 'legal_name|' ELSE '' END) ||
          (CASE WHEN COALESCE(m.flg_wd_match_preferred_name, 0) = 1 THEN 'preferred_name|' ELSE '' END) ||
          (CASE WHEN COALESCE(m.flg_wd_match_email_primary_work, 0) = 1 THEN 'email_primary_work|' ELSE '' END) ||
          (CASE WHEN COALESCE(m.flg_wd_match_email_primary_home, 0) = 1 THEN 'email_primary_home|' ELSE '' END)
          AS VARCHAR
       )                                                                                   detail
     , 'Contingent worker matches Workday employee data; exclude this contingent_worker_id.' message
  FROM src_hcm_contingent_worker cw
       INNER JOIN
       src_hcm_contingent_worker_match m
          ON NULLIF(TRIM(m.nrm_user_id), '') = NULLIF(TRIM(cw.nrm_user_id), '')
 WHERE COALESCE(m.flg_wd_match_employee, 0) = 1
   AND (
          COALESCE(m.flg_wd_match_worker            , 0) = 1
       OR COALESCE(m.flg_wd_match_legal_name        , 0) = 1
       OR COALESCE(m.flg_wd_match_preferred_name    , 0) = 1
       OR COALESCE(m.flg_wd_match_email_primary_work, 0) = 1
       OR COALESCE(m.flg_wd_match_email_primary_home, 0) = 1
       )
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker'                                        table_name
     , 'EXCLUDE'                                                          rule_type
     , 'hcm_workday_worker_all'                                           col
     , t.contingent_worker_id                                             contingent_worker_id
     , t.detail                                                           detail
     , 'Contingent worker overlaps an existing Workday employee; exclude this contingent_worker_id.' message
  FROM (
       WITH cw AS (
         SELECT
                contingent_worker_id
              , nrm_first_name
              , nrm_last_name
              , nrm_primary_email
              , nrm_secondary_email
           FROM src_hcm_contingent_worker
       )
       SELECT DISTINCT
              cw.contingent_worker_id
            , CAST(
                 'worker|' ||
                 COALESCE(NULLIF(TRIM(e.employee_id), ''), '') || '|' ||
                 COALESCE(NULLIF(TRIM(e.worker), ''), '')
                 AS VARCHAR
              ) detail
         FROM cw
              INNER JOIN hcm_workday_worker_all e
                 ON e.worker_type = 'Employee'
                AND UPPER(TRIM(cw.nrm_first_name || ' ' || cw.nrm_last_name)) = UPPER(TRIM(e.worker))

       UNION ALL
       SELECT DISTINCT
              cw.contingent_worker_id
            , CAST(
                 'legal_name|' ||
                 COALESCE(NULLIF(TRIM(e.employee_id), ''), '') || '|' ||
                 COALESCE(NULLIF(TRIM(e.legal_name_first_name), ''), '') || ' ' ||
                 COALESCE(NULLIF(TRIM(e.legal_name_last_name), ''), '')
                 AS VARCHAR
              ) detail
         FROM cw
              INNER JOIN hcm_workday_worker_all e
                 ON e.worker_type = 'Employee'
                AND UPPER(TRIM(cw.nrm_first_name)) = UPPER(TRIM(e.legal_name_first_name))
                AND UPPER(TRIM(cw.nrm_last_name )) = UPPER(TRIM(e.legal_name_last_name ))

       UNION ALL
       SELECT DISTINCT
              cw.contingent_worker_id
            , CAST(
                 'preferred_name|' ||
                 COALESCE(NULLIF(TRIM(e.employee_id), ''), '') || '|' ||
                 COALESCE(NULLIF(TRIM(e.preferred_name_first_name), ''), '') || ' ' ||
                 COALESCE(NULLIF(TRIM(e.preferred_name_last_name), ''), '')
                 AS VARCHAR
              ) detail
         FROM cw
              INNER JOIN hcm_workday_worker_all e
                 ON e.worker_type = 'Employee'
                AND UPPER(TRIM(cw.nrm_first_name)) = UPPER(TRIM(e.preferred_name_first_name))
                AND UPPER(TRIM(cw.nrm_last_name )) = UPPER(TRIM(e.preferred_name_last_name ))

       UNION ALL
       SELECT DISTINCT
              cw.contingent_worker_id
            , CAST(
                 'email_primary_work|' ||
                 COALESCE(NULLIF(TRIM(e.employee_id), ''), '') || '|' ||
                 COALESCE(NULLIF(LOWER(TRIM(e.email_primary_work)), ''), '')
                 AS VARCHAR
              ) detail
         FROM cw
              INNER JOIN hcm_workday_worker_all e
                 ON e.worker_type = 'Employee'
                AND cw.nrm_primary_email IS NOT NULL
                AND LOWER(TRIM(cw.nrm_primary_email)) = LOWER(TRIM(e.email_primary_work))

       UNION ALL
       SELECT DISTINCT
              cw.contingent_worker_id
            , CAST(
                 'email_primary_home|' ||
                 COALESCE(NULLIF(TRIM(e.employee_id), ''), '') || '|' ||
                 COALESCE(NULLIF(LOWER(TRIM(e.email_primary_home)), ''), '')
                 AS VARCHAR
              ) detail
         FROM cw
              INNER JOIN hcm_workday_worker_all e
                 ON e.worker_type = 'Employee'
                AND cw.nrm_secondary_email IS NOT NULL
                AND LOWER(TRIM(cw.nrm_secondary_email)) = LOWER(TRIM(e.email_primary_home))
       ) t
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker'                                          table_name
     , 'EXCLUDE'                                                            rule_type
     , 'hcm_workday_contingent_worker'                                      col
     , cw.contingent_worker_id                                              contingent_worker_id
     , CAST(TRIM(cw.nrm_first_name || ' ' || cw.nrm_last_name || ' [C]') AS VARCHAR) detail
     , 'Contingent worker overlaps an existing Workday contingent worker; exclude this contingent_worker_id.' message
  FROM src_hcm_contingent_worker cw
       INNER JOIN
       hcm_workday_contingent_worker c
          ON UPPER(TRIM(c.contingent_worker)) = UPPER(TRIM(cw.nrm_first_name || ' ' || cw.nrm_last_name || ' [C]'))
;
