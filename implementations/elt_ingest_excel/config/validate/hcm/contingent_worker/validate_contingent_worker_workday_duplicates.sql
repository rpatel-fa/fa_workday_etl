INSERT INTO validation_contingent_worker_issues
SELECT
       'workday_contingent_worker_name'                         table_name
     , 'DUPLICATE'                                               rule_type
     , 'worker_id'                                               col
     , t.worker_id                                               contingent_worker_id
     , CAST(t.cnt AS VARCHAR)                                    detail
     , 'worker_id appears more than once in workday_contingent_worker_name.' message
  FROM (
       SELECT
              worker_id
            , COUNT(*) cnt
         FROM workday_contingent_worker_name
        WHERE NULLIF(TRIM(worker_id), '') IS NOT NULL
        GROUP BY worker_id
       HAVING COUNT(*) > 1
       ) t
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'workday_contingent_worker_manager'                      table_name
     , 'DUPLICATE'                                               rule_type
     , 'worker_id'                                               col
     , t.worker_id                                               contingent_worker_id
     , CAST(t.cnt AS VARCHAR)                                    detail
     , 'worker_id appears more than once in workday_contingent_worker_manager.' message
  FROM (
       SELECT
              worker_id
            , COUNT(*) cnt
         FROM workday_contingent_worker_manager
        WHERE NULLIF(TRIM(worker_id), '') IS NOT NULL
        GROUP BY worker_id
       HAVING COUNT(*) > 1
       ) t
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'workday_contingent_worker_professional_ee'               table_name
     , 'DUPLICATE'                                               rule_type
     , 'worker_id'                                               col
     , t.worker_id                                               contingent_worker_id
     , CAST(t.cnt AS VARCHAR)                                    detail
     , 'worker_id appears more than once in workday_contingent_worker_professional_ee.' message
  FROM (
       SELECT
              worker_id
            , COUNT(*) cnt
         FROM workday_contingent_worker_professional_ee
        WHERE NULLIF(TRIM(worker_id), '') IS NOT NULL
        GROUP BY worker_id
       HAVING COUNT(*) > 1
       ) t
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'workday_contingent_worker_email'                         table_name
     , 'DUPLICATE'                                               rule_type
     , 'worker_id'                                               col
     , t.worker_id                                               contingent_worker_id
     , CAST(t.cnt AS VARCHAR)                                    detail
     , 'worker_id has more than 2 rows in workday_contingent_worker_email.' message
  FROM (
       SELECT
              worker_id
            , COUNT(*) cnt
         FROM workday_contingent_worker_email
        WHERE NULLIF(TRIM(worker_id), '') IS NOT NULL
        GROUP BY worker_id
       HAVING COUNT(*) > 2
       ) t
;
