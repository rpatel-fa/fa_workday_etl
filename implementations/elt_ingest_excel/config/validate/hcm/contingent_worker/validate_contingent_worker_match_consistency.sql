INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_match'                         table_name
     , 'DUPLICATE'                                               rule_type
     , 'user_id'                                                 col
     , CAST(NULL AS VARCHAR)                                     contingent_worker_id
     , t.nrm_user_id                                             detail
     , 'user_id appears more than once in src_hcm_contingent_worker_match.' message
  FROM src_hcm_contingent_worker_match t
 WHERE NULLIF(TRIM(t.nrm_user_id), '') IS NOT NULL
 GROUP BY t.nrm_user_id
HAVING COUNT(*) > 1
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_match'                                                      table_name
     , 'CONSISTENCY'                                                                           rule_type
     , 'flg_wd_match_employee'                                                                 col
     , CAST(NULL AS VARCHAR)                                                                   contingent_worker_id
     , t.nrm_user_id                                                                           detail
     , 'flg_wd_match_employee is inconsistent with the sum of match flags.'                    message
  FROM src_hcm_contingent_worker_match t
 WHERE COALESCE(t.flg_wd_match_employee, 0) <>
       CASE
          WHEN (
                    COALESCE(t.flg_wd_match_worker            , 0)
                  + COALESCE(t.flg_wd_match_legal_name        , 0)
                  + COALESCE(t.flg_wd_match_preferred_name    , 0)
                  + COALESCE(t.flg_wd_match_email_primary_work, 0)
                  + COALESCE(t.flg_wd_match_email_primary_home, 0)
               ) > 0
          THEN 1
          ELSE 0
       END
;
