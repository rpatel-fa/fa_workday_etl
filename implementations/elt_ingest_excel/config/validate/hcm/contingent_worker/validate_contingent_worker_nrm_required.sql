INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_nrm'                table_name
     , 'MISSING'                                      rule_type
     , 'nrm_location'                                 col
     , CAST(NULL AS VARCHAR)                          contingent_worker_id
     , COALESCE(t.user_id, t.username)                detail
     , 'nrm_location is NULL.'                        message
  FROM src_hcm_contingent_worker_nrm t
 WHERE t.nrm_location IS NULL
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_nrm'                table_name
     , 'MISSING'                                      rule_type
     , 'nrm_worker_type'                              col
     , CAST(NULL AS VARCHAR)                          contingent_worker_id
     , COALESCE(t.user_id, t.username)                detail
     , 'nrm_worker_type is NULL.'                     message
  FROM src_hcm_contingent_worker_nrm t
 WHERE t.nrm_worker_type IS NULL
;
