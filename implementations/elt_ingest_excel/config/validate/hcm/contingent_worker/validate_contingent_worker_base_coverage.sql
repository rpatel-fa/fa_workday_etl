INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker'                                         table_name
     , 'CROSS_TAB'                                                         rule_type
     , 'user_id|username'                                                  col
     , s.contingent_worker_id                                              contingent_worker_id
     , COALESCE(NULLIF(TRIM(s.user_id), ''), NULLIF(TRIM(s.username), ''))  detail
     , 'src_hcm_contingent_worker row has no corresponding record in src_hcm_contingent_worker_base (by user_id or username).' message
  FROM src_hcm_contingent_worker s
       LEFT OUTER JOIN
       src_hcm_contingent_worker_base b
          ON (
             (NULLIF(TRIM(s.user_id), '') IS NOT NULL AND NULLIF(TRIM(b.user_id), '') = NULLIF(TRIM(s.user_id), ''))
          OR (NULLIF(TRIM(s.user_id), '') IS NULL AND NULLIF(TRIM(s.username), '') IS NOT NULL AND NULLIF(TRIM(b.username), '') = NULLIF(TRIM(s.username), ''))
             )
 WHERE b.username IS NULL
   AND b.user_id IS NULL
;
