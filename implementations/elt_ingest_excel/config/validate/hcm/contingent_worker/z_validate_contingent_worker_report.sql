SELECT *
  FROM validation_contingent_worker_issues
 ORDER BY
       table_name
     , rule_type
     , col
     , contingent_worker_id
     , detail
;
