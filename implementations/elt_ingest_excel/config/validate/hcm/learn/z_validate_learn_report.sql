SELECT *
  FROM validation_learn_issues
 ORDER BY
       table_name
     , rule_type
     , col
     , learn_id
     , detail
;
