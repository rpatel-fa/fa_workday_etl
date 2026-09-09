DROP TABLE IF EXISTS ref_workday_group_additional
;
CREATE TABLE ref_workday_group_additional
    AS
SELECT
       TRIM(source_type)                   source_type
     , TRIM(source_value)                  source_value
     , TRIM(target_value)                  target_value
  FROM read_json_auto(
             'config/data/ref_workday_group_additional.json'
       )
;
