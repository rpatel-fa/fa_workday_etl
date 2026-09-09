DROP TABLE IF EXISTS ref_worker_type_mapping
;
CREATE TABLE ref_worker_type_mapping
    AS
SELECT
       TRIM(user_type)              user_type
     , TRIM(department_1)               department_1
     , TRIM(mapped_value)               mapped_value
  FROM read_json_auto(
           'config/data/ref_worker_type_mapping.json'
       )
;
