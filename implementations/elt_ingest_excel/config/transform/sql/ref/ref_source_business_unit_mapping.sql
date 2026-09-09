DROP TABLE IF EXISTS ref_source_business_unit_mapping
;
CREATE TABLE ref_source_business_unit_mapping
    AS
SELECT
       TRIM(source_value)               source_value
     , TRIM(target_value)               target_value
  FROM read_json_auto(
           'config/data/ref_source_business_unit_mapping.json'
       )
;
