DROP TABLE IF EXISTS ref_source_country_code_mapping
;
CREATE TABLE ref_source_country_code_mapping
    AS
SELECT
       TRIM(source_country_code)          source_country_code
     , TRIM(country_code)                 country_code
  FROM read_json_auto(
           'config/data/ref_source_country_code_mapping.json'
       )
;
