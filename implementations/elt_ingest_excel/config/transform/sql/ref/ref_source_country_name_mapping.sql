DROP TABLE IF EXISTS ref_source_country_name_mapping
;
CREATE TABLE ref_source_country_name_mapping
    AS
SELECT
       TRIM(source_country_name)         source_country_name
     , TRIM(country_code)                country_code
  FROM read_json_auto(
           'config/data/ref_source_country_name_mapping.json'
       )
;
