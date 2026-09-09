DROP TABLE IF EXISTS ref_country_iso3_mapping
;
CREATE TABLE ref_country_iso3_mapping
    AS
SELECT
       TRIM(country_name)          country_name
     , TRIM(country_code_iso2)     country_code_iso2
     , TRIM(country_code_iso3)     country_code_iso3
  FROM read_json_auto(
           'config/data/ref_country_iso3_mapping.json'
       )
;
