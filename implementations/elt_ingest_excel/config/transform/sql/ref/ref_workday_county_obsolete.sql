DROP TABLE IF EXISTS ref_workday_county_obsolete
;
CREATE TABLE ref_workday_county_obsolete
    AS
SELECT
       TRIM(source_value)                  source_value
     , TRIM(target_value)                  target_value
  FROM read_json_auto(
             'config/data/ref_workday_county_obsolete.json'
       )
;
