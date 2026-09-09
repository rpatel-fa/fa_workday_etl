DROP TABLE IF EXISTS ref_post_code_workday_region
;
CREATE TABLE ref_post_code_workday_region
    AS
SELECT
       TRIM(post_code_region)          post_code_region
     , TRIM(workday_region)            workday_region
  FROM read_json_auto(
           'config/data/ref_post_code_workday_region.json'
       )
;
