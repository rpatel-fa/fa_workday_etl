DROP TABLE IF EXISTS ref_customer_category
;
CREATE TABLE ref_customer_category
    AS
SELECT
       TRIM(source_value)              source_value
     , TRIM(target_value)              target_value
  FROM read_json_auto(
           'config/data/ref_customer_category.json'
       )
;
