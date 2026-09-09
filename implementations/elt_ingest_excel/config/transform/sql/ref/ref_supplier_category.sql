DROP TABLE IF EXISTS ref_supplier_category
;
CREATE TABLE ref_supplier_category
    AS
SELECT
       TRIM(source_supplier_category)      source_supplier_category
     , TRIM(supplier_category)             supplier_category
  FROM read_json_auto(
           'config/data/ref_supplier_category.json'
       )
;
