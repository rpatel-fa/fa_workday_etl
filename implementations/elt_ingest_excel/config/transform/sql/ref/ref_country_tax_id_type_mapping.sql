DROP TABLE IF EXISTS ref_country_tax_id_type_mapping
;
CREATE TABLE ref_country_tax_id_type_mapping
    AS
SELECT
       TRIM(country_code)                     country_code
     , TRIM(tax_id_type_label)                tax_id_type_label
     , CAST(is_default AS BOOLEAN)            is_default
  FROM read_json_auto(
           'config/data/ref_country_tax_id_type_mapping.json'
       )
;
