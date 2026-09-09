DROP TABLE IF EXISTS ref_country
;
CREATE TABLE ref_country
    AS
SELECT
       TRIM(country_code)              country_code
     , TRIM(language_code)             language_code
     , TRIM(currency_code)             currency_code
     , TRIM(phone_code)                phone_code
     , TRIM(tax_id_type)               tax_id_type
     , TRIM(country_name)              country_name
  FROM read_json_auto(
           'config/data/ref_country.json'
       )
;
