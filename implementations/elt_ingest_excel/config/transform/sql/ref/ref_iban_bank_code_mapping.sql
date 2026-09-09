DROP TABLE IF EXISTS ref_iban_bank_code_mapping
;
CREATE TABLE ref_iban_bank_code_mapping
    AS
SELECT
       TRIM(country_code)               country_code
     , TRIM(bank_code)                  bank_code
     , TRIM(bank_name)                  bank_name
     , TRIM("source")                   source_ref
     , TRIM(notes)                      notes
  FROM read_json_auto(
           'config/data/ref_iban_bank_code_mapping.json'
       )
;
