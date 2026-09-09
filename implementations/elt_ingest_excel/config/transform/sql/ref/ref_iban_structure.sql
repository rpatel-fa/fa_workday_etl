DROP TABLE IF EXISTS ref_iban_structure
;
CREATE TABLE ref_iban_structure
    AS
SELECT
       TRIM(country_code)                   country_code
     , CAST(iban_length AS INTEGER)         iban_length
     , CAST(bank_code_offset AS INTEGER)    bank_code_offset
     , CAST(bank_code_length AS INTEGER)    bank_code_length
     , TRIM("source")                       source_ref
     , TRIM(notes)                          notes
  FROM read_json_auto(
           'config/data/ref_iban_structure.json'
       )
;
