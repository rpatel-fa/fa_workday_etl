DROP TABLE IF EXISTS ref_bank_sort_code_prefix_mapping
;
CREATE TABLE ref_bank_sort_code_prefix_mapping
    AS
SELECT
       TRIM(sort_code_prefix)          sort_code_prefix
     , TRIM(bank_name_primary)         bank_name_primary
     , TRIM(banking_group)             banking_group
     , TRIM(prefix_type)               prefix_type
     , TRIM(notes)                     notes
  FROM read_json_auto(
           'config/data/ref_bank_sort_code_prefix_mapping.json'
       )
;
