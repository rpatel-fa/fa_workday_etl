DROP TABLE IF EXISTS ref_source_payment_terms
;
CREATE TABLE ref_source_payment_terms
    AS
SELECT
       TRIM(source_payment_terms)           source_payment_terms
     , TRIM(workday_payment_terms)          workday_payment_terms
  FROM read_json_auto(
           'config/data/ref_source_payment_terms.json'
       )
;
