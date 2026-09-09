DROP TABLE IF EXISTS src_fin_workday_extract_ccs_customer
;
CREATE TABLE src_fin_workday_extract_ccs_customer
    AS
SELECT
       t.*
  FROM fin_workday_extract_customer    t
;
