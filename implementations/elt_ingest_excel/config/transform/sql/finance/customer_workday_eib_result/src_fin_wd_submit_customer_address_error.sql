DROP TABLE IF EXISTS src_fin_wd_submit_customer_address_error
;
CREATE TABLE src_fin_wd_submit_customer_address_error
    AS
SELECT DISTINCT
       a.spreadsheet_key
     , s.error_description
  FROM src_fin_wd_submit_customer_address                             a
       INNER JOIN
       fin_wd_submit_customer_summary                                 s
           ON TRY_CAST(NULLIF(TRIM(s.spreadsheet_key), '') AS INTEGER) = a.spreadsheet_key
 WHERE NULLIF(TRIM(COALESCE(s.error_description, '')), '')            IS NOT NULL
;
