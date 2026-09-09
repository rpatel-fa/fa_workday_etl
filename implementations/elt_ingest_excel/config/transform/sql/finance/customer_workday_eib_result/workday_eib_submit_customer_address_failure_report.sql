DROP TABLE IF EXISTS workday_eib_submit_customer_address_failure_report
;
CREATE TABLE workday_eib_submit_customer_address_failure_report
    AS
  WITH cte_submit_customer_error
    AS (
SELECT t.*
     , e.error_description
  FROM src_fin_wd_submit_customer_address          t
       INNER JOIN
       src_fin_wd_submit_customer_address_error    e
          ON e.spreadsheet_key                     = t.spreadsheet_key
 WHERE regexp_matches(LOWER(e.error_description), 'postal|region|address|municipality|city|country')
       )
SELECT e.*
     , a.*
  FROM cte_submit_customer_error                   e
       INNER JOIN 
       workday_customer_name                       t
          ON t.reference_id                        = e.customer_reference_id
       LEFT OUTER JOIN
       workday_customer_address                    a
          ON a.customer_id                         = t.customer_id
;
