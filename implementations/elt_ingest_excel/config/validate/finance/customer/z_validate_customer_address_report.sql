DROP TABLE IF EXISTS validate_customer_address_report
;
CREATE TABLE validate_customer_address_report 
    AS
SELECT s1.nrm_agg_customer_number   gp_customer_number
     , s1.nrm_agg_business_unit     gp_business_unit
     , w.customer_id
     , w.customer_name
     , w.country
     , w.country_code
     , w.region
     , w.city
     , w.address_line_1
     , w.address_line_2
     , w.address_line_3
     , w.address_line_4
     , w.postal_code
     , w.is_primary
  FROM workday_customer_address w
       INNER JOIN
       src_fin_customer s1
        ON s1.customer_id = w.customer_id
 WHERE w.customer_id IN (
       SELECT customer_id FROM validation_customer_issues WHERE table_name = 'workday_customer_address'
       )
ORDER BY 1
;

DROP TABLE IF EXISTS validate_customer_address_report_base
;
CREATE TABLE validate_customer_address_report_base
    AS
SELECT company
     , customer_number
     , customer_name
     , address_code
     , address_1
     , address_2
     , address_3
     , city
     , county
     , post_code
     , country_code
     , country
  FROM src_fin_customer_base b
       INNER JOIN (
SELECT DISTINCT
       TRIM(u.gp_customer_number) gp_customer_number
  FROM validate_customer_address_report t
     , UNNEST(string_split(t.gp_customer_number, '|')) u(gp_customer_number)
 WHERE NULLIF(TRIM(u.gp_customer_number), '') IS NOT NULL
       ) v
        ON TRIM(b.customer_number) = v.gp_customer_number
 ORDER BY
       country_code
     , customer_number
