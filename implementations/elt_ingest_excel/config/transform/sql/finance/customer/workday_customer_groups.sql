DROP TABLE IF EXISTS workday_customer_groups
;
CREATE TABLE workday_customer_groups
    AS
  WITH cte_group
    AS (
SELECT
       TRIM(c.customer_id)                            customer_id
     , c.nrm_customer_name                            customer_name
     , CASE
           WHEN UPPER(SUBSTR(TRIM(c.nrm_customer_name), 1, 1)) BETWEEN 'A' AND 'L'
               THEN 'Customers A - L'
           WHEN UPPER(SUBSTR(TRIM(c.nrm_customer_name), 1, 1)) BETWEEN 'M' AND 'Z'
               THEN 'Customers M - Z'
           ELSE 'Customers 0 - 9'
       END                                            customer_group
  FROM src_fin_customer                               c
UNION ALL
SELECT
       TRIM(c.customer_id)                            customer_id
     , c.nrm_customer_name                            customer_name
     , r.target_value                                 customer_group
  FROM src_fin_customer                               c
       INNER JOIN
       ref_workday_group_additional                   r
          ON r.source_type                            = 'customer'
         AND UPPER(TRIM(r.source_value))              = UPPER(TRIM(c.customer_class))
    )
SELECT *
  FROM cte_group
 ORDER BY
       customer_id
;
