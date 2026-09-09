-- ----------------------------------------------------------------------------
-- COVERAGE: customer_number from base must exist in src_fin_customer
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'src_fin_customer'                                                              table_name
     , 'COVERAGE'                                                                     rule_type
     , 'customer_number'                                                              col
     , CAST(NULL AS VARCHAR)                                                          customer_id
     , s.customer_number                                                              detail
     , 'customer_number from src_fin_customer_base is missing from src_fin_customer.' message
  FROM (
        SELECT DISTINCT
               TRIM(b.customer_number)                                                customer_number
          FROM src_fin_customer_base                                                  b
         WHERE NULLIF(TRIM(b.customer_number), '') IS NOT NULL
           AND NULLIF(TRIM(UPPER(b.payment_terms_id)), '') IS NOT NULL
       )                                                                              s
       LEFT OUTER JOIN (
        SELECT TRIM(UNNEST(STRING_SPLIT(nrm_agg_customer_number, '|')))               customer_number
          FROM src_fin_customer
       )                                                                              t
          ON t.customer_number = s.customer_number
 WHERE t.customer_number IS NULL
;

-- ----------------------------------------------------------------------------
-- COVERAGE: customer_number from base must appear in workday_customer_name
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_name'                                                              table_name
     , 'COVERAGE'                                                                          rule_type
     , 'customer_id'                                                                       col
     , t.customer_id                                                                       customer_id
     , s.customer_number                                                                   detail
     , 'customer_number from src_fin_customer_base is missing from workday_customer_name.' message
  FROM (
        SELECT DISTINCT
               TRIM(b.customer_number)                                                     customer_number
          FROM src_fin_customer_base                                                       b
         WHERE NULLIF(TRIM(b.customer_number), '') IS NOT NULL
           AND NULLIF(TRIM(UPPER(b.payment_terms_id)), '') IS NOT NULL
       )                                                                                   s
       LEFT OUTER JOIN (
        SELECT customer_id
             , TRIM(UNNEST(STRING_SPLIT(reference_id, '|')))                              customer_number
          FROM workday_customer_name
       )                                                                                   t
          ON t.customer_number = s.customer_number
 WHERE t.customer_id IS NULL
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: customer_number from base must map to exactly one row in src_fin_customer
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'src_fin_customer'                                                                     table_name
     , 'DUPLICATE'                                                                           rule_type
     , 'customer_number'                                                                     col
     , CAST(NULL AS VARCHAR)                                                                 customer_id
     , s.customer_number                                                                     detail
     , 'customer_number from src_fin_customer_base matches multiple rows in src_fin_customer.' message
  FROM (
        SELECT DISTINCT
               TRIM(b.customer_number)                                                       customer_number
          FROM src_fin_customer_base                                                         b
         WHERE NULLIF(TRIM(b.customer_number), '') IS NOT NULL
           AND NULLIF(TRIM(UPPER(b.payment_terms_id)), '') IS NOT NULL
       )                                                                                     s
       INNER JOIN (
        SELECT customer_id
             , TRIM(UNNEST(STRING_SPLIT(nrm_agg_customer_number, '|')))                     customer_number
          FROM src_fin_customer
       )                                                                                     t
          ON t.customer_number = s.customer_number
 GROUP BY s.customer_number
HAVING COUNT(DISTINCT t.customer_id) > 1
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: customer_number from base must map to exactly one row in workday_customer_name
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_name'                                                                    table_name
     , 'DUPLICATE'                                                                               rule_type
     , 'customer_number'                                                                         col
     , CAST(NULL AS VARCHAR)                                                                     customer_id
     , s.customer_number                                                                         detail
     , 'customer_number from src_fin_customer_base matches multiple rows in workday_customer_name.' message
  FROM (
        SELECT DISTINCT
               TRIM(b.customer_number)                                                           customer_number
          FROM src_fin_customer_base                                                             b
         WHERE NULLIF(TRIM(b.customer_number), '') IS NOT NULL
           AND NULLIF(TRIM(UPPER(b.payment_terms_id)), '') IS NOT NULL
       )                                                                                         s
       INNER JOIN (
        SELECT customer_id
             , TRIM(UNNEST(STRING_SPLIT(reference_id, '|')))                                    customer_number
          FROM workday_customer_name
       )                                                                                         t
          ON t.customer_number = s.customer_number
 GROUP BY s.customer_number
HAVING COUNT(DISTINCT t.customer_id) > 1
;
