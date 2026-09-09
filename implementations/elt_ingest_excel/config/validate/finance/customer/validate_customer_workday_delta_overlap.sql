DROP TABLE IF EXISTS validate_customer_workday_in_delta
;
CREATE TABLE validate_customer_workday_in_delta
    AS
  WITH cte_delta_customer_number
    AS (
SELECT DISTINCT
       TRIM(t.customer_number)                          customer_number
  FROM src_fin_customer_base                            t
 WHERE NULLIF(TRIM(t.customer_number), '')              IS NOT NULL
   AND NULLIF(TRIM(UPPER(t.payment_terms_id)), '')      IS NOT NULL
       )
SELECT DISTINCT
       w.*
  FROM fin_workday_extract_customer                     w
     , UNNEST(STRING_SPLIT(w.reference_id, '|'))        u(reference_token)
 WHERE NULLIF(TRIM(w.reference_id), '')                 IS NOT NULL
   AND NULLIF(TRIM(u.reference_token), '')              IS NOT NULL
   AND EXISTS (
       SELECT 1
         FROM cte_delta_customer_number                  d
        WHERE UPPER(TRIM(u.reference_token))             = UPPER(d.customer_number)
   )
;

DROP TABLE IF EXISTS validate_customer_delta_in_workday
;
CREATE TABLE validate_customer_delta_in_workday
    AS
  WITH cte_workday_reference_token
    AS (
SELECT DISTINCT
       UPPER(TRIM(u.reference_token))                   reference_token
  FROM fin_workday_extract_customer                     w
     , UNNEST(STRING_SPLIT(w.reference_id, '|'))        u(reference_token)
 WHERE NULLIF(TRIM(w.reference_id), '')                 IS NOT NULL
   AND NULLIF(TRIM(u.reference_token), '')              IS NOT NULL
       )
SELECT
       s.*
  FROM src_fin_customer_base                            s
 WHERE NULLIF(TRIM(s.customer_number), '')              IS NOT NULL
   AND NULLIF(TRIM(UPPER(s.payment_terms_id)), '')      IS NOT NULL
   AND EXISTS (
       SELECT 1
         FROM cte_workday_reference_token                w
        WHERE w.reference_token                          = UPPER(TRIM(s.customer_number))
   )
;

-- sheet: Workday in Delta
SELECT *
  FROM validate_customer_workday_in_delta
 ORDER BY
       reference_id
;

-- sheet: Delta in Workday
SELECT *
  FROM validate_customer_delta_in_workday
 ORDER BY
       customer_number
;
