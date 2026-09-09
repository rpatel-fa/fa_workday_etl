DROP TABLE IF EXISTS validate_customer_existing_in_delta
;
CREATE TABLE validate_customer_existing_in_delta
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
           OR UPPER(TRIM(u.reference_token))             LIKE UPPER(d.customer_number) || '-%'
   )
;

DROP TABLE IF EXISTS validate_customer_delta_in_existing
;
CREATE TABLE validate_customer_delta_in_existing
    AS
  WITH cte_existing_reference_token
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
         FROM cte_existing_reference_token               w
        WHERE w.reference_token                          = UPPER(TRIM(s.customer_number))
           OR w.reference_token                          LIKE UPPER(TRIM(s.customer_number)) || '-%'
   )
;

-- sheet: Existing in Delta
SELECT *
  FROM validate_customer_existing_in_delta
 ORDER BY
       reference_id
;

-- sheet: Delta in Existing
SELECT *
  FROM validate_customer_delta_in_existing
 ORDER BY
       customer_number
;
