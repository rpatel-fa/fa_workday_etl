DROP TABLE IF EXISTS validate_supplier_existing_in_delta
;
CREATE TABLE validate_supplier_existing_in_delta
    AS
  WITH cte_delta_vendor_id
    AS (
SELECT DISTINCT
       TRIM(t.vendor_id)                               vendor_id
  FROM src_fin_supplier_base                           t
 WHERE NULLIF(TRIM(t.vendor_id), '')                   IS NOT NULL
       )
SELECT DISTINCT
       w.*
  FROM fin_workday_extract_supplier                  w
     , UNNEST(STRING_SPLIT(w.reference_id, '|'))       u(reference_token)
 WHERE NULLIF(TRIM(w.reference_id), '')                IS NOT NULL
   AND NULLIF(TRIM(u.reference_token), '')             IS NOT NULL
   AND EXISTS (
       SELECT 1
         FROM cte_delta_vendor_id                       d
        WHERE UPPER(TRIM(u.reference_token))            = UPPER(d.vendor_id)
           OR UPPER(TRIM(u.reference_token))            LIKE UPPER(d.vendor_id) || '-%'
   )
;

DROP TABLE IF EXISTS validate_supplier_delta_in_existing
;
CREATE TABLE validate_supplier_delta_in_existing
    AS
  WITH cte_existing_reference_token
    AS (
SELECT DISTINCT
       UPPER(TRIM(u.reference_token))                   reference_token
  FROM fin_workday_extract_supplier                  w
     , UNNEST(STRING_SPLIT(w.reference_id, '|'))       u(reference_token)
 WHERE NULLIF(TRIM(w.reference_id), '')                IS NOT NULL
   AND NULLIF(TRIM(u.reference_token), '')             IS NOT NULL
       )
SELECT
       s.*
  FROM src_fin_supplier_base                           s
 WHERE NULLIF(TRIM(s.vendor_id), '')                   IS NOT NULL
   AND EXISTS (
       SELECT 1
         FROM cte_existing_reference_token              w
        WHERE w.reference_token                         = UPPER(TRIM(s.vendor_id))
           OR w.reference_token                         LIKE UPPER(TRIM(s.vendor_id)) || '-%'
   )
;

-- sheet: Existing in Delta
SELECT *
  FROM validate_supplier_existing_in_delta
 ORDER BY
       reference_id
;

-- sheet: Delta in Existing
SELECT *
  FROM validate_supplier_delta_in_existing
 ORDER BY
       vendor_id
;
