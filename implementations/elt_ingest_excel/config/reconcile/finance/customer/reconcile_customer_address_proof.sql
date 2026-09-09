-- ============================================================================
-- RECONCILE: Customer address proof (row-level)
-- One row per distinct (customer_id, address) captured in the aggregated
-- nrm_array_customer_address (src_fin_customer), showing whether that
-- address is present in workday_customer_address_base.
--
-- workday_customer_address_base.sql deliberately NULLs address_line_2/3/4,
-- city and postal_code per country, driven by ref_workday_validate_customer_address
-- (e.g. address_line_3/4 = 'Not Accepted' for the US, Canada, China; postal_code
-- and municipality_city = 'Not Accepted' for the UAE) - Workday's address format
-- simply has no such field for that country. This is correct behaviour, not
-- data loss, so this proof exempts a field from comparison whenever the
-- country's rule says that field is 'Not Accepted'. The join to
-- ref_workday_validate_customer_address is on COUNTRY NAME, matching
-- workday_customer_address_base.sql exactly - country_code cannot be used
-- here, the ref table uses 3-letter ISO codes (USA/CAN/CHN) while the source
-- uses 2-letter codes (US/CA/CN). Hungary has an additional address_line_1/2
-- splitting/truncation rule (workday_customer_address_base.sql lines 44-62)
-- that is excluded from strict comparison for the same reason.
--
-- address_line_1 is always required by Workday and is never suppressed, so it
-- is always compared exactly (Hungary excepted).
-- ============================================================================

DROP TABLE IF EXISTS reconcile_customer_address_proof
;
CREATE TABLE reconcile_customer_address_proof
    AS
  WITH cte_source_address
    AS (
SELECT DISTINCT
       c.customer_id                                                                            customer_id
     , u.address.nrm_country_code                                                               country_code
     , u.address.nrm_country_name                                                               country_name
     , u.address.nrm_address_line_1                                                             address_line_1
     , u.address.nrm_address_line_2                                                             address_line_2
     , u.address.nrm_address_line_3                                                             address_line_3
     , u.address.nrm_address_line_4                                                             address_line_4
     , u.address.nrm_city                                                                       city
     , u.address.nrm_postal_code                                                                postal_code
  FROM src_fin_customer                                                                         c
     , UNNEST(c.nrm_array_customer_address)                                                     u(address)
 WHERE COALESCE(
           NULLIF(TRIM(u.address.nrm_address_line_1), '')
         , NULLIF(TRIM(u.address.nrm_address_line_2), '')
         , NULLIF(TRIM(u.address.nrm_address_line_3), '')
         , NULLIF(TRIM(u.address.nrm_address_line_4), '')
       )                                                                                        IS NOT NULL
       )
     , cte_target_address
    AS (
SELECT DISTINCT
       customer_id
     , address_line_1
     , address_line_2
     , address_line_3
     , address_line_4
     , city
     , postal_code
  FROM workday_customer_address_base
       )
SELECT s.customer_id                                                                            customer_id
     , s.address_line_1                                                                         address_line_1
     , s.country_name                                                                           country_name
     , CASE WHEN t.customer_id IS NOT NULL THEN 'Y' ELSE 'N' END                                found_in_workday_customer_address_base
  FROM cte_source_address                                                                       s
  LEFT JOIN ref_workday_validate_customer_address                                                req
         ON LOWER(TRIM(req.country))                                                            = LOWER(TRIM(s.country_name))
  LEFT JOIN cte_target_address                                                                   t
         ON  t.customer_id                                                                      = s.customer_id
        AND (s.country_code = 'HU' OR t.address_line_1 IS NOT DISTINCT FROM s.address_line_1)
        AND (s.country_code = 'HU'
             OR LOWER(TRIM(req.address_line_2))   = 'not accepted'
             OR LOWER(TRIM(req.postal_code))      = 'not accepted'
             OR t.address_line_2 IS NOT DISTINCT FROM s.address_line_2)
        AND (LOWER(TRIM(req.address_line_3)) = 'not accepted' OR t.address_line_3 IS NOT DISTINCT FROM s.address_line_3)
        AND (LOWER(TRIM(req.address_line_4)) = 'not accepted' OR t.address_line_4 IS NOT DISTINCT FROM s.address_line_4)
        AND (LOWER(TRIM(req.municipality_city)) = 'not accepted' OR t.city IS NOT DISTINCT FROM s.city)
        AND (LOWER(TRIM(req.postal_code)) = 'not accepted' OR t.postal_code IS NOT DISTINCT FROM s.postal_code)
 ORDER BY found_in_workday_customer_address_base
     , customer_id
;
