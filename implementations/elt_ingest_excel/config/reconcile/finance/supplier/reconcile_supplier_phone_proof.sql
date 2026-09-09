-- ============================================================================
-- RECONCILE: Supplier phone proof (row-level, APPROXIMATE)
-- One row per distinct raw phone/fax value aggregated in nrm_agg_phone_1/2/3
-- / nrm_agg_fax (src_fin_supplier), showing whether a matching number is
-- present in workday_supplier_phone for the same supplier_id.
--
-- CAVEAT: this is a best-effort approximate match, not an exact proof like
-- address/email. workday_supplier_phone.sql re-formats numbers heavily
-- (strips leading zeros, guesses/strips country codes, re-pads), so the raw
-- source digits will not literally equal the final phone_number string. This
-- check strips all non-digits from both sides and tests suffix AND prefix
-- containment in both directions (plus a last-9-digit match), which can
-- still produce false positives on coincidental digit overlap for short
-- numbers. Treat found_in_workday_supplier_phone_approx = 'N' as a strong
-- signal, but treat 'Y' as "likely present", not a certainty. The count-based
-- coverage in reconcile_supplier_phone_coverage.sql is the more reliable
-- aggregate guarantee; this table exists to give a per-number drill-down.
--
-- Prefix containment matters here specifically: a lot of supplier source
-- phone values are a valid number with garbage trailing zeros appended
-- (e.g. '75239943970000', 14 digits - likely an Excel numeric-precision
-- export artifact), and workday_supplier_phone.sql correctly recovers the
-- valid embedded number (e.g. '7523994397') by truncating the garbage tail -
-- the valid number is a PREFIX of the padded source value, not a suffix, so
-- a suffix-only match would wrongly flag this as not found even though the
-- transform handled it correctly.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_supplier_phone_proof
;
CREATE TABLE reconcile_supplier_phone_proof
    AS
  WITH cte_source_phones
    AS (
SELECT DISTINCT
       s.supplier_id                                                                            supplier_id
     , TRIM(u.phone)                                                                            phone_raw
     , REGEXP_REPLACE(TRIM(u.phone), '[^0-9]', '', 'g')                                         phone_digits
  FROM src_fin_supplier                                                                         s
     , UNNEST(STRING_SPLIT(
           COALESCE(s.nrm_agg_phone_1, '') || ';' ||
           COALESCE(s.nrm_agg_phone_2, '') || ';' ||
           COALESCE(s.nrm_agg_phone_3, '') || ';' ||
           COALESCE(s.nrm_agg_fax,     '')
       , ';'))                                                                                  u(phone)
 WHERE NULLIF(TRIM(u.phone), '')                                                                IS NOT NULL
       )
     , cte_target_phones
    AS (
SELECT supplier_id                                                                              supplier_id
     , REGEXP_REPLACE(phone_number, '[^0-9]', '', 'g')                                          phone_digits
  FROM workday_supplier_phone
 WHERE NULLIF(phone_number, '')                                                                 IS NOT NULL
       )
SELECT sp.supplier_id                                                                           supplier_id
     , sp.phone_raw                                                                             phone_raw
     , CASE
         WHEN EXISTS (
              SELECT 1
                FROM cte_target_phones                                                          tp
               WHERE tp.supplier_id                                                             = sp.supplier_id
                 AND (
                       RIGHT(tp.phone_digits, 9)                                                = RIGHT(sp.phone_digits, 9)
                    OR tp.phone_digits                                                          LIKE '%' || sp.phone_digits
                    OR sp.phone_digits                                                          LIKE '%' || tp.phone_digits
                    OR tp.phone_digits                                                          LIKE sp.phone_digits || '%'
                    OR sp.phone_digits                                                          LIKE tp.phone_digits || '%'
                     )
              )
         THEN 'Y'
         ELSE 'N'
       END                                                                                       found_in_workday_supplier_phone_approx
  FROM cte_source_phones                                                                        sp
 ORDER BY found_in_workday_supplier_phone_approx
     , supplier_id
;
