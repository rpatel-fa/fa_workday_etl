DROP TABLE IF EXISTS src_fin_customer_raw;

CREATE TABLE src_fin_customer_raw
    AS
  WITH cte_customer_src
    AS (
SELECT *
  FROM src_fin_customer_base 
       )
     , cte_customer_base
    AS (
SELECT DISTINCT
       TRIM(t.customer_number)                                                                        nrm_customer_number
     , UPPER(COALESCE(NULLIF(UPPER(TRIM(t.customer_name)), ''), NULLIF(TRIM(t.customer_number), ''))) nrm_customer_name_base
     , CASE
        WHEN UPPER(TRIM(t.post_code)) LIKE '%KNOWN%' OR TRIM(t.post_code) = '-'
        THEN NULL
        ELSE COALESCE(
             udf_normalize_postal_code(t.post_code, pc.postal_code_validation),
             udf_normalize_postal_code(
               NULLIF(
                 TRIM(
                   COALESCE(t.address_1, '') || ' '
                || COALESCE(t.address_2, '') || ' '
                || COALESCE(t.address_3, '') || ' '
                || COALESCE(t.city, '')      || ' '
                || COALESCE(t.county, '')
                 ),
                 ''
               ),
               pc.postal_code_validation
             ),
             -- Try to extract a full valid UK postcode (e.g. 'AL5 2LG', 'W1B 5TR', 'IM3-1AD')
             -- Handles comma and hyphen separators by replacing them with space first
             -- Normalizes multiple spaces to single space
             NULLIF(
                 REGEXP_REPLACE(
                     REGEXP_EXTRACT(UPPER(REGEXP_REPLACE(t.post_code, '[,.-]', ' ', 'g')), '([A-Z]{1,2}[0-9][0-9A-Z]?\s*[0-9][A-Z]{2})'),
                     '\s+', ' ', 'g'
                 ),
                 ''
             ),
             -- Fallback: Just take the first part (e.g. 'AL5' from 'AL5, Herts') if regex failed
             -- Also replace hyphen with space in the fallback to handle 'IM2- 1AD' if regex fails (though regex should catch it)
             NULLIF(TRIM(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(SPLIT_PART(t.post_code, ',', 1))), '[.-]', ' ', 'g'), '\s+', ' ', 'g')), '')
        )
       END                                                                                            nrm_postal_code_clean
     , t.*
  FROM cte_customer_src                      t
       LEFT JOIN LATERAL (
         SELECT country_code
           FROM ref_source_country_name_mapping
          WHERE UPPER(TRIM(source_country_name)) = UPPER(TRIM(t.country))
          LIMIT 1
       ) m_name
    ON TRUE
       LEFT JOIN LATERAL (
         SELECT country_code
           FROM ref_source_country_code_mapping
          WHERE UPPER(TRIM(source_country_code)) = NULLIF(UPPER(TRIM(t.country_code)), '')
          LIMIT 1
       ) m_code
    ON TRUE
       LEFT JOIN LATERAL (
         SELECT country_name
           FROM ref_country
          WHERE UPPER(TRIM(country_code)) = UPPER(TRIM(COALESCE(m_name.country_code, m_code.country_code, NULLIF(t.country_code, ''))))
          LIMIT 1
       ) rc
    ON TRUE
       LEFT JOIN LATERAL (
         SELECT postal_code_validation
           FROM ref_workday_validate_customer_postal_code
          WHERE LOWER(TRIM(country)) = LOWER(TRIM(COALESCE(rc.country_name, t.country)))
          LIMIT 1
       ) pc
    ON TRUE
       )
     , cte_customer_nrm
    AS (
SELECT
       CASE
         WHEN t.nrm_postal_code_clean IS NULL
         THEN NULL
         WHEN regexp_matches(
                REGEXP_REPLACE(UPPER(TRIM(t.nrm_postal_code_clean)), '\s', '', 'g'),
                '^(JE|GY|IM)[0-9]{2}[A-Z]{2}$'
              )
         THEN REGEXP_REPLACE(UPPER(TRIM(t.nrm_postal_code_clean)), '\s', '', 'g')
         WHEN regexp_matches(
                REGEXP_REPLACE(UPPER(TRIM(t.nrm_postal_code_clean)), '\s', '', 'g'),
                '^[A-Z]{1,2}[0-9][0-9A-Z]?[0-9][A-Z]{2}$'
              )
          AND REGEXP_EXTRACT(
                REGEXP_REPLACE(UPPER(TRIM(t.nrm_postal_code_clean)), '\s', '', 'g'),
                '^[A-Z]+'
              ) NOT IN ('JE', 'GY', 'IM')
         THEN REGEXP_REPLACE(
                REGEXP_REPLACE(UPPER(TRIM(t.nrm_postal_code_clean)), '\s', '', 'g'),
                '(.+)(.{3})$',
                '\1 \2'
              )
         ELSE REGEXP_REPLACE(UPPER(TRIM(t.nrm_postal_code_clean)), '\s+', ' ', 'g')
       END                                                                                            nrm_postal_code
       -- Clean Name: Remove dots, normalize abbreviations (FOOTBALL CLUB→FC, UNITED→UTD, AND→&),
       --             strip parenthesised qualifiers like (UK), then strip trailing legal suffixes
     , TRIM(REGEXP_REPLACE(
           REGEXP_REPLACE(
               REGEXP_REPLACE(
                   REGEXP_REPLACE(
                       REPLACE(
                            REPLACE(
                                REPLACE(
                                    REPLACE(UPPER(unaccent(t.nrm_customer_name_base)), '.', ''),
                                    'WOMEN''S', 'WOMEN'
                                ),
                                'WOMENS', 'WOMEN'
                            ),
                            'FOOTBALL CLUB', 'FC'
                        ),
                       '\s*\([^)]*\)', '', 'g'
                   ),
                   '\bUNITED\b', 'UTD', 'g'
               ),
               '\bAND\b', '&', 'g'
           ),
           '(\s+|^)((LIMITED|LTD|COMPANY|PLC|LLP|INC)\s*)+$', ''
       ))                                                                                             nrm_customer_name
     , COALESCE(
          TRY_STRPTIME(NULLIF(TRIM(t.created_date          ), ''), '%Y-%m-%d %H:%M:%S')
        , TRY_STRPTIME(NULLIF(TRIM(t.created_date          ), ''), '%Y-%m-%d')
        , TRY_STRPTIME(NULLIF(TRIM(t.created_date          ), ''), '%d-%m-%Y')
       )                                                                                              nrm_created_date
     , COALESCE(
          TRY_STRPTIME(NULLIF(TRIM(t.last_payment_date     ), ''), '%Y-%m-%d %H:%M:%S')
        , TRY_STRPTIME(NULLIF(TRIM(t.last_payment_date     ), ''), '%Y-%m-%d')
        , TRY_STRPTIME(NULLIF(TRIM(t.last_payment_date     ), ''), '%d-%m-%Y')
       )                                                                                              nrm_last_payment_date
     , COALESCE(
          TRY_STRPTIME(NULLIF(TRIM(t.last_transaction_date ), ''), '%Y-%m-%d %H:%M:%S')
        , TRY_STRPTIME(NULLIF(TRIM(t.last_transaction_date ), ''), '%Y-%m-%d')
        , TRY_STRPTIME(NULLIF(TRIM(t.last_transaction_date ), ''), '%d-%m-%Y')
       )                                                                                              nrm_last_transaction_date
     , mbu.target_value                                                                               nrm_business_unit
     , t.*
  FROM cte_customer_base                     t
       INNER JOIN
       ref_source_business_unit_mapping      mbu
          ON UPPER(mbu.source_value)         = UPPER(TRIM(t.business_unit))
       )
     , cte_customer_rank
    AS (
SELECT
       ROW_NUMBER() OVER
       (
         PARTITION BY
                   t.nrm_customer_name
                 , t.nrm_customer_number
             ORDER BY
                   t.nrm_created_date           ASC  NULLS LAST
                 , t.nrm_last_payment_date      DESC NULLS LAST
                 , t.nrm_last_transaction_date  DESC NULLS LAST
       )                                                           data_rnk
     , t.*
  FROM cte_customer_nrm                t
       )
     , cte_customer_addr_clean
    AS (
SELECT t.*
     , LIST_REVERSE(LIST_DISTINCT(LIST_FILTER([
         NULLIF(NULLIF(NULLIF(TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(t.address_1    , '[\"`<>|;{}]', '', 'g'), '\\s+', ' ', 'g'), ',+$', '')), '[Not Known]'), ''), '-')
       , NULLIF(NULLIF(NULLIF(TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(t.address_2    , '[\"`<>|;{}]', '', 'g'), '\\s+', ' ', 'g'), ',+$', '')), '[Not Known]'), ''), '-')
       , NULLIF(NULLIF(NULLIF(TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(t.address_3    , '[\"`<>|;{}]', '', 'g'), '\\s+', ' ', 'g'), ',+$', '')), '[Not Known]'), ''), '-')
       ], x -> x IS NOT NULL)))                                                           addr_unique_list
    FROM cte_customer_rank             t
       )
SELECT
       r.country_code                                                                     nrm_country_code
     , r.language_code                                                                    nrm_language_code
     , r.currency_code                                                                    nrm_currency_code
     , r.phone_code                                                                       nrm_phone_code
     , r.tax_id_type                                                                      nrm_tax_id_type
     , r.country_name                                                                     nrm_country_name
     , COALESCE(rco.target_value, rx.instance, rxo.instance)                              nrm_region
     , SPLIT_PART(
       CASE
        WHEN r0.post_town           IS NOT NULL
        THEN r0.post_town
        WHEN r3.town_city_name      IS NOT NULL
        THEN r3.town_city_name
        WHEN r4.town_city_name      IS NOT NULL
        THEN r4.town_city_name
        ELSE NULLIF(TRIM(t.city), '')
       END
       , ',', 1)                                                                          nrm_city
     , t.addr_unique_list[1]                                                              nrm_address_line_1
     , t.addr_unique_list[2]                                                              nrm_address_line_2
     , t.addr_unique_list[3]                                                              nrm_address_line_3
     , CAST(NULL AS STRING)                                                               nrm_address_line_4
       -- Candidate source for city_subdivision, restricted to the specific countries
       -- Workday actually requires it for (ZA/FI/SK/MX/AX) - not computed for anyone else,
       -- so address_3's contents for every other country are never touched here even as an
       -- unused candidate. Mirrors the supplier-side src_fin_supplier_raw.sql pattern.
       -- Whether it's actually exposed downstream is still decided separately in
       -- workday_customer_address_base.sql from ref_workday_validate_customer_address,
       -- keeping that decision table-driven rather than hardcoded. FI/SK/MX/AX default to
       -- raw Address Line 3, cleaned the same way as the address_3 branch of
       -- addr_unique_list above but kept as its own value (not run through that list's
       -- cross-field dedup/compaction, which would shift its position if address_2 were
       -- blank). One-time GP data migration: address_3 for Finland/Slovakia customers needs
       -- the subdivision/district value entered by finance/GP before this can populate - GP
       -- is being frozen after this migration, so this is a single data patch, not an
       -- ongoing field convention.
     , CASE
         WHEN r.country_code = 'ZA' THEN t.addr_unique_list[2]
         WHEN r.country_code IN ('FI', 'SK', 'MX', 'AX')
         THEN NULLIF(NULLIF(NULLIF(TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(t.address_3, '[\"`<>|;{}]', '', 'g'), '\\s+', ' ', 'g'), ',+$', '')), '[Not Known]'), ''), '-')
       END                                                                                nrm_city_subdivision
       -- Candidate source for subregion (Workday's "Region Subdivision 1"), sourced from the
       -- raw county field for every country - unconditionally, since county is already a
       -- county/subregion-level value regardless of country. Whether it's actually exposed
       -- downstream is decided separately in workday_customer_address_base.sql from
       -- ref_workday_validate_customer_address.subregion_county, keeping that decision
       -- table-driven rather than hardcoded to a country list. Mirrors the supplier-side fix.
     , NULLIF(TRIM(t.county), '')                                                        nrm_subregion_county
     , COALESCE(NULLIF(TRIM(UPPER(t.payment_terms_id)), ''), '30 DAYS')                   nrm_payment_terms_id
     , NULLIF(TRIM(UPPER(t.tax_schedule_id)), '')                                         nrm_tax_schedule_id
     , CASE
         WHEN REGEXP_REPLACE(TRIM(COALESCE(t.tax_registration_number, '')), '[^0-9A-Za-z]', '') ~ '^\d+$'
         THEN NULLIF(REGEXP_REPLACE(TRIM(COALESCE(t.tax_registration_number, '')), '[^0-9A-Za-z]', ''), '')
       END                                                                                nrm_tax_registration_number
     , COALESCE(NULLIF(TRIM(UPPER(t.address_code)), ''), 'MAIN')                          nrm_address_code
     , COALESCE(cc.target_value, 'Miscellaneous')                                         nrm_customer_category
     , CASE
         -- No postcode supplied — nothing to validate
         WHEN t.nrm_postal_code IS NULL
         THEN NULL
         -- Non-UK/CI country — foreign postcodes are out of scope, don't flag
         WHEN COALESCE(UPPER(TRIM(m_name.country_code)), UPPER(TRIM(m_code.country_code)), 'GB')
              NOT IN ('GB', 'JE', 'GG', 'IM')
         THEN NULL
         -- UK/CI address with a postcode — Y if Royal Mail district matched, N if not
         WHEN r0.postcode IS NOT NULL THEN 'Y'
         ELSE 'N'
       END                                                                                nrm_postal_code_valid
     , r0.postcode                                                                        nrm_postcode_prefix
     , r0.town_area                                                                       nrm_postcode_town
     , r0.region                                                                          nrm_postcode_region
     , t.*
 FROM cte_customer_addr_clean                         t
       -- Get the Royal Mail postcode for UK addresses
       LEFT OUTER JOIN
       ref_post_code_district                         r0
          ON (
               UPPER(TRIM(t.nrm_postal_code))         LIKE UPPER(TRIM(r0.postcode) || ' %')  -- Matches 'MK19 5AH' to 'MK19'
            OR UPPER(TRIM(t.nrm_postal_code))         =    UPPER(TRIM(r0.postcode))          -- Matches 'MK19'     to 'MK19'
             )
       -- First try: match on country name (higher population)
       LEFT OUTER JOIN
       ref_source_country_name_mapping                      m_name
          ON  UPPER(TRIM(m_name.source_country_name))       = UPPER(TRIM(t.country))
       -- Second try: match on country code (fallback)
       LEFT OUTER JOIN
       ref_source_country_code_mapping                      m_code
          ON  UPPER(TRIM(m_code.source_country_code))       = NULLIF(UPPER(TRIM(t.country_code)), '')
       -- Join to reference table using: name match > code match > default GB
       LEFT OUTER JOIN
       ref_country r
         ON UPPER(TRIM(r.country_code)) = CASE
                                             -- Valid district matched — use region
                                             WHEN r0.region = 'Jersey'        THEN 'JE'
                                             WHEN r0.region = 'Guernsey'      THEN 'GG'
                                             WHEN r0.region = 'Isle of Man'   THEN 'IM'
                                             WHEN r0.region IS NOT NULL       THEN 'GB'
                                             -- Invalid/unmatched district — fall back to area code
                                             -- only when source country is GB/null (i.e. could plausibly be UK/CI)
                                             WHEN r0.postcode IS NULL
                                             AND COALESCE(UPPER(TRIM(m_name.country_code)), UPPER(TRIM(m_code.country_code)), 'GB') IN ('GB', 'JE', 'GG', 'IM')
                                             AND REGEXP_EXTRACT(UPPER(TRIM(t.nrm_postal_code)), '^[A-Z]+') = 'JE' THEN 'JE'
                                             WHEN r0.postcode IS NULL
                                             AND COALESCE(UPPER(TRIM(m_name.country_code)), UPPER(TRIM(m_code.country_code)), 'GB') IN ('GB', 'JE', 'GG', 'IM')
                                             AND REGEXP_EXTRACT(UPPER(TRIM(t.nrm_postal_code)), '^[A-Z]+') = 'GY' THEN 'GG'
                                             WHEN r0.postcode IS NULL
                                             AND COALESCE(UPPER(TRIM(m_name.country_code)), UPPER(TRIM(m_code.country_code)), 'GB') IN ('GB', 'JE', 'GG', 'IM')
                                             AND REGEXP_EXTRACT(UPPER(TRIM(t.nrm_postal_code)), '^[A-Z]+') = 'IM' THEN 'IM'
                                             -- Everything else — use name/code mapping or default GB
                                             ELSE UPPER(TRIM(COALESCE(m_name.country_code, m_code.country_code, 'GB')))
                                          END
       -- Customer category normalization
       LEFT OUTER JOIN
       ref_customer_category                                cc
         ON UPPER(TRIM(cc.source_value))                    = COALESCE(NULLIF(UPPER(TRIM(t.customer_class)), ''), 'NULL')
       -- Address handling
       LEFT OUTER JOIN
       ref_post_code_workday_region                         r01
          ON UPPER(TRIM(r01.post_code_region))              = UPPER(TRIM(r0.region))
       LEFT OUTER JOIN
       ref_workday_country_state_region                     r1
         ON UPPER(TRIM(r1.country))                         = UPPER(TRIM(r.country_name))
        AND UPPER(TRIM(r1.instance))                        = UPPER(TRIM(t.county))
       LEFT OUTER JOIN
       ref_workday_country_state_region                     r2
         ON UPPER(TRIM(r2.country))                         = UPPER(TRIM(r.country_name))
        AND UPPER(TRIM(r2.instance))                        = UPPER(TRIM(t.city))
       LEFT OUTER JOIN
       ref_country_county_state_town_mapping                r3
         ON UPPER(TRIM(r3.country_code))                    = UPPER(TRIM(r.country_code))
        AND UPPER(TRIM(r3.town_city_name))                  = UPPER(TRIM(t.county))
       LEFT OUTER JOIN
       ref_country_county_state_town_mapping                r4
         ON UPPER(TRIM(r4.country_code))                    = UPPER(TRIM(r.country_code))
        AND UPPER(TRIM(r4.town_city_name))                  = UPPER(TRIM(t.city))
       LEFT OUTER JOIN
       ref_workday_country_state_region                     rx
         ON UPPER(TRIM(rx.country))                         = UPPER(TRIM(r.country_name))
        AND UPPER(TRIM(rx.instance))                        = UPPER(TRIM(
                                                              CASE
                                                               WHEN r0.region             IS NOT NULL THEN UPPER(TRIM(COALESCE(r01.workday_region, r0.region)))
                                                               WHEN r1.instance           IS NOT NULL THEN UPPER(TRIM(r1.instance))
                                                               WHEN r2.instance           IS NOT NULL THEN UPPER(TRIM(r2.instance))
                                                               WHEN r3.county_state_name  IS NOT NULL THEN UPPER(TRIM(r3.county_state_name))
                                                               WHEN r4.county_state_name  IS NOT NULL THEN UPPER(TRIM(r4.county_state_name))
                                                               ELSE NULLIF(UPPER(TRIM(t.county)), '')
                                                              END))
       LEFT OUTER JOIN
       ref_workday_country_state_region                     rxo
         ON UPPER(TRIM(rxo.country))                        = UPPER(TRIM(r.country_name))
        AND UPPER(TRIM(rxo.instance))                       = UPPER(TRIM(
                                                              CASE
                                                               WHEN r0.region             IS NOT NULL THEN UPPER(TRIM(COALESCE(r01.workday_region, r0.region)))
                                                               WHEN r1.instance           IS NOT NULL THEN UPPER(TRIM(r1.instance))
                                                               WHEN r2.instance           IS NOT NULL THEN UPPER(TRIM(r2.instance))
                                                               WHEN r3.county_state_name  IS NOT NULL THEN UPPER(TRIM(r3.county_state_name))
                                                               WHEN r4.county_state_name  IS NOT NULL THEN UPPER(TRIM(r4.county_state_name))
                                                              ELSE NULLIF(UPPER(TRIM(t.county)), '')
                                                              END || '(obsolete)'))
       LEFT OUTER JOIN
       ref_workday_county_obsolete                          rco
         ON UPPER(TRIM(rco.source_value))                   = UPPER(TRIM(COALESCE(rx.instance, rxo.instance)))

;
