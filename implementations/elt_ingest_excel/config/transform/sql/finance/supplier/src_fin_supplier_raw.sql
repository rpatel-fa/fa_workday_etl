DROP TABLE IF EXISTS src_fin_supplier_raw
;
CREATE TABLE src_fin_supplier_raw
    AS
  WITH cte_supplier_src
    AS (
SELECT *
  FROM src_fin_supplier_base
       )
     , cte_supplier_base
    AS (
SELECT DISTINCT
       TRIM(t.vendor_id)                                                                              nrm_vendor_id
     , UPPER(COALESCE(NULLIF(UPPER(TRIM(t.vendor_name)), ''), NULLIF(TRIM(t.vendor_id), '')))         nrm_supplier_name_base
     , CASE
        WHEN UPPER(TRIM(t.post_code)) LIKE '%KNOWN%' OR TRIM(t.post_code) = '-'
        THEN NULL
        ELSE COALESCE(
             udf_normalize_postal_code(t.post_code, pc.postal_code_validation),
             udf_normalize_postal_code(
               NULLIF(
                 TRIM(
                   COALESCE(t.address_1, '')     || ' '
                || COALESCE(t.address_2, '')     || ' '
                || COALESCE(t.address_3, '')     || ' '
                || COALESCE(t.addressline_4, '') || ' '
                || COALESCE(t.city, '')          || ' '
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
  FROM cte_supplier_src                t
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
           FROM ref_workday_validate_supplier_postal_code
          WHERE LOWER(TRIM(country)) = LOWER(TRIM(COALESCE(rc.country_name, t.country)))
          LIMIT 1
       ) pc
    ON TRUE
       )
     , cte_supplier_nrm
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
                       REGEXP_REPLACE(
                           REPLACE(
                                REPLACE(
                                    REPLACE(
                                        REPLACE(UPPER(unaccent(t.nrm_supplier_name_base)), '.', ''),
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
           ),
           '\s+', ' ', 'g'
       ))                                                                                             nrm_supplier_name
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
          TRY_STRPTIME(NULLIF(TRIM(t.last_purchase_date    ), ''), '%Y-%m-%d %H:%M:%S')
        , TRY_STRPTIME(NULLIF(TRIM(t.last_purchase_date    ), ''), '%Y-%m-%d')
        , TRY_STRPTIME(NULLIF(TRIM(t.last_purchase_date    ), ''), '%d-%m-%Y')
       )                                                                                              nrm_last_purchase_date
     , mbu.target_value                                                                               nrm_business_unit
     , t.*
  FROM cte_supplier_base                     t
       INNER JOIN
       ref_source_business_unit_mapping      mbu
          ON UPPER(mbu.source_value)         = UPPER(TRIM(t.business_unit))
       )
     , cte_supplier_id
    AS (
-- GP vendor_id is scoped per business_unit in the source system, not globally unique -
-- the same code is independently reused for unrelated vendors in different company
-- books (e.g. 'HAL003' = three different companies across FA/NFC/WNSL). Detected by
-- the fully-normalised nrm_supplier_name disagreeing for the same vendor_id: when it
-- does, suffix the business_unit onto the key so each real vendor gets its own
-- canonical supplier number instead of being merged into one composite record.
SELECT t.*
     , CASE
         WHEN COUNT(DISTINCT t.nrm_supplier_name) OVER (PARTITION BY t.nrm_vendor_id) > 1
         THEN t.nrm_vendor_id || '-' || t.business_unit
         ELSE t.nrm_vendor_id
       END                                                                                            nrm_supplier_number
  FROM cte_supplier_nrm                       t
       )
     , cte_supplier_addr_clean
    AS (
SELECT t.*
     , LIST_REVERSE(LIST_DISTINCT(LIST_FILTER([
         NULLIF(NULLIF(NULLIF(TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(t.address_1    , '[\"`<>|;{}]', '', 'g'), '\\s+', ' ', 'g'), ',+$', '')), '[Not Known]'), ''), '-')
       , NULLIF(NULLIF(NULLIF(TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(t.address_2    , '[\"`<>|;{}]', '', 'g'), '\\s+', ' ', 'g'), ',+$', '')), '[Not Known]'), ''), '-')
       , NULLIF(NULLIF(NULLIF(TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(t.address_3    , '[\"`<>|;{}]', '', 'g'), '\\s+', ' ', 'g'), ',+$', '')), '[Not Known]'), ''), '-')
       , NULLIF(NULLIF(NULLIF(TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(t.addressline_4, '[\"`<>|;{}]', '', 'g'), '\\s+', ' ', 'g'), ',+$', '')), '[Not Known]'), ''), '-')
       ], x -> x IS NOT NULL)))                                                           addr_unique_list
  FROM cte_supplier_id                       t
       )
     , cte_supplier_clean
    AS (
SELECT
       r.country_code                                                                     nrm_country_code
     , r.language_code                                                                    nrm_language_code
     , r.currency_code                                                                    nrm_currency_code
     , r.phone_code                                                                       nrm_phone_code
     , r.tax_id_type                                                                      nrm_tax_id_type
     , r.country_name                                                                     nrm_country_name
     , COALESCE(scm.supplier_category, 'Miscellaneous')                                   nrm_supplier_category
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
     , t.addr_unique_list[4]                                                              nrm_address_line_4
       -- Candidate source for city_subdivision, restricted to the specific countries
       -- Workday actually requires it for (ZA/FI/SK/MX/AX) - not computed for anyone else,
       -- so address_3's contents for every other country are never touched here even as an
       -- unused candidate. Whether it's actually exposed downstream is still decided
       -- separately in workday_supplier_address_base.sql from
       -- ref_workday_validate_supplier_address, keeping that decision table-driven rather
       -- than hardcoded. South Africa already has real data correctly in address_2 (e.g.
       -- 'Allen's Nek', a Johannesburg suburb) - kept as-is, no rework needed. FI/SK/MX/AX
       -- default to raw Address Line 3, cleaned the same way as the address_3 branch of
       -- addr_unique_list above but kept as its own value (not run through that list's
       -- cross-field dedup/compaction, which would shift its position if address_2 were
       -- blank). One-time GP data migration: address_3 for the Finland/Slovakia suppliers
       -- needs the subdivision/district value entered by finance/GP before this can
       -- populate - GP is being frozen after this migration, so this is a single data
       -- patch, not an ongoing field convention.
     , CASE
         WHEN r.country_code = 'ZA' THEN t.addr_unique_list[2]
         WHEN r.country_code IN ('FI', 'SK', 'MX', 'AX')
         THEN NULLIF(NULLIF(NULLIF(TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(t.address_3, '[\"`<>|;{}]', '', 'g'), '\\s+', ' ', 'g'), ',+$', '')), '[Not Known]'), ''), '-')
       END                                                                                nrm_city_subdivision
       -- Candidate source for subregion (Workday's "Region Subdivision 1"), sourced from the
       -- raw county field for every country - unconditionally, since county is already a
       -- county/subregion-level value regardless of country. Whether it's actually exposed
       -- downstream is decided separately in workday_supplier_address_base.sql from
       -- ref_workday_validate_supplier_address.subregion_county, keeping that decision
       -- table-driven rather than hardcoded to a country list.
     , NULLIF(TRIM(t.county), '')                                                        nrm_subregion_county
     , COALESCE(NULLIF(TRIM(UPPER(t.vendor_address_code_primary)), ''), 'MAIN')           nrm_address_code
     , NULLIF(UPPER(TRIM(t.tax_schedule_id)), '')                                         nrm_tax_schedule_id
     , NULLIF(UPPER(TRIM(t.tax_id_number  )), '')                                         nrm_tax_id_number
     , CASE
         WHEN REGEXP_REPLACE(TRIM(COALESCE(t.tax_registration_number, '')), '[^0-9A-Za-z]', '') ~ '^\d+$'
         THEN NULLIF(REGEXP_REPLACE(TRIM(COALESCE(t.tax_registration_number, '')), '[^0-9A-Za-z]', ''), '')
       END                                                                                nrm_tax_registration_number
     , pt.workday_payment_terms                                                           nrm_payment_terms_id
     , COALESCE(
           -- 1. US: tenant-authoritative routing -> bank name, from
           --    ref_bank_routing_number_mapping (the tenant's own AT/AU named range), so the
           --    name matches the exact spelling the EIB expects.
           rbrn.bank_name
           -- 2. UK sort-code prefix -> a CONSISTENT canonical bank name (the reason this
           --    derivation is preferred over GP's inconsistent free-text). Skips the shared/
           --    fintech prefixes whose ref value is only a "Various …" / "range owner"
           --    placeholder (04, 23, 70-72, 93) rather than a specific bank.
         , CASE WHEN LOWER(COALESCE(rbsc.bank_name_primary, '')) NOT LIKE '%various%'
                 AND LOWER(COALESCE(rbsc.bank_name_primary, '')) NOT LIKE '%range owner%'
                THEN rbsc.bank_name_primary
           END
           -- 3. GP's own bank name (real, not the "Bank Name" placeholder) - supplies the
           --    ambiguous 04/23 accounts where GP recorded the actual bank (e.g. MONZO).
         , CASE WHEN UPPER(TRIM(COALESCE(t.bank_name, ''))) NOT IN ('', 'BANK NAME')
                THEN TRIM(t.bank_name)
           END
           -- 4. IBAN-embedded UK sort code (same placeholder skip) / international IBAN name.
         , CASE WHEN LOWER(COALESCE(rbsc_iban.bank_name_primary, '')) NOT LIKE '%various%'
                 AND LOWER(COALESCE(rbsc_iban.bank_name_primary, '')) NOT LIKE '%range owner%'
                THEN rbsc_iban.bank_name_primary
           END
         , iban_intl.bank_name
           -- 5. Last resort: the prefix's own placeholder ("Various …"), so a valid
           --    sort+account still ships rather than being dropped over a cosmetic name.
         , rbsc.bank_name_primary
       )                                                                                  nrm_bank_name
       -- GP never populates bank_country_code/central_bank_code (confirmed 0% across
       -- all records with a bank account) - there is no source-provided bank country.
       -- A sort-code match (rbsc/rbsc_iban for GB, iban_intl for other countries via
       -- ref_iban_bank_code_mapping) is treated as evidence of that account's actual
       -- country regardless of the supplier's own registered country (a supplier can
       -- genuinely hold a bank account in a different country); NULL here (no match)
       -- means "no evidence either way", and the supplier's own country is used as
       -- the default in workday_supplier_settlement_account.sql, not asserted here.
     , COALESCE(
           CASE WHEN rbsc.sort_code_prefix IS NOT NULL
             OR   rbsc_iban.sort_code_prefix IS NOT NULL
                THEN 'United Kingdom'
           END
         , iban_intl.country_name
       )                                                                                  nrm_bank_country_name
     , NULLIF(TRIM(t.eft_bank_code), '')                                                  nrm_bank_sort_code
     , NULLIF(COALESCE(NULLIF(TRIM(t.vendor_check_name), ''), t.nrm_supplier_name), '')   nrm_bank_account_name
     , COALESCE(
       CASE
         -- UK account numbers are 8 digits; a 7-digit value is a dropped leading
         -- zero from numeric storage - pad it (same cause as the sort code above).
         -- Only the exactly-7 case is padded: 6-digit (could be a legacy/building-
         -- society account) and 9/10-digit (roll-number concatenation) are left
         -- as-is so the account-length validation check surfaces them for review.
         WHEN r.country_code = 'GB'
          AND regexp_matches(REPLACE(TRIM(COALESCE(t.eft_bank_account, '')), ' ', ''), '^[0-9]{7}$')
         THEN '0' || REPLACE(TRIM(t.eft_bank_account), ' ', '')
         ELSE NULLIF(TRIM(REPLACE(t.eft_bank_account, ' ', '')), '')
       END
       -- Fill-blank-only recovery: GP frequently stored the bank account in its IBAN
       -- column instead of EFT Bank Account (confirmed against source - e.g. US 2TH001,
       -- GB AIR003, SG AIR010). Workday requires the local Financial Account Number for
       -- non-SEPA / local-clearing countries (US, GB, SG, JP, AU, ...), so recover it
       -- here rather than shipping an empty account. COALESCE means a real GP-supplied
       -- account is never overwritten; GP source columns are read only.
     , CASE
         -- Structured GB IBAN (CCkk BBBB SSSSSS NNNNNNNN, 22 chars): the 8-digit
         -- Financial Account Number is the final 8 characters (chars 15-22).
         WHEN regexp_matches(REPLACE(REPLACE(UPPER(TRIM(COALESCE(t.iban, ''))), ' ', ''), '-', ''), '^GB[0-9]{2}[A-Z]{4}[0-9]{14}$')
         THEN SUBSTR(REPLACE(REPLACE(UPPER(TRIM(t.iban)), ' ', ''), '-', ''), 15, 8)
         -- Not an ISO 13616 IBAN (no leading 2-letter country + 2 check digits): GP used
         -- the IBAN column to hold the raw local account number for a non-IBAN country -
         -- use it directly as the Financial Account Number.
         WHEN NULLIF(TRIM(t.iban), '') IS NOT NULL
          AND NOT regexp_matches(REPLACE(REPLACE(UPPER(TRIM(t.iban)), ' ', ''), '-', ''), '^[A-Z]{2}[0-9]{2}')
         THEN REPLACE(REPLACE(UPPER(TRIM(t.iban)), ' ', ''), '-', '')
       END
       )                                                                                  nrm_bank_account_number
     , COALESCE(
       CASE
         WHEN r.country_code = 'GB'
         THEN CASE
                -- UK sort codes are 6 numeric digits. Strip stray formatting first -
                -- space, hyphen, and an Excel text-marker apostrophe (e.g. "040003'")
                -- - then pad a dropped leading zero (5 -> 6; padded prefix 00-09 all
                -- exist in ref_bank_sort_code_prefix_mapping). Letters are kept, so a
                -- BIC wrongly placed in this field (e.g. 'MIDLGB22') stays visible and
                -- flagged rather than silently mangled.
                WHEN regexp_matches(REPLACE(REPLACE(REPLACE(TRIM(COALESCE(t.eft_bank_code, '')), ' ', ''), '-', ''), '''', ''), '^[0-9]{5}$')
                THEN '0' || REPLACE(REPLACE(REPLACE(TRIM(t.eft_bank_code), ' ', ''), '-', ''), '''', '')
                ELSE NULLIF(REPLACE(REPLACE(REPLACE(TRIM(t.eft_bank_code), ' ', ''), '-', ''), '''', ''), '')
              END
         WHEN r.country_code IN ('IE', 'FR', 'DE', 'ES', 'IT', 'NL', 'BE', 'PT', 'AT', 'SE', 'DK', 'FI')
         THEN NULLIF(TRIM(REPLACE(REPLACE(t.eft_bank_code, ' ', ''), '-', '')), '')
         WHEN r.country_code = 'US'
         THEN COALESCE(
                NULLIF(TRIM(REPLACE(REPLACE(t.eft_transit_routing_no, ' ', ''), '-', '')), '')
              , NULLIF(TRIM(REPLACE(REPLACE(t.eft_bank_code, ' ', ''), '-', '')), '')
              )
         ELSE NULLIF(TRIM(REPLACE(REPLACE(t.eft_bank_code, ' ', ''), '-', '')), '')
       END
       -- Fill-blank-only fallback: the national bank code that ISO 13616 embeds in the
       -- supplier's own IBAN (positions per ref_iban_structure). COALESCE means a real
       -- GP-supplied bank/routing code is never overwritten; the join's iban_length
       -- guard means only a structurally complete IBAN yields a code (a truncated or
       -- placeholder IBAN yields nothing, not a mangled value). Source columns are read
       -- only - GP data is never altered.
     , iban_bankcode.bank_code
       -- Fill-blank-only recovery: the 6-digit UK sort code embedded in a structured GB
       -- IBAN (chars 9-14) when GP left EFT Bank Code blank and stored the details only in
       -- the IBAN column. Workday requires the Routing/Sort field for GB; without this the
       -- record fails "Routing Transit or Institution Number is required".
     , CASE
         WHEN regexp_matches(REPLACE(REPLACE(UPPER(TRIM(COALESCE(t.iban, ''))), ' ', ''), '-', ''), '^GB[0-9]{2}[A-Z]{4}[0-9]{14}$')
         THEN SUBSTR(REPLACE(REPLACE(UPPER(TRIM(t.iban)), ' ', ''), '-', ''), 9, 6)
       END
       )                                                                                  nrm_bank_code_routing_number
       -- Only keep a value that is structurally an ISO 13616 IBAN (2-letter country + 2
       -- check digits). GP frequently stored a raw local account number in this column for
       -- non-IBAN countries (US/SG/JP/...); that value is recovered into
       -- nrm_bank_account_number above and must not also ship as a fabricated IBAN, which
       -- Workday would reject on format.
     , CASE
         WHEN regexp_matches(REPLACE(REPLACE(UPPER(TRIM(COALESCE(t.iban, ''))), ' ', ''), '-', ''), '^[A-Z]{2}[0-9]{2}[A-Z0-9]+$')
         THEN NULLIF(TRIM(REPLACE(REPLACE(UPPER(t.iban), ' ', ''), '-', '')), '')
       END                                                                                nrm_bank_iban
       -- A genuine SWIFT/BIC code is 8-11 alphanumeric characters (bank + country +
       -- location [+ branch]) - it is never just a bare sort code/routing number.
       -- Confirmed real cases where swift_address was populated with the identical
       -- value as the sort code/routing field (whitespace aside) - a data-entry
       -- error, not a real SWIFT code - so treat it as unusable rather than submit
       -- a fabricated-looking BIC to Workday.
     , CASE
         WHEN NULLIF(REPLACE(UPPER(TRIM(COALESCE(t.swift_address, ''))), ' ', ''), '') IN (
                REPLACE(UPPER(TRIM(COALESCE(t.eft_bank_code, ''))), ' ', '')
              , REPLACE(UPPER(TRIM(COALESCE(t.eft_transit_routing_no, ''))), ' ', '')
              )
         THEN NULL
         ELSE NULLIF(TRIM(UPPER(t.swift_address)), '')
       END                                                                                nrm_bank_swift_code
     , NULLIF(TRIM(t.building_society_roll_no), '')                                       nrm_bank_roll_number
     , NULLIF(TRIM(t.eft_bank_check_digit), '')                                           nrm_bank_check_digit
     , NULLIF(TRIM(t.eft_bank_branch_code), '')                                           nrm_bank_branch_id
     , NULLIF(TRIM(t.eft_bank_branch), '')                                                nrm_bank_branch_name
     , NULLIF(TRIM(t.eft_transfer_method), '')                                            nrm_bank_transfer_method
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
  FROM cte_supplier_addr_clean                        t
       -- Get the Royal Mail postcode for UK addresses
       LEFT OUTER JOIN
       ref_post_code_district                         r0
          ON (
               UPPER(TRIM(t.nrm_postal_code))         LIKE UPPER(TRIM(r0.postcode) || ' %')  -- Matches 'MK19 5AH' to 'MK19'
            OR UPPER(TRIM(t.nrm_postal_code))         =    UPPER(TRIM(r0.postcode))          -- Matches 'MK19'     to 'MK19'
             )
       -- First try: match on country name (higher population)
       LEFT OUTER JOIN
       ref_source_country_name_mapping                m_name
         ON  UPPER(TRIM(m_name.source_country_name))  = UPPER(TRIM(t.country))
       -- Second try: match on country code (fallback)
       LEFT OUTER JOIN
       ref_source_country_code_mapping                m_code
         ON  UPPER(TRIM(m_code.source_country_code))  = NULLIF(UPPER(TRIM(t.country_code)), '')
       -- Join to reference table using: name match > code match > default GB
       LEFT OUTER JOIN
       ref_country                                    r
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
       -- Supplier category normalization
       LEFT OUTER JOIN
       ref_supplier_category                          scm
         ON UPPER(TRIM(scm.source_supplier_category)) = COALESCE(NULLIF(UPPER(TRIM(t.vendor_class_id)), ''), 'NULL')
       -- Payment Terms normalization
       LEFT OUTER JOIN
       ref_source_payment_terms                       pt
         ON  UPPER(TRIM(pt.source_payment_terms))     = COALESCE(NULLIF(UPPER(TRIM(t.payment_terms_id)), ''), 'NULL')
       -- Sort Code normalisation - ref_bank_sort_code_prefix_mapping is UK clearing
       -- data (NatWest/Nationwide/Santander sort-code ranges), matched regardless of
       -- the supplier's own registered country - a non-UK supplier can genuinely hold
       -- a UK bank account, and a match is treated as evidence of that (see
       -- nrm_bank_name/nrm_bank_country_name below).
       LEFT OUTER JOIN
        ref_bank_sort_code_prefix_mapping             rbsc
          ON rbsc.sort_code_prefix                    = SUBSTR(NULLIF(TRIM(t.eft_bank_code), ''), 1, 2)
       -- A GB-format IBAN embeds the sort code at characters 9-14 (GBkk BBBB SSSSSS
       -- NNNNNNNN - 4-letter bank code, 6-digit sort code, 8-digit account number,
       -- 22 characters total) - falls back to this when eft_bank_code didn't resolve
       -- a bank name, recovering real accounts that would otherwise have no usable
       -- bank name and be dropped from workday_supplier_settlement_account entirely.
       LEFT OUTER JOIN
        ref_bank_sort_code_prefix_mapping             rbsc_iban
          ON rbsc_iban.sort_code_prefix                = SUBSTR(REPLACE(REPLACE(UPPER(COALESCE(t.iban, '')), ' ', ''), '-', ''), 9, 2)
         AND UPPER(SUBSTR(REPLACE(REPLACE(UPPER(COALESCE(t.iban, '')), ' ', ''), '-', ''), 1, 2)) = 'GB'
         AND LENGTH(REPLACE(REPLACE(UPPER(COALESCE(t.iban, '')), ' ', ''), '-', '')) = 22
       -- International bank-code-from-IBAN lookup (ref_iban_bank_code_mapping,
       -- researched per-country from official/primary sources - see that file's
       -- 'source' field). Matches the IBAN's country code and prefix-matches its
       -- BBAN (from character 5 onward) against each country's own bank-code length,
       -- taking the longest/most specific match (handles variable-length codes,
       -- e.g. Finland, without per-country hardcoded lengths).
       LEFT OUTER JOIN LATERAL (
         SELECT m.bank_name                                                             bank_name
              , rc_iban.country_name                                                    country_name
           FROM ref_iban_bank_code_mapping                                              m
           JOIN ref_country                                                             rc_iban
             ON UPPER(TRIM(rc_iban.country_code))                                       = UPPER(TRIM(m.country_code))
          WHERE m.country_code                                                          = SUBSTR(REPLACE(REPLACE(UPPER(COALESCE(t.iban, '')), ' ', ''), '-', ''), 1, 2)
            AND SUBSTR(REPLACE(REPLACE(UPPER(COALESCE(t.iban, '')), ' ', ''), '-', ''), 5) LIKE m.bank_code || '%'
          ORDER BY LENGTH(m.bank_code) DESC
          LIMIT 1
       )                                                    iban_intl
         ON TRUE
       -- National bank code embedded in a well-formed IBAN (ISO 13616 / SWIFT IBAN
       -- Registry bank-identifier field, ref_iban_structure). Used only as a fill-blank
       -- fallback for nrm_bank_code_routing_number above, for IBAN countries where GP
       -- populates no domestic bank/routing code field. The iban_length guard restricts
       -- extraction to a structurally complete IBAN, so a truncated or placeholder IBAN
       -- (e.g. a lone '0') yields nothing rather than a mangled code.
       LEFT OUTER JOIN LATERAL (
         SELECT NULLIF(TRIM(SUBSTR(
                    REPLACE(REPLACE(UPPER(t.iban), ' ', ''), '-', '')
                  , st.bank_code_offset
                  , st.bank_code_length)), '')                                          bank_code
           FROM ref_iban_structure                                                      st
          WHERE st.country_code
                  = SUBSTR(REPLACE(REPLACE(UPPER(COALESCE(t.iban, '')), ' ', ''), '-', ''), 1, 2)
            AND LENGTH(REPLACE(REPLACE(UPPER(COALESCE(t.iban, '')), ' ', ''), '-', ''))
                  = st.iban_length
       )                                                    iban_bankcode
         ON TRUE
       -- US routing number -> tenant bank name. ref_bank_routing_number_mapping is the
       -- tenant's own US routing->name directory (cloned EIB workbook, "Tenant Default Named
       -- Ranges" AT/AU), so the derived name is the exact spelling the EIB expects. Matched
       -- for US accounts on the 9-digit routing (transit routing preferred, else bank code;
       -- a dropped leading zero is restored by LPAD to 9).
       LEFT OUTER JOIN
        ref_bank_routing_number_mapping               rbrn
          ON r.country_code                           = 'US'
         AND rbrn.routing_number                      = LPAD(
               NULLIF(REGEXP_REPLACE(COALESCE(t.eft_transit_routing_no, t.eft_bank_code), '[^0-9]', '', 'g'), ''),
               9, '0')
       -- Address handling
       LEFT OUTER JOIN
       ref_post_code_workday_region                   r01
          ON UPPER(TRIM(r01.post_code_region))        = UPPER(TRIM(r0.region))
       LEFT OUTER JOIN
       ref_workday_country_state_region               r1
         ON UPPER(TRIM(r1.country))                   = UPPER(TRIM(r.country_name))
        AND UPPER(TRIM(r1.instance))                  = UPPER(TRIM(t.county))
       LEFT OUTER JOIN
       ref_workday_country_state_region               r2
         ON UPPER(TRIM(r2.country))                   = UPPER(TRIM(r.country_name))
        AND UPPER(TRIM(r2.instance))                  = UPPER(TRIM(t.city))
       LEFT OUTER JOIN
       ref_country_county_state_town_mapping          r3
         ON UPPER(TRIM(r3.country_code))              = UPPER(TRIM(r.country_code))
        AND UPPER(TRIM(r3.town_city_name))            = UPPER(TRIM(t.county))
       LEFT OUTER JOIN
       ref_country_county_state_town_mapping          r4
         ON UPPER(TRIM(r4.country_code))              = UPPER(TRIM(r.country_code))
        AND UPPER(TRIM(r4.town_city_name))            = UPPER(TRIM(t.city))
       LEFT OUTER JOIN
       ref_workday_country_state_region               rx
         ON UPPER(TRIM(rx.country))                   = UPPER(TRIM(r.country_name))
        AND UPPER(TRIM(rx.instance))                  = UPPER(TRIM(
                                                        CASE
                                                         WHEN r0.region             IS NOT NULL THEN UPPER(TRIM(COALESCE(r01.workday_region, r0.region)))
                                                         WHEN r1.instance           IS NOT NULL THEN UPPER(TRIM(r1.instance))
                                                         WHEN r2.instance           IS NOT NULL THEN UPPER(TRIM(r2.instance))
                                                         WHEN r3.county_state_name  IS NOT NULL THEN UPPER(TRIM(r3.county_state_name))
                                                         WHEN r4.county_state_name  IS NOT NULL THEN UPPER(TRIM(r4.county_state_name))
                                                         ELSE NULLIF(UPPER(TRIM(t.county)), '')
                                                        END))
       LEFT OUTER JOIN
       ref_workday_country_state_region               rxo
         ON UPPER(TRIM(rxo.country))                  = UPPER(TRIM(r.country_name))
        AND UPPER(TRIM(rxo.instance))                 = UPPER(TRIM(
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
       )
SELECT
       ROW_NUMBER() OVER
       (
         PARTITION BY
                   t.nrm_supplier_name
                 , t.nrm_supplier_number
           ORDER BY
                   CASE
                     WHEN t.nrm_bank_name IS NOT NULL
                     THEN 0
                     ELSE 1
                   END ASC NULLS LAST
                 , CASE
                     WHEN t.nrm_bank_account_number IS NOT NULL
                     THEN 0
                     ELSE 1
                   END ASC NULLS LAST
                 , t.nrm_created_date           ASC  NULLS LAST
                 , t.nrm_last_payment_date      DESC NULLS LAST
                 , t.nrm_last_purchase_date     DESC NULLS LAST
       )                                                                   data_rnk
     , t.*
  FROM cte_supplier_clean              t
;
