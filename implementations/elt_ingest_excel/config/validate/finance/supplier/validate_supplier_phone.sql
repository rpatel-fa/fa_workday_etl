-- ============================================================================
-- VALIDATE: Supplier Phone
-- Source table: workday_supplier_phone
-- VBA rules: Match (cross-tab), Duplicatecolumn, Missing, Match (Yes/No),
--            OneSelectionPerInstance, MatchValueorFormat (phone digits by country)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Supplier ID must exist in Supplier Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_phone'                                       table_name
     , 'CROSS_TAB'                                                  rule_type
     , 'supplier_id'                                                col
     , p.supplier_id                                                supplier_id
     , p.supplier_id                                                detail
     , 'Supplier ID must be listed first on the Supplier Name tab.' message
  FROM workday_supplier_phone                                       p
  LEFT OUTER JOIN
       workday_supplier_name                                        n
         ON  n.supplier_id                                          = p.supplier_id
 WHERE n.supplier_id IS NULL
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: Phone ID must be unique
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_phone'                                           table_name
     , 'DUPLICATE'                                                       rule_type
     , 'phone_id'                                                        col
     , p.supplier_id                                                     supplier_id
     , p.phone_id                                                        detail
     , 'Phone ID must be unique and cannot be duplicated.'               message
  FROM workday_supplier_phone                                            p
 WHERE p.phone_id IN
       (
         SELECT d.phone_id
           FROM workday_supplier_phone                                   d
          GROUP BY d.phone_id
         HAVING COUNT(*) > 1
       )
;

-- ----------------------------------------------------------------------------
-- MISSING: Phone Device Type is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_phone'                                           table_name
     , 'MISSING'                                                         rule_type
     , 'phone_device_type'                                               col
     , p.supplier_id                                                     supplier_id
     , NULL                                                              detail
     , 'Phone device type is a required field.'                          message
  FROM workday_supplier_phone                                            p
 WHERE p.phone_device_type IS NULL
    OR TRIM(p.phone_device_type) = ''
;

-- ----------------------------------------------------------------------------
-- MISSING: Public flag is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_phone'                                           table_name
     , 'MISSING'                                                         rule_type
     , 'public_flag'                                                     col
     , p.supplier_id                                                     supplier_id
     , NULL                                                              detail
     , 'The public column is a required column.'                         message
  FROM workday_supplier_phone                                            p
 WHERE p.public_flag IS NULL
    OR TRIM(p.public_flag) = ''
;

-- ----------------------------------------------------------------------------
-- MATCH: Public flag must be Yes or No
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_phone'                                                 table_name
     , 'MATCH'                                                                 rule_type
     , 'public_flag'                                                           col
     , p.supplier_id                                                           supplier_id
     , p.public_flag                                                           detail
     , 'The Public column must be populated with either "Yes" or "No".'        message
  FROM workday_supplier_phone                                                  p
 WHERE p.public_flag IS NOT NULL
   AND TRIM(p.public_flag) != ''
   AND p.public_flag NOT IN ('Yes', 'No')
;

-- ----------------------------------------------------------------------------
-- MISSING: Primary flag is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_phone'                                           table_name
     , 'MISSING'                                                         rule_type
     , 'primary_flag'                                                    col
     , p.supplier_id                                                     supplier_id
     , NULL                                                              detail
     , 'The primary column is a required column.'                        message
  FROM workday_supplier_phone                                            p
 WHERE p.primary_flag IS NULL
    OR TRIM(p.primary_flag) = ''
;

-- ----------------------------------------------------------------------------
-- MATCH: Primary flag must be Yes or No
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_phone'                                                 table_name
     , 'MATCH'                                                                 rule_type
     , 'primary_flag'                                                          col
     , p.supplier_id                                                           supplier_id
     , p.primary_flag                                                          detail
     , 'The Primary column must be populated with either "Yes" or "No".'       message
  FROM workday_supplier_phone                                                  p
 WHERE p.primary_flag IS NOT NULL
   AND TRIM(p.primary_flag) != ''
   AND p.primary_flag NOT IN ('Yes', 'No')
;

-- ----------------------------------------------------------------------------
-- ONE_PRIMARY: Each Supplier ID + Use For group must have exactly one primary
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_phone'                                                  table_name
     , 'ONE_PRIMARY'                                                            rule_type
     , 'primary_flag'                                                           col
     , grp.supplier_id                                                          supplier_id
     , 'use_for=' || COALESCE(grp.use_for, 'ALL') || ' primary_count=' || grp.primary_count detail
     , 'One business phone must be marked as primary and one only.'             message
  FROM
       (
         SELECT
                p.supplier_id                                                   supplier_id
              , p.use_for                                                       use_for
              , COUNT(*) FILTER (WHERE p.primary_flag = 'Yes')                 primary_count
           FROM workday_supplier_phone                                          p
          GROUP BY p.supplier_id, p.use_for
         HAVING COUNT(*) FILTER (WHERE p.primary_flag = 'Yes') != 1
       )                                                                        grp
;

-- ----------------------------------------------------------------------------
-- FORMAT: UK phone number must be 9–10 digits (no special characters)
-- Note: workday_supplier_phone.phone_number is populated from udf_parse_phone(...).full_phone_number
--       which is the national significant number (no country code, no trunk prefix 0).
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_phone'                                                  table_name
     , 'FORMAT'                                                                 rule_type
     , 'phone_number'                                                           col
     , p.supplier_id                                                            supplier_id
     , COALESCE(p.area_code, '') || p.phone_number                             detail
     , 'Phone Number must be 9–10 digits for United Kingdom (no special characters).' message
  FROM workday_supplier_phone                                                   p
 WHERE p.phone_country                                                          = 'United Kingdom'
   AND p.phone_number IS NOT NULL
   AND NOT regexp_matches(
             regexp_replace(
               COALESCE(p.area_code, '') || COALESCE(p.phone_number, '')
             , '[^0-9]'
             , ''
             , 'g'
             )
           , '^\d{9,10}$'
       )
;

-- ----------------------------------------------------------------------------
-- FORMAT: Netherlands phone number must be 9 digits (no special characters)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_phone'                                                  table_name
     , 'FORMAT'                                                                 rule_type
     , 'phone_number'                                                           col
     , p.supplier_id                                                            supplier_id
     , COALESCE(p.area_code, '') || p.phone_number                             detail
     , 'Phone Number must be 9 digits for Netherlands (no special characters).' message
  FROM workday_supplier_phone                                                   p
 WHERE p.phone_country                                                          = 'Netherlands'
   AND p.phone_number IS NOT NULL
   AND NOT regexp_matches(
             regexp_replace(
               COALESCE(p.area_code, '') || COALESCE(p.phone_number, '')
             , '[^0-9]'
             , ''
             , 'g'
             )
           , '^\d{9}$'
       )
;

-- ----------------------------------------------------------------------------
-- MATCH: Phone Country must be a recognised Workday country name
-- ref: ref_workday_validate_supplier_phone_format (Phone Number sheet)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_phone'                                                         table_name
     , 'MATCH'                                                                        rule_type
     , 'phone_country'                                                                col
     , p.supplier_id                                                                  supplier_id
     , p.phone_country                                                                detail
     , 'Phone Country "' || p.phone_country || '" is not a recognised Workday country name.' message
  FROM workday_supplier_phone                                                          p
  LEFT JOIN
       ref_workday_validate_supplier_phone_format                                      r
          ON LOWER(TRIM(r.country))                                                   = LOWER(TRIM(p.phone_country))
 WHERE p.phone_country IS NOT NULL
   AND TRIM(p.phone_country) != ''
   AND r.country IS NULL
;

-- ----------------------------------------------------------------------------
-- FORMAT: International phone code must match the country
-- ref: ref_workday_validate_supplier_phone_format (Phone Number sheet)
-- workday_supplier_phone.international_phone_code is digits-only (e.g. "44")
-- ref table international_phone_code is also digits-only
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_phone'                                                              table_name
     , 'FORMAT'                                                                            rule_type
     , 'international_phone_code'                                                          col
     , p.supplier_id                                                                       supplier_id
     , p.international_phone_code || ' (expected ' || r.international_phone_code || ')'    detail
     , 'International phone code does not match the expected code for '
       || p.phone_country || ' (' || r.international_phone_code || ').'                    message
  FROM workday_supplier_phone                                                               p
  INNER JOIN
       ref_workday_validate_supplier_phone_format                                           r
          ON LOWER(TRIM(r.country))                                                        = LOWER(TRIM(p.phone_country))
 WHERE p.international_phone_code IS NOT NULL
   AND TRIM(p.international_phone_code) != ''
   AND TRIM(p.international_phone_code)                                                    != TRIM(r.international_phone_code)
;

-- ----------------------------------------------------------------------------
-- MISSING: International phone code is required when a phone number is present.
-- The MATCH/FORMAT checks above only fire when a code exists; this closes the
-- "phone country code missing" gap Workday flags but the internal validation did not.
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_phone'                                                              table_name
     , 'MISSING'                                                                            rule_type
     , 'international_phone_code'                                                          col
     , p.supplier_id                                                                       supplier_id
     , p.phone_number                                                                      detail
     , 'International phone code is required when a phone number is present.'               message
  FROM workday_supplier_phone                                                               p
 WHERE p.phone_number IS NOT NULL
   AND TRIM(p.phone_number) != ''
   AND (p.international_phone_code IS NULL OR TRIM(p.international_phone_code) = '')
;
