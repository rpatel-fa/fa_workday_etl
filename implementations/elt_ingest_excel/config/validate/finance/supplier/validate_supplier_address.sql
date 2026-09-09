-- ============================================================================
-- VALIDATE: Supplier Address
-- Source table: workday_supplier_address
-- VBA rules: Match (cross-tab), Duplicatecolumn, Missing, OneSelectionPerInstance,
--            Match (Yes/No flags), addressrequirementsvalidation
-- ============================================================================

DROP VIEW IF EXISTS v_validate_supplier_address
;
CREATE TEMP VIEW v_validate_supplier_address AS
SELECT 'workday_supplier_address_base' AS table_name
     , supplier_id
     , supplier_name
     , formatted_address
     , address_format_type
     , defaulted_business_site_address
     , delete_flag
     , do_not_replace_all
     , last_modified
     , descriptor
     , address_id
     , country
     , country_code
     , region
     , subregion
     , city
     , submunicipality
     , city_subdivision
     , address_line_1
     , address_line_2
     , address_line_3
     , address_line_4
     , postal_code
     , public_flag
     , is_primary
     , effective_date
     , address_type
     , use_for
     , use_for_tenanted
     , address_comments
     , number_of_days
     , municipality_local
  FROM workday_supplier_address_base
 UNION ALL
SELECT 'workday_supplier_address' AS table_name
     , supplier_id
     , supplier_name
     , formatted_address
     , address_format_type
     , defaulted_business_site_address
     , delete_flag
     , do_not_replace_all
     , last_modified
     , descriptor
     , address_id
     , country
     , country_code
     , region
     , subregion
     , city
     , submunicipality
     , submunicipality AS city_subdivision
     , address_line_1
     , address_line_2
     , address_line_3
     , address_line_4
     , postal_code
     , public_flag
     , is_primary
     , effective_date
     , address_type
     , use_for
     , use_for_tenanted
     , address_comments
     , number_of_days
     , municipality_local
  FROM workday_supplier_address
;

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Supplier ID must exist in Supplier Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       a.table_name                                                   table_name
     , 'CROSS_TAB'                                                   rule_type
     , 'supplier_id'                                                 col
     , a.supplier_id                                                 supplier_id
     , a.supplier_id                                                 detail
     , 'Supplier ID must be listed first on the Supplier Name tab.'  message
  FROM v_validate_supplier_address                                   a
  LEFT OUTER JOIN
       workday_supplier_name                                         n
         ON  n.supplier_id                                           = a.supplier_id
 WHERE n.supplier_id IS NULL
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: Address ID must be unique
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       a.table_name                                                       table_name
     , 'DUPLICATE'                                                        rule_type
     , 'address_id'                                                       col
     , a.supplier_id                                                      supplier_id
     , a.address_id                                                       detail
     , 'Address ID must be unique and cannot be duplicated.'              message
  FROM v_validate_supplier_address                                        a
 WHERE a.address_id IS NOT NULL
   AND a.address_id IN
       (
         SELECT d.address_id
           FROM v_validate_supplier_address                               d
          WHERE d.address_id IS NOT NULL
            AND d.table_name = a.table_name
          GROUP BY d.address_id
         HAVING COUNT(*) > 1
       )
;

-- ----------------------------------------------------------------------------
-- MISSING: Country is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       a.table_name                                                              table_name
     , 'MISSING'                                                                 rule_type
     , 'country'                                                                 col
     , a.supplier_id                                                             supplier_id
     , NULL                                                                      detail
     , 'Country is a required column and must be populated for the used row.'    message
  FROM v_validate_supplier_address                                               a
 WHERE a.country IS NULL
    OR TRIM(a.country) = ''
;

-- ----------------------------------------------------------------------------
-- MISSING: Address Line 1 is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       a.table_name                                                              table_name
     , 'MISSING'                                                                 rule_type
     , 'address_line_1'                                                          col
     , a.supplier_id                                                             supplier_id
     , NULL                                                                      detail
     , 'Address Line 1 is a required column and must be populated for the used row.' message
  FROM v_validate_supplier_address                                               a
 WHERE a.address_line_1 IS NULL
    OR TRIM(a.address_line_1) = ''
;

-- ----------------------------------------------------------------------------
-- MISSING: Public flag is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       a.table_name                                                   table_name
     , 'MISSING'                                                       rule_type
     , 'public_flag'                                                   col
     , a.supplier_id                                                   supplier_id
     , NULL                                                            detail
     , 'The public column is a required column.'                       message
  FROM v_validate_supplier_address                                     a
 WHERE a.public_flag IS NULL
    OR TRIM(a.public_flag) = ''
;

-- ----------------------------------------------------------------------------
-- MATCH: Public flag must be Yes or No
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       a.table_name                                                               table_name
     , 'MATCH'                                                                   rule_type
     , 'public_flag'                                                              col
     , a.supplier_id                                                              supplier_id
     , a.public_flag                                                              detail
     , 'The Public column must be populated with either "Yes" or "No".'           message
  FROM v_validate_supplier_address                                                a
 WHERE a.public_flag IS NOT NULL
   AND TRIM(a.public_flag) != ''
   AND a.public_flag NOT IN ('Yes', 'No')
;

-- ----------------------------------------------------------------------------
-- MISSING: Primary flag is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       a.table_name                                                   table_name
     , 'MISSING'                                                       rule_type
     , 'is_primary'                                                    col
     , a.supplier_id                                                   supplier_id
     , NULL                                                            detail
     , 'The primary column is a required column.'                      message
  FROM v_validate_supplier_address                                     a
 WHERE a.is_primary IS NULL
    OR TRIM(a.is_primary) = ''
;

-- ----------------------------------------------------------------------------
-- MATCH: Primary flag must be Yes or No
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       a.table_name                                                               table_name
     , 'MATCH'                                                                   rule_type
     , 'is_primary'                                                               col
     , a.supplier_id                                                              supplier_id
     , a.is_primary                                                               detail
     , 'The Primary column must be populated with either "Yes" or "No".'          message
  FROM v_validate_supplier_address                                                a
 WHERE a.is_primary IS NOT NULL
   AND TRIM(a.is_primary) != ''
   AND a.is_primary NOT IN ('Yes', 'No')
;

-- ----------------------------------------------------------------------------
-- ONE_PRIMARY: Each Supplier ID + Use For group must have exactly one primary
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       grp.table_name                                                             table_name
     , 'ONE_PRIMARY'                                                              rule_type
     , 'is_primary'                                                               col
     , grp.supplier_id                                                            supplier_id
     , 'use_for=' || COALESCE(grp.use_for, 'ALL') || ' primary_count=' || grp.primary_count detail
     , 'One business address must be marked as primary and one only.'             message
  FROM
       (
         SELECT
                a.table_name                                                      table_name
              , a.supplier_id                                                     supplier_id
              , a.use_for                                                         use_for
              , COUNT(*) FILTER (WHERE a.is_primary = 'Yes')                     primary_count
           FROM v_validate_supplier_address                                       a
          GROUP BY a.table_name, a.supplier_id, a.use_for
         HAVING COUNT(*) FILTER (WHERE a.is_primary = 'Yes') != 1
       )                                                                          grp
;

INSERT INTO validation_supplier_issues
SELECT
       a.table_name                                                               table_name
     , 'ADDRESS_REQ'                                                              rule_type
     , 'country'                                                                  col
     , a.supplier_id                                                              supplier_id
     , COALESCE(a.country, a.country_code)                                        detail
     , 'No address requirements defined for this country.'                        message
  FROM v_validate_supplier_address                                                a
  LEFT JOIN
       (
         SELECT DISTINCT
                r.country                                                         country
           FROM ref_workday_validate_supplier_address                             r
       )                                                                          r
    ON LOWER(TRIM(r.country))                                                     = LOWER(TRIM(a.country))
 WHERE NULLIF(TRIM(a.country), '') IS NOT NULL
   AND r.country IS NULL
;

INSERT INTO validation_supplier_issues
SELECT
       addr.table_name                                                            table_name
     , 'ADDRESS_REQ'                                                              rule_type
     , req.workday_col                                                            col
     , addr.supplier_id                                                           supplier_id
     , COALESCE(addr.country, addr.country_code)                                  detail
     , CASE
           WHEN req.req_level                                                     = 'required'
                THEN REPLACE(req.workday_col, '_', ' ') || ' is required for ' || COALESCE(addr.country, addr.country_code) || ' addresses.'
           WHEN req.req_level                                                     = 'recommended'
                THEN REPLACE(req.workday_col, '_', ' ') || ' is recommended for ' || COALESCE(addr.country, addr.country_code) || ' addresses.'
       END                                                                        message
  FROM
       (
         SELECT
                table_name                                                        table_name
              , supplier_id                                                       supplier_id
              , country                                                           country
              , country_code                                                      country_code
              , field_name                                                        col
              , NULLIF(TRIM(CAST(field_value AS VARCHAR)), '')                     val
           FROM v_validate_supplier_address
                UNPIVOT INCLUDE NULLS(field_value FOR field_name IN (* EXCLUDE (table_name, supplier_id, supplier_name, country, country_code)))
       )                                                                          addr
  INNER JOIN
       (
         SELECT
                country                                                           country
              , CASE
                    WHEN req_col IN ('postcode', 'post_code', 'zip', 'zip_code')
                         THEN 'postal_code'
                    WHEN req_col IN ('state', 'province', 'state_province', 'county', 'county_state', 'region_county', 'region_state')
                         THEN 'region'
                    WHEN req_col IN ('town')
                         THEN 'city'
                    ELSE req_col
                END                                                               workday_col
              , NULLIF(TRIM(CAST(req_value AS VARCHAR)), '')                       req_val
              , CASE
                    WHEN LOWER(NULLIF(TRIM(CAST(req_value AS VARCHAR)), '')) LIKE 'required%'
                         OR LOWER(NULLIF(TRIM(CAST(req_value AS VARCHAR)), '')) IN ('required', 'yes', 'y', 'mandatory', 'must')
                         THEN 'required'
                    WHEN LOWER(NULLIF(TRIM(CAST(req_value AS VARCHAR)), '')) LIKE 'recommend%'
                         THEN 'recommended'
                END                                                               req_level
           FROM ref_workday_validate_supplier_address
                UNPIVOT(req_value FOR req_col IN (* EXCLUDE (country, country_code)))
          WHERE NULLIF(TRIM(CAST(req_value AS VARCHAR)), '') IS NOT NULL
       )                                                                          req
    ON LOWER(TRIM(req.country))                                                   = LOWER(TRIM(addr.country))
   AND req.workday_col                                                            = addr.col
 WHERE req.req_level IS NOT NULL
   AND addr.val IS NULL
;

INSERT INTO validation_supplier_issues
SELECT
       a.table_name                                                                    table_name
     , 'NOT_ACCEPTED'                                                                  rule_type
     , v.col                                                                           col
     , a.supplier_id                                                                   supplier_id
     , a.address_id || ' | ' || COALESCE(a.country, a.country_code)                    detail
     , REPLACE(v.col, '_', ' ') || ' is not accepted for ' || COALESCE(a.country, a.country_code)
       || ' addresses. Value: ' || v.val                                              message
  FROM v_validate_supplier_address                                                     a
  INNER JOIN
       ref_workday_validate_supplier_address                                           r
          ON LOWER(TRIM(r.country))                                                   = LOWER(TRIM(a.country))
  INNER JOIN LATERAL (
       SELECT *
         FROM (VALUES
               ('region',        a.region,        r.region_state),
               ('city',          a.city,          r.municipality_city),
               ('address_line_1', a.address_line_1, r.address_line_1),
               ('address_line_2', a.address_line_2, r.address_line_2),
               ('address_line_3', a.address_line_3, r.address_line_3),
               ('address_line_4', a.address_line_4, r.address_line_4),
               ('postal_code',   a.postal_code,   r.postal_code)
         ) AS x(col, val, req_val)
        WHERE NULLIF(TRIM(CAST(x.val AS VARCHAR)), '') IS NOT NULL
          AND LOWER(NULLIF(TRIM(CAST(x.req_val AS VARCHAR)), '')) = 'not accepted'
  )                                                                                   v
    ON TRUE
;

-- ----------------------------------------------------------------------------
-- MATCH: Region must be a valid Workday region for the address country
-- ref: ref_workday_validate_supplier_region (Country States-Regions sheet)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       a.table_name                                                                    table_name
     , 'MATCH'                                                                        rule_type
     , 'region'                                                                       col
     , a.supplier_id                                                                  supplier_id
     , a.region || ' (' || COALESCE(a.country, a.country_code) || ')'                 detail
     , 'Region "' || a.region || '" is not a valid Workday region for '
       || COALESCE(a.country, a.country_code) || '.'                                  message
  FROM v_validate_supplier_address                                                     a
  LEFT JOIN
       (
         SELECT LOWER(TRIM(country))       AS country
              , LOWER(TRIM(instance))      AS region_value
           FROM ref_workday_validate_supplier_region
          WHERE NULLIF(TRIM(instance), '') IS NOT NULL
         UNION
         SELECT LOWER(TRIM(country))
              , LOWER(TRIM(reference_ids))
           FROM ref_workday_validate_supplier_region
          WHERE NULLIF(TRIM(reference_ids), '') IS NOT NULL
       )                                                                               r
    ON r.country                                                                       = LOWER(TRIM(a.country))
   AND r.region_value                                                                  = LOWER(TRIM(a.region))
 WHERE a.region IS NOT NULL
   AND TRIM(a.region) != ''
   AND r.region_value IS NULL
;

-- ----------------------------------------------------------------------------
-- FORMAT: Postal code required for countries with a defined validation rule
-- ref: ref_workday_validate_supplier_postal_code (Postal Code Formatting sheet)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       a.table_name                                                                    table_name
     , 'FORMAT'                                                                       rule_type
     , 'postal_code'                                                                  col
     , a.supplier_id                                                                  supplier_id
     , a.address_id || ' | ' || COALESCE(a.country, a.country_code)                   detail
     , 'Postal code is required for ' || COALESCE(a.country, a.country_code)
       || '. Format: ' || r.postal_code_validation                                    message
  FROM v_validate_supplier_address                                                     a
  INNER JOIN
       ref_workday_validate_supplier_postal_code                                       r
          ON LOWER(TRIM(r.country))                                                   = LOWER(TRIM(a.country))
 WHERE LOWER(TRIM(r.postal_code_validation))                                           != 'no validation'
   AND NULLIF(TRIM(r.postal_code_validation), '')                                      IS NOT NULL
   AND (a.postal_code IS NULL OR TRIM(a.postal_code) = '')
;

INSERT INTO validation_supplier_issues
SELECT
       a.table_name                                                                    table_name
     , 'FORMAT'                                                                       rule_type
     , 'postal_code'                                                                  col
     , a.supplier_id                                                                  supplier_id
     , a.address_id || ' | ' || a.postal_code || ' (' || COALESCE(a.country, a.country_code) || ')' detail
     , 'Postal code "' || a.postal_code || '" does not match expected format for '
       || COALESCE(a.country, a.country_code)
       || '. Expected: ' || r.postal_code_validation
       || '. Normalized: ' || COALESCE(x.expected_postal_code, '<none>')             message
  FROM v_validate_supplier_address                                                     a
  INNER JOIN
       ref_workday_validate_supplier_postal_code                                       r
          ON LOWER(TRIM(r.country))                                                   = LOWER(TRIM(a.country))
  INNER JOIN LATERAL (
       SELECT udf_normalize_postal_code(a.postal_code, r.postal_code_validation)      expected_postal_code
  )                                                                                   x
    ON TRUE
 WHERE LOWER(TRIM(r.postal_code_validation))                                           != 'no validation'
   AND NULLIF(TRIM(r.postal_code_validation), '')                                      IS NOT NULL
   AND NULLIF(TRIM(a.postal_code), '')                                                 IS NOT NULL
   AND (
         x.expected_postal_code IS NULL
      OR UPPER(TRIM(a.postal_code))                                                    != UPPER(TRIM(x.expected_postal_code))
   )
;
