DROP TABLE IF EXISTS workday_customer_address_base
;
CREATE TABLE workday_customer_address_base
    AS
  WITH cte_address
    AS (
SELECT
       t.customer_id                                                          customer_id
     , t.nrm_customer_name                                                    customer_name
     , CAST(NULL AS VARCHAR)                                                  formatted_address
     , CAST(NULL AS VARCHAR)                                                  address_format_type
     , CAST(NULL AS VARCHAR)                                                  defaulted_business_site_address
     , CAST(NULL AS VARCHAR)                                                  delete_flag
     , CAST(NULL AS VARCHAR)                                                  do_not_replace_all
     , CAST(NULL AS VARCHAR)                                                  last_modified
     , CAST(NULL AS VARCHAR)                                                  descriptor
     , ROW_NUMBER() OVER (
           PARTITION BY t.customer_id
               ORDER BY u.address.nrm_address_code
                      , u.address.nrm_address_line_1
                      , u.address.nrm_address_line_2
                      , u.address.nrm_address_line_3
                      , u.address.nrm_address_line_4
                      , u.address.nrm_city
                      , u.address.nrm_region
                      , u.address.nrm_postal_code
                      , u.address.nrm_country_name
       )                                                                      address_id_rnk
     , u.address.nrm_address_code                                             address_id_code
     , u.address.nrm_country_name                                             country
     , u.address.nrm_country_code                                             country_code
     , CASE
           WHEN LOWER(TRIM(req.region_state)) = 'not accepted'
           THEN NULL
           ELSE u.address.nrm_region
       END                                                                    region
       -- Only exposed for countries Workday actually requires it for (per
       -- ref_workday_validate_customer_address.subregion_county), table-driven rather than a
       -- hardcoded country list. nrm_subregion_county (see src_fin_customer_raw.sql) sources
       -- from the raw county field for every country - whether it's actually surfaced here is
       -- decided by this gate alone.
     , CASE
           WHEN LOWER(TRIM(req.subregion_county)) = 'required'
           THEN u.address.nrm_subregion_county
           ELSE NULL
       END                                                                    subregion
     , CASE
           WHEN LOWER(TRIM(req.municipality_city)) = 'not accepted'
           THEN NULL
           ELSE u.address.nrm_city
       END                                                                    city
     , CAST(NULL AS VARCHAR)                                                  submunicipality
       -- Only exposed for countries Workday actually requires it for (currently ZA/FI/SK/
       -- MX/AX per ref_workday_validate_customer_address) - driven entirely by that table,
       -- not a hardcoded country list, so it automatically covers any future country
       -- Workday adds this requirement for. nrm_city_subdivision itself (see
       -- src_fin_customer_raw.sql) already carries the right source value per country -
       -- South Africa's real suburb data, Address Line 3 for everyone else. Still NULL for
       -- Finland/Slovakia until that one-time GP data migration lands.
     , CASE
           WHEN LOWER(TRIM(req.city_subdivision)) = 'required'
           THEN u.address.nrm_city_subdivision
           ELSE NULL
       END                                                                    city_subdivision
     , CASE
           WHEN u.address.nrm_country_code = 'HU'
            AND u.address.nrm_address_line_2 IS NULL
            AND u.address.nrm_address_line_1 IS NOT NULL
            AND LENGTH(u.address.nrm_address_line_1) > 35
            AND regexp_matches(u.address.nrm_address_line_1, '^.{1,35}\\s+.+$')
           THEN TRIM(REGEXP_REPLACE(u.address.nrm_address_line_1, '^(.{1,35})\\s+.+$', '\1'))
           ELSE u.address.nrm_address_line_1
       END                                                                    address_line_1
     , CASE
           WHEN u.address.nrm_country_code = 'HU'
            AND u.address.nrm_address_line_2 IS NULL
            AND u.address.nrm_address_line_1 IS NOT NULL
            AND LENGTH(u.address.nrm_address_line_1) > 35
            AND regexp_matches(u.address.nrm_address_line_1, '^.{1,35}\\s+.+$')
           THEN NULLIF(TRIM(REGEXP_REPLACE(u.address.nrm_address_line_1, '^.{1,35}\\s+(.+)$', '\1')), '')
           WHEN u.address.nrm_country_code = 'HU'
            AND u.address.nrm_address_line_2 IS NULL
           THEN NULLIF(TRIM(COALESCE(u.address.nrm_city, u.address.nrm_address_line_1)), '')
           WHEN LOWER(TRIM(req.address_line_2)) = 'not accepted'
           THEN NULL
           WHEN LOWER(TRIM(req.postal_code)) = 'not accepted'
            AND NULLIF(TRIM(u.address.nrm_postal_code), '') IS NOT NULL
            AND UPPER(TRIM(u.address.nrm_postal_code)) NOT LIKE '%PO BOX%'
            AND LENGTH(TRIM(u.address.nrm_postal_code)) <= 35
            AND (u.address.nrm_address_line_2 IS NULL OR TRIM(u.address.nrm_address_line_2) = '')
           THEN TRIM(u.address.nrm_postal_code)
           WHEN LOWER(TRIM(req.postal_code)) = 'not accepted'
            AND NULLIF(TRIM(u.address.nrm_postal_code), '') IS NOT NULL
            AND UPPER(TRIM(u.address.nrm_postal_code)) NOT LIKE '%PO BOX%'
            AND LENGTH(TRIM(u.address.nrm_postal_code)) <= 35
            AND u.address.nrm_address_line_2 IS NOT NULL
            AND TRIM(u.address.nrm_address_line_2) != ''
            AND LENGTH(TRIM(u.address.nrm_address_line_2)) + 1 + LENGTH(TRIM(u.address.nrm_postal_code)) <= 35
           THEN TRIM(u.address.nrm_address_line_2) || ' ' || TRIM(u.address.nrm_postal_code)
           ELSE u.address.nrm_address_line_2
       END                                                                    address_line_2
     , CASE
           WHEN LOWER(TRIM(req.address_line_3)) = 'not accepted'
           THEN NULL
           ELSE u.address.nrm_address_line_3
       END                                                                    address_line_3
     , CASE
           WHEN LOWER(TRIM(req.address_line_4)) = 'not accepted'
           THEN NULL
           ELSE u.address.nrm_address_line_4
       END                                                                    address_line_4
     , CASE
           WHEN LOWER(TRIM(req.postal_code)) = 'not accepted'
           THEN NULL
           ELSE u.address.nrm_postal_code
       END                                                                    postal_code
     , 'Yes'                                                                  is_public
     , 'Yes'                                                                  is_primary
     , CAST(NULL AS VARCHAR)                                                  address_type   
     , CAST(NULL AS VARCHAR)                                                  effective_date
     , CAST(NULL AS VARCHAR)                                                  use_for
     , CAST(NULL AS VARCHAR)                                                  use_for_tenanted
     , CAST(NULL AS VARCHAR)                                                  customer_comment
     , CAST(NULL AS VARCHAR)                                                  number_of_days
     , CAST(NULL AS VARCHAR)                                                  municipality_local
     , CAST(NULL AS VARCHAR)                                                  optional_address
  FROM src_fin_customer                         t
  CROSS JOIN UNNEST(nrm_array_customer_address) u(address)
  LEFT JOIN
       ref_workday_validate_customer_address    req
    ON LOWER(TRIM(req.country))                 = LOWER(TRIM(u.address.nrm_country_name))
 WHERE COALESCE( NULLIF(TRIM(u.address.nrm_address_line_1), '')
               , NULLIF(TRIM(u.address.nrm_address_line_2), '')
               , NULLIF(TRIM(u.address.nrm_address_line_3), '')
               , NULLIF(TRIM(u.address.nrm_address_line_4), '')
               )                                      IS NOT NULL
-- AND NULLIF(TRIM(u.address.nrm_address_line_1), '') NOT IN ('[Not Known]')
-- AND COALESCE(u.address.nrm_postal_code_valid, 'Y') = 'Y'
       )
SELECT t.customer_id
     , t.customer_name
     , t.formatted_address
     , t.address_format_type
     , t.defaulted_business_site_address
     , t.delete_flag
     , t.do_not_replace_all
     , t.last_modified
     , t.descriptor
     , TRIM(t.customer_id) || '_' || t.address_id_code || '_' || t.address_id_rnk   address_id
     , t.country
     , t.country_code
     , t.region
     , t.subregion
     , t.city
     , t.submunicipality
     , t.city_subdivision
     , t.address_line_1
     , t.address_line_2
     , t.address_line_3
     , t.address_line_4
     , t.postal_code
     , t.is_public
     , CASE t.address_id_rnk WHEN 1 THEN 'Yes' ELSE 'No' END                        is_primary
     , t.address_type   
     , t.effective_date
     , t.use_for
     , t.use_for_tenanted
     , t.customer_comment
     , t.number_of_days
     , t.municipality_local
     , t.optional_address
  FROM cte_address                     t
;
