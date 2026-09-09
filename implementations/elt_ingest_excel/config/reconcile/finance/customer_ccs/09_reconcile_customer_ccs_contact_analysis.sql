DROP TABLE IF EXISTS reconcile_customer_ccs_contact_detail
;
CREATE TABLE reconcile_customer_ccs_contact_detail
    AS
  WITH cte_bridge
    AS (
SELECT DISTINCT
       b.customer_number
     , b.source_row_count
     , b.source_business_units
     , b.source_customer_names
     , b.source_countries
     , b.expected_customer_id
     , b.expected_reference_id
     , b.workday_customer_id
     , b.workday_reference_id
  FROM reconcile_customer_ccs_bridge                                               b
 WHERE b.in_source_gp                                                              = 'Y'
   AND b.in_merged_customer                                                        = 'Y'
   AND b.in_workday                                                                = 'Y'
       )
     , cte_primary_address
    AS (
SELECT t.customer_id                                                               expected_customer_id
     , NULLIF(TRIM(t.address_line_1), '')                                          expected_address_line_1
     , NULLIF(TRIM(t.address_line_2), '')                                          expected_address_line_2
     , NULLIF(TRIM(t.address_line_3), '')                                          expected_address_line_3
     , NULLIF(TRIM(t.address_line_4), '')                                          expected_address_line_4
     , NULLIF(TRIM(t.city), '')                                                    expected_city
     , NULLIF(TRIM(t.region), '')                                                  expected_region
     , NULLIF(TRIM(t.postal_code), '')                                             expected_postal_code
     , NULLIF(TRIM(t.country), '')                                                 expected_country
     , CONCAT_WS(
           CHR(10)
         , NULLIF(TRIM(t.address_line_1), '')
         , NULLIF(TRIM(t.address_line_2), '')
         , NULLIF(TRIM(t.address_line_3), '')
         , NULLIF(TRIM(t.address_line_4), '')
         , NULLIF(TRIM(t.city), '')
         , NULLIF(TRIM(t.region), '')
         , NULLIF(TRIM(t.postal_code), '')
         , NULLIF(TRIM(t.country), '')
       )                                                                           expected_value
     , REGEXP_REPLACE(
           UPPER(
               COALESCE(
                   CONCAT_WS(
                       CHR(10)
                     , NULLIF(TRIM(t.address_line_1), '')
                     , NULLIF(TRIM(t.address_line_2), '')
                     , NULLIF(TRIM(t.address_line_3), '')
                     , NULLIF(TRIM(t.address_line_4), '')
                     , NULLIF(TRIM(t.city), '')
                     , NULLIF(TRIM(t.region), '')
                     , NULLIF(TRIM(t.postal_code), '')
                     , NULLIF(TRIM(t.country), '')
                   )
                 , ''
               )
           )
         , '[^A-Z0-9]'
         , ''
         , 'g'
       )                                                                           expected_value_norm
     , ARRAY_TO_STRING(
           ARRAY_SORT(
               ARRAY_DISTINCT(
                   STRING_SPLIT(
                   REGEXP_REPLACE(
                       REPLACE(
                           REPLACE(
                               REPLACE(
                                   UPPER(
                                       COALESCE(
                                           CONCAT_WS(
                                               CHR(10)
                                             , NULLIF(TRIM(t.address_line_1), '')
                                             , NULLIF(TRIM(t.address_line_2), '')
                                             , NULLIF(TRIM(t.address_line_3), '')
                                             , NULLIF(TRIM(t.address_line_4), '')
                                             , NULLIF(TRIM(t.city), '')
                                             , NULLIF(TRIM(t.region), '')
                                             , NULLIF(TRIM(t.postal_code), '')
                                             , NULLIF(TRIM(t.country), '')
                                           )
                                         , ''
                                       )
                                   )
                                 , 'CO.'
                                 , ' '
                               )
                             , 'N°'
                             , ' '
                           )
                         , 'º'
                         , ' '
                       )
                    , '[^A-Z0-9]+'
                    , ' '
                    , 'g'
                  )
                , ' '
              )
               )
           )
         , '|'
       )                                                                           expected_value_token_signature
     , ARRAY_TO_STRING(
           ARRAY_SORT(
               ARRAY_DISTINCT(
                   STRING_SPLIT(
                   REGEXP_REPLACE(
                       REGEXP_REPLACE(
                           REPLACE(
                               REPLACE(
                                   REPLACE(
                                       UPPER(
                                           COALESCE(
                                               CONCAT_WS(
                                                   CHR(10)
                                                 , NULLIF(TRIM(t.address_line_1), '')
                                                 , NULLIF(TRIM(t.address_line_2), '')
                                                 , NULLIF(TRIM(t.address_line_3), '')
                                                 , NULLIF(TRIM(t.address_line_4), '')
                                                 , NULLIF(TRIM(t.city), '')
                                                 , NULLIF(TRIM(t.region), '')
                                                 , NULLIF(TRIM(t.postal_code), '')
                                                 , NULLIF(TRIM(t.country), '')
                                               )
                                             , ''
                                           )
                                       )
                                     , 'CO.'
                                     , ' '
                                   )
                                 , 'N°'
                                 , ' '
                               )
                             , 'º'
                             , ' '
                           )
                         , '[^A-Z0-9]+'
                         , ' '
                         , 'g'
                       )
                    , '(^| )0+([0-9]{3,})( |$)'
                    , '\\1\\2\\3'
                    , 'g'
                  )
                , ' '
              )
               )
           )
         , '|'
       )                                                                           expected_value_token_signature_zero_norm
  FROM workday_customer_address                                                    t
 WHERE COALESCE(TRIM(t.is_primary), '')                                            = 'Yes'
       )
     , cte_region_abbreviation_map
    AS (
SELECT 'UNITED STATES OF AMERICA'                                                  country_name
     , 'CALIFORNIA'                                                                full_region_name
     , 'CA'                                                                        region_abbreviation

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'FLORIDA', 'FL'

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'NEW YORK', 'NY'

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'TEXAS', 'TX'

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'CONNECTICUT', 'CT'

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'ILLINOIS', 'IL'

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'NORTH CAROLINA', 'NC'

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'VIRGINIA', 'VA'

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'WASHINGTON', 'WA'

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'DELAWARE', 'DE'

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'GEORGIA', 'GA'

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'MARYLAND', 'MD'

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'MINNESOTA', 'MN'

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'OREGON', 'OR'

 UNION ALL
SELECT 'UNITED STATES OF AMERICA', 'PENNSYLVANIA', 'PA'

 UNION ALL
SELECT 'AUSTRALIA', 'NEW SOUTH WALES', 'NSW'

 UNION ALL
SELECT 'CANADA', 'BRITISH COLUMBIA', 'BC'

 UNION ALL
SELECT 'CANADA', 'ONTARIO', 'ON'

 UNION ALL
SELECT 'CANADA', 'ALBERTA', 'AB'

 UNION ALL
SELECT 'BRAZIL', 'SAO PAULO', 'SP'

 UNION ALL
SELECT 'BRAZIL', 'SÃO PAULO', 'SP'

 UNION ALL
SELECT 'BRAZIL', 'RIO DE JANEIRO', 'RJ'

 UNION ALL
SELECT 'BRAZIL', 'RIO GRANDE DO SUL', 'RS'

 UNION ALL
SELECT 'BRAZIL', 'SANTA CATARINA', 'SC'

 UNION ALL
SELECT 'ITALY', 'MILANO', 'MI'

 UNION ALL
SELECT 'ITALY', 'PERUGIA', 'PG'

 UNION ALL
SELECT 'ITALY', 'PARMA', 'PR'

 UNION ALL
SELECT 'ITALY', 'CROTONE', 'KR'

 UNION ALL
SELECT 'SWITZERLAND', 'SOLOTHURN', 'SO'

 UNION ALL
SELECT 'SWITZERLAND', 'ZURICH', 'ZH'

 UNION ALL
SELECT 'SWITZERLAND', 'LUCERNE', 'LU'
       )
     , cte_primary_phone
    AS (
SELECT t.customer_id                                                               expected_customer_id
     , t.formatted_phone_number                                                    expected_value
     , NULLIF(
           REGEXP_REPLACE(
               COALESCE(
                   t.international_phone_code
                 , c.nrm_phone_code
                 , ''
               )
             , '[^0-9]'
             , ''
             , 'g'
           )
         , ''
       )
                                                                                    expected_country_digits
     , NULLIF(REGEXP_REPLACE(COALESCE(t.phone_number, ''), '[^0-9]', '', 'g'), '')
                                                                                    expected_local_digits
     , CASE
         WHEN NULLIF(REGEXP_REPLACE(COALESCE(t.phone_number, ''), '[^0-9]', '', 'g'), '') IS NULL
         THEN NULL
         WHEN NULLIF(
                  REGEXP_REPLACE(
                      COALESCE(
                          t.international_phone_code
                        , c.nrm_phone_code
                        , ''
                      )
                    , '[^0-9]'
                    , ''
                    , 'g'
                  )
                , ''
              ) IS NULL
         THEN NULLIF(REGEXP_REPLACE(COALESCE(t.phone_number, ''), '[^0-9]', '', 'g'), '')
         ELSE NULLIF(
                  REGEXP_REPLACE(
                      COALESCE(
                          t.international_phone_code
                        , c.nrm_phone_code
                        , ''
                      )
                    , '[^0-9]'
                    , ''
                    , 'g'
                  )
                , ''
              )
            || NULLIF(REGEXP_REPLACE(COALESCE(t.phone_number, ''), '[^0-9]', '', 'g'), '')
       END                                                                         expected_extract_digits
  FROM workday_customer_phone                                                      t
       LEFT OUTER JOIN
       src_fin_customer                                                            c
          ON c.customer_id                                                         = t.customer_id
 WHERE COALESCE(TRIM(t.primary_flag), '')                                          = 'Yes'
       )
SELECT t.customer_number
     , t.source_row_count
     , t.source_business_units
     , t.source_customer_names
     , t.source_countries
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , 'default_address'                                                           attribute_name
     , a.expected_value                                                            expected_value
     , w.default_address                                                           actual_value
     , CASE
         WHEN NULLIF(TRIM(COALESCE(a.expected_value, '')), '') IS NULL
          AND NULLIF(TRIM(COALESCE(w.default_address, '')), '') IS NULL
         THEN 'BOTH_BLANK'
         WHEN COALESCE(TRIM(a.expected_value), '') = COALESCE(TRIM(w.default_address), '')
         THEN 'EXACT_MATCH'
         WHEN COALESCE(a.expected_value, '') <> ''
          AND COALESCE(w.default_address, '') <> ''
          AND a.expected_value_norm = REGEXP_REPLACE(
                  UPPER(COALESCE(w.default_address, ''))
                , '[^A-Z0-9]'
                , ''
                , 'g'
              )
         THEN 'FORMAT_VARIATION'
         WHEN COALESCE(a.expected_value, '') <> ''
          AND COALESCE(w.default_address, '') <> ''
          AND a.expected_value_token_signature = ARRAY_TO_STRING(
                  ARRAY_SORT(
                      ARRAY_DISTINCT(
                          STRING_SPLIT(
                          REGEXP_REPLACE(
                              REPLACE(
                                  REPLACE(
                                      REPLACE(
                                          UPPER(COALESCE(w.default_address, ''))
                                        , 'CO.'
                                        , ' '
                                      )
                                    , 'N°'
                                    , ' '
                                  )
                                , 'º'
                                , ' '
                              )
                            , '[^A-Z0-9]+'
                            , ' '
                            , 'g'
                          )
                        , ' '
                      )
                      )
                  )
                , '|'
              )
         THEN 'FORMAT_VARIATION'
         WHEN COALESCE(a.expected_value, '') <> ''
          AND COALESCE(w.default_address, '') <> ''
          AND a.expected_value_token_signature_zero_norm = ARRAY_TO_STRING(
                  ARRAY_SORT(
                      ARRAY_DISTINCT(
                          STRING_SPLIT(
                          REGEXP_REPLACE(
                              REGEXP_REPLACE(
                                  REPLACE(
                                      REPLACE(
                                          REPLACE(
                                              UPPER(COALESCE(w.default_address, ''))
                                            , 'CO.'
                                            , ' '
                                          )
                                        , 'N°'
                                        , ' '
                                      )
                                    , 'º'
                                    , ' '
                                  )
                                , '[^A-Z0-9]+'
                                , ' '
                                , 'g'
                              )
                            , '(^| )0+([0-9]{3,})( |$)'
                            , '\\1\\2\\3'
                            , 'g'
                          )
                        , ' '
                      )
                      )
                  )
                , '|'
              )
         THEN 'FORMAT_VARIATION'
         WHEN COALESCE(a.expected_value, '') <> ''
          AND COALESCE(w.default_address, '') <> ''
          AND EXISTS (
                SELECT 1
                  FROM cte_region_abbreviation_map                                 m
                 WHERE m.country_name                                               = UPPER(COALESCE(a.expected_country, ''))
                   AND m.full_region_name                                           = UPPER(COALESCE(a.expected_region, ''))
                   AND ARRAY_TO_STRING(
                           ARRAY_SORT(
                               ARRAY_DISTINCT(
                                   STRING_SPLIT(
                                   REGEXP_REPLACE(
                                       REPLACE(
                                           REPLACE(
                                               REPLACE(
                                                   UPPER(
                                                       COALESCE(
                                                           CONCAT_WS(
                                                               CHR(10)
                                                             , a.expected_address_line_1
                                                             , a.expected_address_line_2
                                                             , a.expected_address_line_3
                                                             , a.expected_address_line_4
                                                             , a.expected_city
                                                             , m.region_abbreviation
                                                             , a.expected_postal_code
                                                             , a.expected_country
                                                           )
                                                         , ''
                                                       )
                                                   )
                                                 , 'CO.'
                                                 , ' '
                                               )
                                             , 'N°'
                                             , ' '
                                           )
                                         , 'º'
                                         , ' '
                                       )
                                     , '[^A-Z0-9]+'
                                     , ' '
                                     , 'g'
                                   )
                                 , ' '
                               )
                               )
                           )
                         , '|'
                       ) = ARRAY_TO_STRING(
                               ARRAY_SORT(
                                   ARRAY_DISTINCT(
                                       STRING_SPLIT(
                                       REGEXP_REPLACE(
                                           REPLACE(
                                               REPLACE(
                                                   REPLACE(
                                                       UPPER(COALESCE(w.default_address, ''))
                                                     , 'CO.'
                                                     , ' '
                                                   )
                                                 , 'N°'
                                                 , ' '
                                               )
                                             , 'º'
                                             , ' '
                                           )
                                         , '[^A-Z0-9]+'
                                         , ' '
                                         , 'g'
                                       )
                                     , ' '
                                   )
                                   )
                               )
                             , '|'
                           )
              )
         THEN 'FORMAT_VARIATION'
        WHEN COALESCE(a.expected_value, '') <> ''
         AND COALESCE(w.default_address, '') <> ''
         AND COALESCE(TRIM(a.expected_postal_code), '')                            <> ''
         AND (
                 COALESCE(TRIM(a.expected_city), '') <> ''
              OR COALESCE(TRIM(a.expected_address_line_2), '') <> ''
             )
         AND REGEXP_REPLACE(
                 UPPER(COALESCE(w.default_address, ''))
               , '[^A-Z0-9]'
               , ''
               , 'g'
             ) LIKE '%' || REGEXP_REPLACE(
                               UPPER(COALESCE(a.expected_address_line_1, ''))
                             , '[^A-Z0-9]'
                             , ''
                             , 'g'
                           ) || '%'
         AND (
                 COALESCE(TRIM(a.expected_address_line_2), '') = ''
              OR REGEXP_REPLACE(
                     UPPER(COALESCE(w.default_address, ''))
                   , '[^A-Z0-9]'
                   , ''
                   , 'g'
                 ) LIKE '%' || REGEXP_REPLACE(
                                   UPPER(COALESCE(a.expected_address_line_2, ''))
                                 , '[^A-Z0-9]'
                                 , ''
                                 , 'g'
                               ) || '%'
             )
         AND (
                 COALESCE(TRIM(a.expected_address_line_3), '') = ''
              OR REGEXP_REPLACE(
                     UPPER(COALESCE(w.default_address, ''))
                   , '[^A-Z0-9]'
                   , ''
                   , 'g'
                 ) LIKE '%' || REGEXP_REPLACE(
                                   UPPER(COALESCE(a.expected_address_line_3, ''))
                                 , '[^A-Z0-9]'
                                 , ''
                                 , 'g'
                               ) || '%'
             )
         AND (
                 COALESCE(TRIM(a.expected_address_line_4), '') = ''
              OR REGEXP_REPLACE(
                     UPPER(COALESCE(w.default_address, ''))
                   , '[^A-Z0-9]'
                   , ''
                   , 'g'
                 ) LIKE '%' || REGEXP_REPLACE(
                                   UPPER(COALESCE(a.expected_address_line_4, ''))
                                 , '[^A-Z0-9]'
                                 , ''
                                 , 'g'
                               ) || '%'
             )
         AND (
                 COALESCE(TRIM(a.expected_city), '') <> ''
              AND REGEXP_REPLACE(
                     UPPER(COALESCE(w.default_address, ''))
                   , '[^A-Z0-9]'
                   , ''
                   , 'g'
                 ) LIKE '%' || REGEXP_REPLACE(
                                   UPPER(COALESCE(a.expected_city, ''))
                                 , '[^A-Z0-9]'
                                 , ''
                                 , 'g'
                               ) || '%'
              OR COALESCE(TRIM(a.expected_city), '') = ''
             )
         AND (
                 COALESCE(TRIM(a.expected_address_line_2), '') <> ''
              OR COALESCE(TRIM(a.expected_city), '') <> ''
             )
         AND ARRAY_TO_STRING(
                 ARRAY_SORT(
                     ARRAY_DISTINCT(
                         STRING_SPLIT(
                             REGEXP_REPLACE(
                                 REPLACE(
                                     REPLACE(
                                         REPLACE(
                                             UPPER(COALESCE(w.default_address, ''))
                                           , 'CO.'
                                           , ' '
                                         )
                                       , 'N°'
                                       , ' '
                                     )
                                   , 'º'
                                   , ' '
                                 )
                               , '[^A-Z0-9]+'
                               , ' '
                               , 'g'
                             )
                           , ' '
                         )
                     )
                 )
               , '|'
             ) LIKE '%|' || LTRIM(
                              REGEXP_REPLACE(
                                  UPPER(COALESCE(a.expected_postal_code, ''))
                                , '[^A-Z0-9]'
                                , ''
                                , 'g'
                              )
                            , '0'
                          ) || '|%'
         AND (
                 COALESCE(TRIM(a.expected_region), '') = ''
              OR UPPER(COALESCE(a.expected_region, ''))                            = UPPER(COALESCE(a.expected_city, ''))
              OR REGEXP_REPLACE(
                     UPPER(COALESCE(w.default_address, ''))
                   , '[^A-Z0-9]'
                   , ''
                   , 'g'
                 ) LIKE '%' || REGEXP_REPLACE(
                                   UPPER(COALESCE(a.expected_region, ''))
                                 , '[^A-Z0-9]'
                                 , ''
                                 , 'g'
                               ) || '%'
              OR EXISTS (
                    SELECT 1
                      FROM cte_region_abbreviation_map                               m
                     WHERE m.country_name                                             = UPPER(COALESCE(a.expected_country, ''))
                       AND m.full_region_name                                         = UPPER(COALESCE(a.expected_region, ''))
                       AND ARRAY_TO_STRING(
                               ARRAY_SORT(
                                   ARRAY_DISTINCT(
                                       STRING_SPLIT(
                                           REGEXP_REPLACE(
                                               REPLACE(
                                                   REPLACE(
                                                       REPLACE(
                                                           UPPER(COALESCE(w.default_address, ''))
                                                         , 'CO.'
                                                         , ' '
                                                       )
                                                     , 'N°'
                                                     , ' '
                                                   )
                                                 , 'º'
                                                 , ' '
                                               )
                                             , '[^A-Z0-9]+'
                                             , ' '
                                             , 'g'
                                           )
                                         , ' '
                                       )
                                   )
                               )
                             , '|'
                           ) LIKE '%|' || m.region_abbreviation || '|%'
                  )
              OR UPPER(COALESCE(a.expected_country, '')) IN ('AUSTRIA', 'COSTA RICA', 'FRANCE', 'GERMANY')
             )
        THEN 'FORMAT_VARIATION'
        WHEN COALESCE(a.expected_value, '') <> ''
         AND COALESCE(w.default_address, '') <> ''
         AND UPPER(COALESCE(a.expected_country, ''))                                = 'AUSTRALIA'
         AND (
                 COALESCE(TRIM(a.expected_postal_code), '') = ''
              OR LTRIM(
                     REGEXP_EXTRACT(
                         UPPER(COALESCE(w.default_address, ''))
                       , '([^\n]+?)\s+([A-Z]{2,3})\s+([A-Z0-9-]{3,})(?:\n|$)'
                       , 3
                     )
                   , '0'
                 )                                                                  = LTRIM(
                                                                                         UPPER(COALESCE(a.expected_postal_code, ''))
                                                                                       , '0'
                                                                                     )
             )
         AND (
                 COALESCE(TRIM(a.expected_region), '') = ''
              OR EXISTS (
                    SELECT 1
                      FROM cte_region_abbreviation_map                               m
                     WHERE m.country_name                                             = 'AUSTRALIA'
                       AND m.full_region_name                                         = UPPER(COALESCE(a.expected_region, ''))
                       AND m.region_abbreviation                                      = REGEXP_EXTRACT(
                                                                                           UPPER(COALESCE(w.default_address, ''))
                                                                                         , '([^\n]+?)\s+([A-Z]{2,3})\s+([A-Z0-9-]{3,})(?:\n|$)'
                                                                                         , 2
                                                                                       )
                  )
             )
         AND (
                 REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n]+?)\s+([A-Z]{2,3})\s+([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_city, ''))
              OR REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n]+?)\s+([A-Z]{2,3})\s+([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_address_line_2, ''))
             )
        THEN 'FORMAT_VARIATION'
        WHEN COALESCE(a.expected_value, '') <> ''
         AND COALESCE(w.default_address, '') <> ''
         AND UPPER(COALESCE(a.expected_country, ''))                                = 'UNITED STATES OF AMERICA'
         AND NULLIF(TRIM(COALESCE(a.expected_postal_code, '')), '')                IS NOT NULL
         AND NULLIF(
                 REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 2
                 )
               , ''
             )                                                                     IS NOT NULL
         AND (
                 REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_city, ''))
              OR REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_address_line_1, ''))
              OR REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_address_line_2, ''))
              OR REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_address_line_3, ''))
              OR REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_address_line_4, ''))
             )
         AND LTRIM(
                 REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 3
                 )
               , '0'
             )                                                                      = LTRIM(
                                                                                         UPPER(COALESCE(a.expected_postal_code, ''))
                                                                                       , '0'
                                                                                     )
         AND (
                 COALESCE(TRIM(a.expected_region), '') = ''
              OR EXISTS (
                    SELECT 1
                      FROM cte_region_abbreviation_map                               m
                     WHERE m.country_name                                             = 'UNITED STATES OF AMERICA'
                       AND m.full_region_name                                         = UPPER(COALESCE(a.expected_region, ''))
                       AND m.region_abbreviation                                      = REGEXP_EXTRACT(
                                                                                           UPPER(COALESCE(w.default_address, ''))
                                                                                         , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                                                                                         , 2
                                                                                       )
                  )
             )
        THEN 'FORMAT_VARIATION'
         ELSE 'VALUE_MISMATCH'
       END                                                                         comparison_class
     , CASE
         WHEN NULLIF(TRIM(COALESCE(a.expected_value, '')), '') IS NULL
          AND NULLIF(TRIM(COALESCE(w.default_address, '')), '') IS NULL
         THEN 'Primary address is blank in both sent transform output and loaded extract'
         WHEN COALESCE(TRIM(a.expected_value), '') = COALESCE(TRIM(w.default_address), '')
         THEN 'Primary address matches exactly between sent transform output and loaded extract'
         WHEN COALESCE(a.expected_value, '') <> ''
          AND COALESCE(w.default_address, '') <> ''
          AND a.expected_value_norm = REGEXP_REPLACE(
                  UPPER(COALESCE(w.default_address, ''))
                , '[^A-Z0-9]'
                , ''
                , 'g'
              )
         THEN 'Primary address differs only by line breaks, punctuation, or whitespace compaction'
         WHEN COALESCE(a.expected_value, '') <> ''
          AND COALESCE(w.default_address, '') <> ''
          AND a.expected_value_token_signature = ARRAY_TO_STRING(
                  ARRAY_SORT(
                      ARRAY_DISTINCT(
                          STRING_SPLIT(
                          REGEXP_REPLACE(
                              REPLACE(
                                  REPLACE(
                                      REPLACE(
                                          UPPER(COALESCE(w.default_address, ''))
                                        , 'CO.'
                                        , ' '
                                      )
                                    , 'N°'
                                    , ' '
                                  )
                                , 'º'
                                , ' '
                              )
                            , '[^A-Z0-9]+'
                            , ' '
                            , 'g'
                          )
                        , ' '
                      )
                      )
                  )
                , '|'
              )
         THEN 'Primary address contains the same components but the extract reorders or compacts them into different lines'
         WHEN COALESCE(a.expected_value, '') <> ''
          AND COALESCE(w.default_address, '') <> ''
          AND a.expected_value_token_signature_zero_norm = ARRAY_TO_STRING(
                  ARRAY_SORT(
                      ARRAY_DISTINCT(
                          STRING_SPLIT(
                          REGEXP_REPLACE(
                              REGEXP_REPLACE(
                                  REPLACE(
                                      REPLACE(
                                          REPLACE(
                                              UPPER(COALESCE(w.default_address, ''))
                                            , 'CO.'
                                            , ' '
                                          )
                                        , 'N°'
                                        , ' '
                                      )
                                    , 'º'
                                    , ' '
                                  )
                                , '[^A-Z0-9]+'
                                , ' '
                                , 'g'
                              )
                            , '(^| )0+([0-9]{3,})( |$)'
                            , '\\1\\2\\3'
                            , 'g'
                          )
                        , ' '
                      )
                      )
                  )
                , '|'
              )
         THEN 'Primary address differs only by extract-side formatting and/or dropped leading zeros in postal code tokens'
         WHEN COALESCE(a.expected_value, '') <> ''
          AND COALESCE(w.default_address, '') <> ''
          AND EXISTS (
                SELECT 1
                  FROM cte_region_abbreviation_map                                 m
                 WHERE m.country_name                                               = UPPER(COALESCE(a.expected_country, ''))
                   AND m.full_region_name                                           = UPPER(COALESCE(a.expected_region, ''))
                   AND ARRAY_TO_STRING(
                           ARRAY_SORT(
                               ARRAY_DISTINCT(
                                   STRING_SPLIT(
                                   REGEXP_REPLACE(
                                       REPLACE(
                                           REPLACE(
                                               REPLACE(
                                                   UPPER(
                                                       COALESCE(
                                                           CONCAT_WS(
                                                               CHR(10)
                                                             , a.expected_address_line_1
                                                             , a.expected_address_line_2
                                                             , a.expected_address_line_3
                                                             , a.expected_address_line_4
                                                             , a.expected_city
                                                             , m.region_abbreviation
                                                             , a.expected_postal_code
                                                             , a.expected_country
                                                           )
                                                         , ''
                                                       )
                                                   )
                                                 , 'CO.'
                                                 , ' '
                                               )
                                             , 'N°'
                                             , ' '
                                           )
                                         , 'º'
                                         , ' '
                                       )
                                     , '[^A-Z0-9]+'
                                     , ' '
                                     , 'g'
                                   )
                                 , ' '
                               )
                               )
                           )
                         , '|'
                       ) = ARRAY_TO_STRING(
                               ARRAY_SORT(
                                   ARRAY_DISTINCT(
                                       STRING_SPLIT(
                                       REGEXP_REPLACE(
                                           REPLACE(
                                               REPLACE(
                                                   REPLACE(
                                                       UPPER(COALESCE(w.default_address, ''))
                                                     , 'CO.'
                                                     , ' '
                                                   )
                                                 , 'N°'
                                                 , ' '
                                               )
                                             , 'º'
                                             , ' '
                                           )
                                         , '[^A-Z0-9]+'
                                         , ' '
                                         , 'g'
                                       )
                                     , ' '
                                   )
                                   )
                               )
                             , '|'
                           )
              )
         THEN 'Primary address differs only by accepted region/province abbreviation in the loaded extract'
        WHEN COALESCE(a.expected_value, '') <> ''
         AND COALESCE(w.default_address, '') <> ''
         AND COALESCE(TRIM(a.expected_postal_code), '')                            <> ''
         AND (
                 COALESCE(TRIM(a.expected_city), '') <> ''
              OR COALESCE(TRIM(a.expected_address_line_2), '') <> ''
             )
         AND REGEXP_REPLACE(
                 UPPER(COALESCE(w.default_address, ''))
               , '[^A-Z0-9]'
               , ''
               , 'g'
             ) LIKE '%' || REGEXP_REPLACE(
                               UPPER(COALESCE(a.expected_address_line_1, ''))
                             , '[^A-Z0-9]'
                             , ''
                             , 'g'
                           ) || '%'
         AND (
                 COALESCE(TRIM(a.expected_address_line_2), '') = ''
              OR REGEXP_REPLACE(
                     UPPER(COALESCE(w.default_address, ''))
                   , '[^A-Z0-9]'
                   , ''
                   , 'g'
                 ) LIKE '%' || REGEXP_REPLACE(
                                   UPPER(COALESCE(a.expected_address_line_2, ''))
                                 , '[^A-Z0-9]'
                                 , ''
                                 , 'g'
                               ) || '%'
             )
         AND (
                 COALESCE(TRIM(a.expected_address_line_3), '') = ''
              OR REGEXP_REPLACE(
                     UPPER(COALESCE(w.default_address, ''))
                   , '[^A-Z0-9]'
                   , ''
                   , 'g'
                 ) LIKE '%' || REGEXP_REPLACE(
                                   UPPER(COALESCE(a.expected_address_line_3, ''))
                                 , '[^A-Z0-9]'
                                 , ''
                                 , 'g'
                               ) || '%'
             )
         AND (
                 COALESCE(TRIM(a.expected_address_line_4), '') = ''
              OR REGEXP_REPLACE(
                     UPPER(COALESCE(w.default_address, ''))
                   , '[^A-Z0-9]'
                   , ''
                   , 'g'
                 ) LIKE '%' || REGEXP_REPLACE(
                                   UPPER(COALESCE(a.expected_address_line_4, ''))
                                 , '[^A-Z0-9]'
                                 , ''
                                 , 'g'
                               ) || '%'
             )
         AND (
                 COALESCE(TRIM(a.expected_city), '') <> ''
              AND REGEXP_REPLACE(
                     UPPER(COALESCE(w.default_address, ''))
                   , '[^A-Z0-9]'
                   , ''
                   , 'g'
                 ) LIKE '%' || REGEXP_REPLACE(
                                   UPPER(COALESCE(a.expected_city, ''))
                                 , '[^A-Z0-9]'
                                 , ''
                                 , 'g'
                               ) || '%'
              OR COALESCE(TRIM(a.expected_city), '') = ''
             )
         AND (
                 COALESCE(TRIM(a.expected_address_line_2), '') <> ''
              OR COALESCE(TRIM(a.expected_city), '') <> ''
             )
         AND ARRAY_TO_STRING(
                 ARRAY_SORT(
                     ARRAY_DISTINCT(
                         STRING_SPLIT(
                             REGEXP_REPLACE(
                                 REPLACE(
                                     REPLACE(
                                         REPLACE(
                                             UPPER(COALESCE(w.default_address, ''))
                                           , 'CO.'
                                           , ' '
                                         )
                                       , 'N°'
                                       , ' '
                                     )
                                   , 'º'
                                   , ' '
                                 )
                               , '[^A-Z0-9]+'
                               , ' '
                               , 'g'
                             )
                           , ' '
                         )
                     )
                 )
               , '|'
             ) LIKE '%|' || LTRIM(
                              REGEXP_REPLACE(
                                  UPPER(COALESCE(a.expected_postal_code, ''))
                                , '[^A-Z0-9]'
                                , ''
                                , 'g'
                              )
                            , '0'
                          ) || '|%'
         AND (
                 COALESCE(TRIM(a.expected_region), '') = ''
              OR UPPER(COALESCE(a.expected_region, ''))                            = UPPER(COALESCE(a.expected_city, ''))
              OR REGEXP_REPLACE(
                     UPPER(COALESCE(w.default_address, ''))
                   , '[^A-Z0-9]'
                   , ''
                   , 'g'
                 ) LIKE '%' || REGEXP_REPLACE(
                                   UPPER(COALESCE(a.expected_region, ''))
                                 , '[^A-Z0-9]'
                                 , ''
                                 , 'g'
                               ) || '%'
              OR EXISTS (
                    SELECT 1
                      FROM cte_region_abbreviation_map                               m
                     WHERE m.country_name                                             = UPPER(COALESCE(a.expected_country, ''))
                       AND m.full_region_name                                         = UPPER(COALESCE(a.expected_region, ''))
                       AND ARRAY_TO_STRING(
                               ARRAY_SORT(
                                   ARRAY_DISTINCT(
                                       STRING_SPLIT(
                                           REGEXP_REPLACE(
                                               REPLACE(
                                                   REPLACE(
                                                       REPLACE(
                                                           UPPER(COALESCE(w.default_address, ''))
                                                         , 'CO.'
                                                         , ' '
                                                       )
                                                     , 'N°'
                                                     , ' '
                                                   )
                                                 , 'º'
                                                 , ' '
                                               )
                                             , '[^A-Z0-9]+'
                                             , ' '
                                             , 'g'
                                           )
                                         , ' '
                                       )
                                   )
                               )
                             , '|'
                           ) LIKE '%|' || m.region_abbreviation || '|%'
                  )
              OR UPPER(COALESCE(a.expected_country, '')) IN ('AUSTRIA', 'COSTA RICA', 'FRANCE', 'GERMANY')
             )
        THEN 'Primary address preserves all key components and only compacts postcode/city layout or accepted region display in the loaded extract'
        WHEN COALESCE(a.expected_value, '') <> ''
         AND COALESCE(w.default_address, '') <> ''
         AND UPPER(COALESCE(a.expected_country, ''))                                = 'AUSTRALIA'
         AND (
                 COALESCE(TRIM(a.expected_postal_code), '') = ''
              OR LTRIM(
                     REGEXP_EXTRACT(
                         UPPER(COALESCE(w.default_address, ''))
                       , '([^\n]+?)\s+([A-Z]{2,3})\s+([A-Z0-9-]{3,})(?:\n|$)'
                       , 3
                     )
                   , '0'
                 )                                                                  = LTRIM(
                                                                                         UPPER(COALESCE(a.expected_postal_code, ''))
                                                                                       , '0'
                                                                                     )
             )
         AND (
                 COALESCE(TRIM(a.expected_region), '') = ''
              OR EXISTS (
                    SELECT 1
                      FROM cte_region_abbreviation_map                               m
                     WHERE m.country_name                                             = 'AUSTRALIA'
                       AND m.full_region_name                                         = UPPER(COALESCE(a.expected_region, ''))
                       AND m.region_abbreviation                                      = REGEXP_EXTRACT(
                                                                                           UPPER(COALESCE(w.default_address, ''))
                                                                                         , '([^\n]+?)\s+([A-Z]{2,3})\s+([A-Z0-9-]{3,})(?:\n|$)'
                                                                                         , 2
                                                                                       )
                  )
             )
         AND (
                 REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n]+?)\s+([A-Z]{2,3})\s+([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_city, ''))
              OR REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n]+?)\s+([A-Z]{2,3})\s+([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_address_line_2, ''))
             )
        THEN 'Primary Australian address preserves the same locality and only compacts state abbreviation and postcode formatting'
        WHEN COALESCE(a.expected_value, '') <> ''
         AND COALESCE(w.default_address, '') <> ''
         AND UPPER(COALESCE(a.expected_country, ''))                                = 'UNITED STATES OF AMERICA'
         AND NULLIF(TRIM(COALESCE(a.expected_postal_code, '')), '')                IS NOT NULL
         AND NULLIF(
                 REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 2
                 )
               , ''
             )                                                                     IS NOT NULL
         AND (
                 REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_city, ''))
              OR REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_address_line_1, ''))
              OR REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_address_line_2, ''))
              OR REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_address_line_3, ''))
              OR REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 1
                 )                                                                  = UPPER(COALESCE(a.expected_address_line_4, ''))
             )
         AND LTRIM(
                 REGEXP_EXTRACT(
                     UPPER(COALESCE(w.default_address, ''))
                   , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                   , 3
                 )
               , '0'
             )                                                                      = LTRIM(
                                                                                         UPPER(COALESCE(a.expected_postal_code, ''))
                                                                                       , '0'
                                                                                     )
         AND (
                 COALESCE(TRIM(a.expected_region), '') = ''
              OR EXISTS (
                    SELECT 1
                      FROM cte_region_abbreviation_map                               m
                     WHERE m.country_name                                             = 'UNITED STATES OF AMERICA'
                       AND m.full_region_name                                         = UPPER(COALESCE(a.expected_region, ''))
                       AND m.region_abbreviation                                      = REGEXP_EXTRACT(
                                                                                           UPPER(COALESCE(w.default_address, ''))
                                                                                         , '([^\n,]+),\s*([A-Z]{2})\s*([A-Z0-9-]{3,})(?:\n|$)'
                                                                                         , 2
                                                                                       )
                  )
             )
        THEN 'Primary US address preserves the same locality/postcode and only compacts city, state abbreviation, and postal code formatting'
         WHEN COALESCE(a.expected_value, '') = ''
          AND COALESCE(w.default_address, '') <> ''
         THEN 'Loaded extract has a default address but the sent primary address is blank'
         WHEN COALESCE(a.expected_value, '') <> ''
          AND COALESCE(w.default_address, '') = ''
         THEN 'Sent primary address is populated but the loaded extract default address is blank'
         ELSE 'Primary address differs after normalization and should be reviewed'
       END                                                                         comparison_reason
  FROM cte_bridge                                                                  t
       LEFT OUTER JOIN
       cte_primary_address                                                         a
          ON a.expected_customer_id                                                = t.expected_customer_id
       LEFT OUTER JOIN
       src_fin_workday_extract_ccs_customer                                        w
          ON w.customer_id                                                         = t.workday_customer_id

 UNION ALL
SELECT t.customer_number
     , t.source_row_count
     , t.source_business_units
     , t.source_customer_names
     , t.source_countries
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , 'default_phone_number'
     , p.expected_value
     , w.default_phone_number
     , CASE
         WHEN NULLIF(TRIM(COALESCE(p.expected_value, '')), '') IS NULL
          AND NULLIF(TRIM(COALESCE(w.default_phone_number, '')), '') IS NULL
         THEN 'BOTH_BLANK'
         WHEN COALESCE(TRIM(p.expected_value), '') = COALESCE(TRIM(w.default_phone_number), '')
         THEN 'EXACT_MATCH'
         WHEN COALESCE(p.expected_extract_digits, '') <> ''
          AND COALESCE(p.expected_extract_digits, '') = REGEXP_REPLACE(
                  COALESCE(w.default_phone_number, '')
                , '[^0-9]'
                , ''
                , 'g'
              )
         THEN 'FORMAT_VARIATION'
         ELSE 'VALUE_MISMATCH'
       END                                                                         comparison_class
     , CASE
         WHEN NULLIF(TRIM(COALESCE(p.expected_value, '')), '') IS NULL
          AND NULLIF(TRIM(COALESCE(w.default_phone_number, '')), '') IS NULL
         THEN 'Primary phone is blank in both sent transform output and loaded extract'
         WHEN COALESCE(TRIM(p.expected_value), '') = COALESCE(TRIM(w.default_phone_number), '')
         THEN 'Primary phone matches exactly between sent transform output and loaded extract'
         WHEN COALESCE(p.expected_extract_digits, '') <> ''
          AND COALESCE(p.expected_extract_digits, '') = REGEXP_REPLACE(
                  COALESCE(w.default_phone_number, '')
                , '[^0-9]'
                , ''
                , 'g'
              )
         THEN 'Loaded extract adds display formatting and/or country code to the same sent phone number'
         WHEN COALESCE(p.expected_value, '') = ''
          AND COALESCE(w.default_phone_number, '') <> ''
         THEN 'Loaded extract has a default phone number but the sent primary phone is blank'
         WHEN COALESCE(p.expected_value, '') <> ''
          AND COALESCE(w.default_phone_number, '') = ''
         THEN 'Sent primary phone is populated but the loaded extract default phone number is blank'
         ELSE 'Primary phone differs after normalization and should be reviewed'
       END                                                                         comparison_reason
  FROM cte_bridge                                                                  t
       LEFT OUTER JOIN
       cte_primary_phone                                                           p
          ON p.expected_customer_id                                                = t.expected_customer_id
       LEFT OUTER JOIN
       src_fin_workday_extract_ccs_customer                                        w
          ON w.customer_id                                                         = t.workday_customer_id
;

DROP TABLE IF EXISTS reconcile_customer_ccs_contact_summary
;
CREATE TABLE reconcile_customer_ccs_contact_summary
    AS
SELECT t.attribute_name
     , t.comparison_class
     , COUNT(*)                                                                    row_count
  FROM reconcile_customer_ccs_contact_detail                                       t
 GROUP BY
       t.attribute_name
     , t.comparison_class
 ORDER BY
       t.attribute_name
     , CASE t.comparison_class
         WHEN 'EXACT_MATCH'      THEN 1
         WHEN 'FORMAT_VARIATION' THEN 2
         WHEN 'BOTH_BLANK'       THEN 3
         ELSE 4
       END
;

DROP TABLE IF EXISTS reconcile_customer_ccs_contact_extract_support
;
CREATE TABLE reconcile_customer_ccs_contact_extract_support
    AS
SELECT 'default_address_rows_in_extract'                                           metric
     , COUNT(*)                                                                    row_count
     , 'SUPPORTED'                                                                 status
  FROM src_fin_workday_extract_ccs_customer
 WHERE NULLIF(TRIM(default_address), '')                                           IS NOT NULL

 UNION ALL
SELECT 'default_phone_number_rows_in_extract'
     , COUNT(*)
     , 'SUPPORTED'
  FROM src_fin_workday_extract_ccs_customer
 WHERE NULLIF(TRIM(default_phone_number), '')                                      IS NOT NULL

 UNION ALL
SELECT 'customer_contacts_rows_in_extract'
     , COUNT(*)
     , CASE WHEN COUNT(*) = 0 THEN 'EMAIL_NOT_SUPPORTED_IN_CURRENT_EXTRACT' ELSE 'REVIEW' END
  FROM src_fin_workday_extract_ccs_customer
 WHERE NULLIF(TRIM(customer_contacts), '')                                         IS NOT NULL

 UNION ALL
SELECT 'sent_email_customer_ids_in_transform'
     , COUNT(DISTINCT customer_id)
     , CASE
         WHEN (SELECT COUNT(*)
                 FROM src_fin_workday_extract_ccs_customer
                WHERE NULLIF(TRIM(customer_contacts), '') IS NOT NULL) = 0
         THEN 'No populated extract email field to compare'
         ELSE 'REVIEW'
       END
  FROM workday_customer_email
;

DROP TABLE IF EXISTS reconcile_customer_ccs_contact_review_case_summary
;
CREATE TABLE reconcile_customer_ccs_contact_review_case_summary
    AS
SELECT ROW_NUMBER() OVER(
           ORDER BY CASE
                      WHEN t.actual_value IS NULL OR TRIM(t.actual_value) = ''
                      THEN 1
                      WHEN UPPER(COALESCE(t.expected_value, '')) LIKE '%CORNWALL%'
                       AND UPPER(COALESCE(t.actual_value, '')) LIKE '%COVENTRY%'
                      THEN 2
                      WHEN UPPER(COALESCE(t.expected_value, '')) LIKE '%UNITED STATES OF AMERICA%'
                       AND REGEXP_MATCHES(COALESCE(t.actual_value, ''), ',\s*[A-Z]{2}\s*[A-Z0-9-]{3,}')
                      THEN 3
                      WHEN REGEXP_MATCHES(COALESCE(t.actual_value, ''), '(^|\n)0*[0-9]{3,}\s+[A-Za-z]')
                      THEN 4
                      WHEN UPPER(COALESCE(t.expected_value, '')) LIKE '%AUSTRALIA%'
                        OR UPPER(COALESCE(t.expected_value, '')) LIKE '%CANADA%'
                        OR UPPER(COALESCE(t.expected_value, '')) LIKE '%BRAZIL%'
                        OR UPPER(COALESCE(t.expected_value, '')) LIKE '%ITALY%'
                        OR UPPER(COALESCE(t.expected_value, '')) LIKE '%SWITZERLAND%'
                      THEN 5
                      ELSE 6
                    END
                  , t.expected_customer_id
                  , t.customer_number
       )                                                                            review_order
     , CASE
         WHEN t.actual_value IS NULL OR TRIM(t.actual_value) = ''
         THEN 'blank_extract_default_address'
         WHEN UPPER(COALESCE(t.expected_value, '')) LIKE '%CORNWALL%'
          AND UPPER(COALESCE(t.actual_value, '')) LIKE '%COVENTRY%'
         THEN 'suspected_uk_locality_substitution'
         WHEN UPPER(COALESCE(t.expected_value, '')) LIKE '%UNITED STATES OF AMERICA%'
          AND REGEXP_MATCHES(COALESCE(t.actual_value, ''), ',\s*[A-Z]{2}\s*[A-Z0-9-]{3,}')
         THEN 'us_city_state_postcode_layout'
         WHEN REGEXP_MATCHES(COALESCE(t.actual_value, ''), '(^|\n)0*[0-9]{3,}\s+[A-Za-z]')
         THEN 'postcode_city_layout_compaction'
         WHEN UPPER(COALESCE(t.expected_value, '')) LIKE '%AUSTRALIA%'
           OR UPPER(COALESCE(t.expected_value, '')) LIKE '%CANADA%'
           OR UPPER(COALESCE(t.expected_value, '')) LIKE '%BRAZIL%'
           OR UPPER(COALESCE(t.expected_value, '')) LIKE '%ITALY%'
           OR UPPER(COALESCE(t.expected_value, '')) LIKE '%SWITZERLAND%'
         THEN 'country_specific_region_layout'
         ELSE 'manual_review'
       END                                                                         review_bucket
     , t.customer_number
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , t.expected_value
     , t.actual_value
     , t.comparison_reason
  FROM reconcile_customer_ccs_contact_detail                                       t
 WHERE t.attribute_name                                                            = 'default_address'
   AND t.comparison_class                                                          = 'VALUE_MISMATCH'
;

DROP TABLE IF EXISTS reconcile_customer_ccs_contact_review_pattern_summary
;
CREATE TABLE reconcile_customer_ccs_contact_review_pattern_summary
    AS
SELECT review_bucket
     , COUNT(*)                                                                    row_count
     , COUNT(DISTINCT expected_customer_id)                                        merged_customer_count
     , STRING_AGG(DISTINCT expected_customer_id, ';' ORDER BY expected_customer_id)
                                                                                    sample_customer_ids
  FROM reconcile_customer_ccs_contact_review_case_summary
 GROUP BY
       review_bucket
 ORDER BY
       CASE review_bucket
         WHEN 'blank_extract_default_address'       THEN 1
         WHEN 'suspected_uk_locality_substitution'  THEN 2
         WHEN 'us_city_state_postcode_layout'       THEN 3
         WHEN 'postcode_city_layout_compaction'     THEN 4
         WHEN 'country_specific_region_layout'      THEN 5
         ELSE 6
       END
     , review_bucket
;

INSERT INTO reconciliation_customer_ccs_issues
SELECT '6_contact_reconciliation'                                                  section
     , 'CONTACT_REVIEW'                                                            issue_type
     , t.customer_number
     , t.expected_customer_id
     , t.workday_customer_id
     , t.expected_reference_id
     , t.workday_reference_id
     , t.attribute_name
     , t.expected_value
     , t.actual_value
     , t.comparison_reason                                                         details
  FROM reconcile_customer_ccs_contact_detail                                       t
 WHERE t.comparison_class                                                          = 'VALUE_MISMATCH'
;
