DROP TABLE IF EXISTS workday_customer_phone
;
CREATE TABLE workday_customer_phone AS
  WITH cte_customer_phone
    AS (
SELECT c.customer_id                         customer_id
     , c.nrm_customer_name                   customer_name
     , c.nrm_country_name                    customer_country_name
     , c.nrm_country_code                    customer_country_code
     , COALESCE(c.nrm_phone_code, '+44')     international_phone_code
     , c.nrm_agg_phone_1                     phone_raw
     , 'primary'                             phone_type
     , '_PH1'                                suffix
  FROM src_fin_customer                      c
 WHERE NULLIF(UPPER(TRIM(c.nrm_agg_phone_1)), '') IS NOT NULL
UNION ALL
SELECT c.customer_id                         customer_id
     , c.nrm_customer_name                   customer_name
     , c.nrm_country_name                    customer_country_name
     , c.nrm_country_code                    customer_country_code
     , COALESCE(c.nrm_phone_code, '+44')     international_phone_code
     , c.nrm_agg_phone_2                     phone_raw
     , 'secondary'                           phone_type
     , '_PH2'                                suffix
  FROM src_fin_customer                      c
 WHERE NULLIF(UPPER(TRIM(c.nrm_agg_phone_2)), '') IS NOT NULL
UNION ALL
SELECT c.customer_id                         customer_id
     , c.nrm_customer_name                   customer_name
     , c.nrm_country_name                    customer_country_name
     , c.nrm_country_code                    customer_country_code
     , COALESCE(c.nrm_phone_code, '+44')     international_phone_code
     , c.nrm_agg_phone_3                     phone_raw
     , 'tertiary'                            phone_type
     , '_PH3'                                suffix
  FROM src_fin_customer                      c
 WHERE NULLIF(UPPER(TRIM(c.nrm_agg_phone_3)), '') IS NOT NULL
       )
     , cte_phone_split
    AS (
SELECT DISTINCT
       p.customer_id                                                       customer_id
     , p.customer_name                                                     customer_name
     , p.customer_country_name                                             customer_country_name
     , p.customer_country_code                                             customer_country_code
     , REGEXP_REPLACE(TRIM(p.international_phone_code), '[^0-9]', '', 'g') international_phone_code
     , TRIM(u.phone)                                                       phone_value
     , CASE
         WHEN TRIM(u.phone) LIKE '+%'
         THEN '+' || REGEXP_REPLACE(TRIM(u.phone), '[^0-9]', '', 'g')
         ELSE REGEXP_REPLACE(TRIM(u.phone), '[^0-9]', '', 'g')
       END                                                                 phone_number_raw
     , p.phone_type                                                        phone_type
     , p.suffix                                                            suffix
  FROM cte_customer_phone p
     , UNNEST(STRING_SPLIT(p.phone_raw, ';')) u(phone)
       )
    , cte_cleaned
    AS (
SELECT t.*
     , CASE
         WHEN t.phone_number_raw_norm LIKE '+%'
           THEN t.phone_number_raw_norm
         WHEN t.phone_number_raw_norm LIKE '00%'
           THEN '+' || SUBSTR(t.phone_number_raw_norm, 3)
         WHEN REGEXP_REPLACE(t.phone_number_raw_norm, '^0+', '') LIKE t.international_phone_code || '%'
           THEN REGEXP_REPLACE(
                  REGEXP_REPLACE(
                    REGEXP_REPLACE(t.phone_number_raw_norm, '^0+', ''),
                    '^' || t.international_phone_code,
                    ''
                  ),
                  '^0+', ''
                )
         -- No 0/+ prefix and >10 digits: country code is embedded without + (e.g. 353xxxxxxx, 81xxxxxxxx)
         WHEN t.phone_number_raw_norm NOT LIKE '0%'
          AND LENGTH(t.phone_number_raw_norm) > 10
           THEN '+' || t.phone_number_raw_norm
         ELSE REGEXP_REPLACE(t.phone_number_raw_norm, '^0+', '')
       END                                                                 phone_number
  FROM (
        SELECT s.*
             , CASE
                 WHEN s.phone_number_raw LIKE '+0%'
                   THEN SUBSTR(s.phone_number_raw, 2)
                 WHEN s.phone_number_raw LIKE '+440%'
                   THEN '+44' || SUBSTR(s.phone_number_raw, 5)
                 ELSE s.phone_number_raw
               END                                                         phone_number_raw_norm
          FROM cte_phone_split s
         WHERE NULLIF(TRIM(s.phone_number_raw), '') IS NOT NULL
       )                                                                   t
       )
     , cte_phone_parse
    AS (
SELECT t.*
     , udf_parse_phone(t.phone_number, t.customer_country_code)            parsed_phone
  FROM cte_cleaned   t
       )
    , cte_phone_guess
   AS (
SELECT t.*
     , g.country_code                                                    phone_country_code_guess
     , g.country_name                                                    phone_country_name_guess
     , g.phone_code_digits                                               phone_country_phone_code_guess
     , CASE
         WHEN g.phone_code_digits IS NOT NULL
           THEN REGEXP_REPLACE(
                  REGEXP_REPLACE(
                    REGEXP_REPLACE(
                      REGEXP_REPLACE(t.phone_number, '^\+', ''),
                      '^' || g.phone_code_digits,
                      ''
                    ),
                    '^0',
                    ''
                  ),
                  '[^0-9]',
                  '',
                  'g'
                )
         ELSE t.parsed_phone.full_phone_number
       END                                                               full_phone_number
     , COALESCE(g.country_code, t.parsed_phone.phone_country_code)        phone_country_code
  FROM cte_phone_parse t
  LEFT JOIN LATERAL (
         SELECT
                r.country_code                                           country_code
              , r.country_name                                           country_name
              , REGEXP_REPLACE(TRIM(r.phone_code), '[^0-9]', '', 'g')     phone_code_digits
              , LENGTH(REGEXP_REPLACE(TRIM(r.phone_code), '[^0-9]', '', 'g')) code_len
           FROM ref_country r
          WHERE t.phone_number LIKE '+%'
            AND REGEXP_REPLACE(TRIM(r.phone_code), '[^0-9]', '', 'g') IS NOT NULL
            AND REGEXP_REPLACE(TRIM(r.phone_code), '[^0-9]', '', 'g') != ''
            AND REGEXP_REPLACE(t.phone_number, '^\+', '') LIKE REGEXP_REPLACE(TRIM(r.phone_code), '[^0-9]', '', 'g') || '%'
          ORDER BY code_len DESC
          LIMIT 1
       )                                                                   g
    ON TRUE
       )
     , cte_phone_rnk
    AS (
SELECT t.*
     , ROW_NUMBER() OVER (
           PARTITION BY t.customer_id
               ORDER BY CASE t.phone_type
                           WHEN 'primary'    THEN 1
                           WHEN 'secondary'  THEN 2
                           WHEN 'tertiary'   THEN 3
                           ELSE 5
                        END
                      , t.full_phone_number
       )                                     phone_rank
  FROM cte_phone_guess   t
 WHERE t.full_phone_number         IS NOT NULL
   AND LENGTH(t.full_phone_number) >= 7
   AND NOT regexp_matches(t.full_phone_number, '^0*12345')
       )
SELECT s.customer_id                                                       customer_id
     , s.customer_name                                                     customer_name
     , TRIM(s.customer_id) || s.suffix || '_' || ROW_NUMBER() OVER (
           PARTITION BY s.customer_id, s.suffix
               ORDER BY s.phone_number_raw
       )                                                                   phone_id
     , r.country_name                                                      phone_country
     , r.country_code                                                      country_code
     , REGEXP_REPLACE(TRIM(r.phone_code), '[^0-9]', '', 'g')               international_phone_code
     , CAST(NULL AS VARCHAR)                                               area_code
     , s.full_phone_number                                                 phone_number
     , s.full_phone_number                                                 formatted_phone_number
     , CAST(NULL AS VARCHAR)                                               phone_number_extension
     , s.parsed_phone.device_type                                          phone_device_type
     , 'Yes'                                                               is_public
     , CASE s.phone_rank WHEN 1 THEN 'Yes' ELSE 'No' END                   primary_flag
     , s.phone_type                                                        phone_type
     , CAST(NULL AS VARCHAR)                                               use_for
     , CAST(NULL AS VARCHAR)                                               use_for_tenanted
     , CAST(NULL AS VARCHAR)                                               tenant_formatted_phone
     , CAST(NULL AS VARCHAR)                                               international_formatted_phone
     , CAST(NULL AS VARCHAR)                                               delete_flag
     , CAST(NULL AS VARCHAR)                                               do_not_replace_all
     , CAST(NULL AS VARCHAR)                                               phone_comment
     , CAST(NULL AS VARCHAR)                                               phone
 FROM cte_phone_rnk                    s
      LEFT OUTER JOIN
      ref_country                      r
          ON r.country_code            = s.phone_country_code
 ORDER BY
       customer_id
     , primary_flag DESC
     , CASE phone_device_type
         WHEN 'Landline'   THEN 1
         WHEN 'Mobile'     THEN 2
         ELSE 4
       END
     , phone_id
;
