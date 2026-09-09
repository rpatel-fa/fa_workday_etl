DROP TABLE IF EXISTS src_fin_wd_submit_customer_address
;
CREATE TABLE src_fin_wd_submit_customer_address
    AS
SELECT
       TRY_CAST(NULLIF(TRIM(t.spreadsheet_key), '') AS INTEGER)      spreadsheet_key
     , MAX(NULLIF(TRIM(t.customer_id), ''))                           customer_id
     , MAX(NULLIF(TRIM(t.customer_reference_id), ''))                 customer_reference_id
     , MAX(NULLIF(TRIM(t.customer_name), ''))                         customer_name
     , MAX(NULLIF(TRIM(t.country), ''))                               country
     , MAX(NULLIF(TRIM(t.municipality), ''))                          municipality
     , MAX(NULLIF(TRIM(t.country_city), ''))                          country_city
     , MAX(NULLIF(TRIM(t.iso_3166_1_alpha_2_code), ''))               iso_3166_1_alpha_2_code
     , MAX(TRY_CAST(NULLIF(TRIM(t.row_id_4), '') AS INTEGER))         row_id_country_city
     , MAX(NULLIF(TRIM(t.address_component_name), ''))                address_component_name
     , MAX(NULLIF(TRIM(t.submunicipality_data), ''))                  submunicipality_data
     , MAX(NULLIF(TRIM(t.country_region), ''))                        country_region
     , MAX(NULLIF(TRIM(t.country_region_descriptor), ''))             country_region_descriptor
     , MAX(TRY_CAST(NULLIF(TRIM(t.row_id_5), '') AS INTEGER))         row_id_country_region
     , MAX(NULLIF(TRIM(t.descriptor_1), ''))                          descriptor
     , MAX(NULLIF(TRIM(t.type_2), ''))                                descriptor_type
     , MAX(NULLIF(TRIM(t.subregion_data), ''))                        subregion_data
     , MAX(NULLIF(TRIM(t.postal_code), ''))                           postal_code
     , MAX(CASE
           WHEN UPPER(TRIM(COALESCE(t.type, ''))) = 'ADDRESS_LINE_1'
           THEN NULLIF(TRIM(t.address_line_data), '')
       END)                                                           address_line_1
     , MAX(CASE
           WHEN UPPER(TRIM(COALESCE(t.type, ''))) = 'ADDRESS_LINE_2'
           THEN NULLIF(TRIM(t.address_line_data), '')
       END)                                                           address_line_2
     , MAX(CASE
           WHEN UPPER(TRIM(COALESCE(t.type, ''))) = 'ADDRESS_LINE_3'
           THEN NULLIF(TRIM(t.address_line_data), '')
       END)                                                           address_line_3
     , MAX(CASE
           WHEN UPPER(TRIM(COALESCE(t.type, ''))) = 'ADDRESS_LINE_4'
           THEN NULLIF(TRIM(t.address_line_data), '')
       END)                                                           address_line_4
  FROM fin_wd_submit_customer                                        t
 WHERE TRY_CAST(NULLIF(TRIM(t.spreadsheet_key), '') AS INTEGER)      IS NOT NULL
 GROUP BY 1
;
