DROP TABLE IF EXISTS workday_customer_address
;
CREATE TABLE workday_customer_address
    AS
-- Publish only the columns that exist on the Customer Address workbook sheet, in
-- the workbook's exact order. workday_customer_address_base carries an internal
-- city_subdivision field for validation / diagnostic use, but the Customer
-- workbook has no matching column; selecting * shifts every later address field
-- one column to the right during publish.
SELECT t.customer_id
     , t.customer_name
     , t.formatted_address
     , t.address_format_type
     , t.defaulted_business_site_address
     , t.delete_flag
     , t.do_not_replace_all
     , t.last_modified
     , t.descriptor
     , t.address_id
     , t.country
     , t.country_code
     , t.region
     , t.subregion
     , t.city
     -- The Customer Address workbook exposes only one locality slot between City
     -- and Address Line #1: "Submunicipality". Workday validation for some
     -- countries calls this required component "City Subdivision 1". Publish our
     -- derived city_subdivision value into that workbook slot when no native
     -- submunicipality exists so the outbound file carries the locality Workday
     -- is actually asking for.
     , COALESCE(t.submunicipality, t.city_subdivision)                          submunicipality
     , t.address_line_1
     , t.address_line_2
     , t.address_line_3
     , t.address_line_4
     , t.postal_code
     , t.is_public
     , t.is_primary
     , t.address_type
     , t.effective_date
     , t.use_for
     , t.use_for_tenanted
     , t.customer_comment
     , t.number_of_days
     , t.municipality_local
     , t.optional_address
  FROM workday_customer_address_base   t
 WHERE NOT (
       t.address_line_2                IS NULL 
   AND t.country_code                  = 'HU'
       )
 ORDER BY 
       customer_id
     , is_primary DESC
     , address_id
;
