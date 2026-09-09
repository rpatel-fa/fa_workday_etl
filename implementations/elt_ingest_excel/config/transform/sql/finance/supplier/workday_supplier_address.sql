DROP TABLE IF EXISTS workday_supplier_address
;
CREATE TABLE workday_supplier_address 
    AS
-- Publish only the columns that exist on the Supplier Address workbook sheet, in
-- the workbook's exact order. workday_supplier_address_base carries an internal
-- city_subdivision field for validation / diagnostic use, but the 3+ supplier
-- workbook has no matching column; selecting * shifts every later address field
-- one column to the right during publish.
SELECT t.supplier_id
     , t.supplier_name
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
       -- The Supplier Address workbook exposes only one locality slot between City
       -- and Address Line #1: "Submunicipality". Workday validation for some
       -- countries (for example Finland / Slovakia / South Africa) calls this required
       -- component "City Subdivision 1". Publish our derived city_subdivision value into
       -- that workbook slot when no native submunicipality exists so the outbound file
       -- carries the locality Workday is actually asking for.
       -- EIB TYPE TAG (learned from the 2026-07-24 load): this value alone is NOT enough -
       -- the Submit Supplier EIB requires the paired component Type = 'CITY_SUBDIVISION_1'
       -- (the column immediately left of "Submunicipality Data"), exactly as address lines
       -- carry ADDRESS_LINE_n and subregions carry REGION_SUBDIVISION_1. Without that Type
       -- tag Workday reports "City Subdivision 1 required" even though the value is present.
       -- The 30-column publish template has no Type column, so the tag is applied during EIB
       -- transcription (see scripts/populate_eib_corrections.py).
     , COALESCE(t.submunicipality, t.city_subdivision)                          submunicipality
     , t.address_line_1
     , t.address_line_2
     , t.address_line_3
     , t.address_line_4
     , t.postal_code
     , t.public_flag
     , t.is_primary
     , t.effective_date
     , t.address_type
     , t.use_for
     , t.use_for_tenanted
     , t.address_comments
     , t.number_of_days
     , t.municipality_local
  FROM workday_supplier_address_base   t
 WHERE NOT (
       t.region                        IS NULL 
   AND t.country_code                  IN ('IT', 'BR', 'DO')
       )
 ORDER BY 
       supplier_id
     , is_primary DESC
     , address_id
;
