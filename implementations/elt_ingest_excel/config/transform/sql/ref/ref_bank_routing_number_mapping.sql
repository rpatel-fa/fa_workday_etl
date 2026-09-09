DROP TABLE IF EXISTS ref_bank_routing_number_mapping
;
-- US routing number -> bank name, sourced from the Workday tenant's "Tenant Default
-- Named Ranges" sheet (columns AT "Routing Numbers" / AU "Bank Names") in the cloned
-- supplier EIB workbook. This is the tenant-authoritative US bank list, so a US bank
-- name derived from the routing number here matches exactly what the EIB expects.
-- Routing numbers are normalised to 9 digits (US ABA) with the leading zero restored.
CREATE TABLE ref_bank_routing_number_mapping
    AS
SELECT
       TRIM(CAST(routing_number AS VARCHAR))    routing_number
     , TRIM(bank_name)                          bank_name
  FROM read_json_auto(
           'config/data/ref_bank_routing_number_mapping.json'
       )
;
