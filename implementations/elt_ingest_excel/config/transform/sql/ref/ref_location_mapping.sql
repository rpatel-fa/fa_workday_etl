DROP TABLE IF EXISTS ref_location_mapping
;
-- Maps Okta location values to 4 canonical locations
-- Rules:
--   • Only 4 valid target locations permitted: Homebased - SGP | Homebased - Wembley | St George's Park | Wembley Stadium
--   • NULL = EXCLUDE from contingent worker population (non-physical locations, external entities, or ambiguous values)
--   • 'County' explicitly excluded per business rule
--   • Cognizant and The FA Group mapped to Wembley Stadium (temporary assumption for testing)
CREATE TABLE ref_location_mapping
    AS
SELECT
       TRIM(source_column)              source_column
     , TRIM(source_value)               source_value
     , TRIM(target_value)               target_value
  FROM read_json_auto(
           'config/data/ref_location_mapping.json'
       )
;
