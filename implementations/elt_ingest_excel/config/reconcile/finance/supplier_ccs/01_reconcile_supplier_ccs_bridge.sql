-- Population bridge: GP-source suppliers <-> Workday extract, keyed on the GP
-- Vendor ID in its business-unit-suffixed form (nrm_supplier_number, e.g.
-- LON004-FA). The raw src_fin_supplier_base.vendor_id (LON004) does NOT carry the
-- BU suffix and is reused across FA/NFC/WNSL, so it cannot be the reconcile key —
-- the transform's nrm_agg_supplier_number is the authoritative source key and is
-- exactly what Workday stores in reference_id.

-- GP-source population at supplier_number grain (exploded from the merged winners).
DROP TABLE IF EXISTS reconcile_supplier_ccs_source_profile
;
CREATE TABLE reconcile_supplier_ccs_source_profile
    AS
SELECT TRIM(u.supplier_number)                                                       supplier_number
     , c.supplier_id                                                                 expected_supplier_id
     , c.nrm_supplier_name                                                           source_supplier_name
     , c.nrm_business_unit                                                           source_business_units
     , c.nrm_country_name                                                            source_country
     , c.nrm_agg_supplier_number                                                     expected_reference_id
  FROM src_fin_supplier                                                              c
     , UNNEST(STRING_SPLIT(c.nrm_agg_supplier_number, '|'))                          u(supplier_number)
 WHERE c.rnk_score                                                                   = 1
   AND NULLIF(TRIM(u.supplier_number), '')                                           IS NOT NULL
;

-- Merge bridge: supplier_number -> its merged transform supplier (parallel to the
-- customer_ccs structure; for suppliers it is the same source as the profile).
DROP TABLE IF EXISTS reconcile_supplier_ccs_source_merge_bridge
;
CREATE TABLE reconcile_supplier_ccs_source_merge_bridge
    AS
SELECT DISTINCT
       supplier_number
     , expected_supplier_id
     , source_supplier_name                                                          expected_supplier_name
     , expected_reference_id
  FROM reconcile_supplier_ccs_source_profile
;

-- Workday extract population at supplier_number grain (exploded reference_id).
DROP TABLE IF EXISTS reconcile_supplier_ccs_workday_reference_bridge
;
CREATE TABLE reconcile_supplier_ccs_workday_reference_bridge
    AS
SELECT DISTINCT
       TRIM(u.supplier_number)                                                       supplier_number
     , w.supplier_id                                                                 workday_supplier_id
     , w.supplier                                                                    workday_supplier_name
     , w.reference_id                                                                workday_reference_id
  FROM fin_workday_extract_supplier                                                  w
     , UNNEST(STRING_SPLIT(w.reference_id, '|'))                                     u(supplier_number)
 WHERE NULLIF(TRIM(u.supplier_number), '')                                           IS NOT NULL
;

DROP TABLE IF EXISTS reconcile_supplier_ccs_bridge
;
CREATE TABLE reconcile_supplier_ccs_bridge
    AS
SELECT COALESCE(s.supplier_number, w.supplier_number)                                supplier_number
     , s.source_supplier_name
     , s.source_business_units
     , s.source_country
     , s.expected_supplier_id
     , s.expected_reference_id
     , w.workday_supplier_id
     , w.workday_supplier_name
     , w.workday_reference_id
     , CASE WHEN s.supplier_number IS NOT NULL THEN 'Y' ELSE 'N' END                 in_source_gp
     , CASE WHEN w.supplier_number IS NOT NULL THEN 'Y' ELSE 'N' END                 in_workday
  FROM reconcile_supplier_ccs_source_profile                                         s
       FULL OUTER JOIN
       reconcile_supplier_ccs_workday_reference_bridge                               w
          ON w.supplier_number                                                       = s.supplier_number
;
