DROP TABLE IF EXISTS workday_supplier_alternate_name
;
CREATE TABLE workday_supplier_alternate_name
    AS
  WITH cte_supplier_alternate_name
    AS (
-- Short Name
SELECT
       t.supplier_id                                  supplier_id
     , t.nrm_supplier_name                            supplier_name
     , TRIM(t.vendor_short_name)                      alternate_name
     , 'Short Name'                                   alternate_name_usage_plus
  FROM src_fin_supplier                t
 WHERE t.vendor_short_name             IS NOT NULL
   AND TRIM(t.vendor_short_name)       != ''
   AND TRIM(t.vendor_short_name)       != t.nrm_supplier_name
UNION ALL
-- Check Name
SELECT
       t.supplier_id                                  supplier_id
     , t.nrm_supplier_name                            supplier_name
     , TRIM(t.vendor_check_name)                      alternate_name
     , 'Check Name'                                   alternate_name_usage_plus
  FROM src_fin_supplier                t
 WHERE t.vendor_check_name             IS NOT NULL
   AND TRIM(t.vendor_check_name)       != ''
   AND TRIM(t.vendor_check_name)       != t.nrm_supplier_name
   AND (
       TRIM(t.vendor_check_name)       != TRIM(t.vendor_short_name)
    OR t.vendor_short_name             IS NULL
       )
       )
SELECT *
  FROM cte_supplier_alternate_name
 WHERE 1 = 2
;
