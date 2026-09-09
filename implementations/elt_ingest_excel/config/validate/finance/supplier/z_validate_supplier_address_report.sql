DROP TABLE IF EXISTS validate_supplier_address_report
;
CREATE TABLE validate_supplier_address_report 
    AS
SELECT s1.nrm_agg_supplier_number gp_supplier_number
     , w.supplier_id
     , w.supplier_name
     , w.country
     , w.country_code
     , w.region
     , w.city
     , w.address_line_1
     , w.address_line_2
     , w.address_line_3
     , w.address_line_4
     , w.postal_code
     , w.is_primary     
  FROM workday_supplier_address w
       INNER JOIN 
       src_fin_supplier s1
        ON s1.supplier_id = w.supplier_id
 WHERE w.supplier_id IN (
       SELECT supplier_id FROM validation_supplier_issues WHERE table_name = 'workday_supplier_address'
       )
ORDER BY 1
;

DROP TABLE IF EXISTS validate_supplier_address_report_base
;
CREATE TABLE validate_supplier_address_report_base
    AS
SELECT company
     , vendor_id
     , vendor_name
     , vendor_address_code_primary
     , address_1
     , address_2
     , address_3
     , city
     , county
     , post_code
     , country_code
     , country
  FROM src_fin_supplier_base b
       INNER JOIN (
SELECT DISTINCT
       TRIM(u.gp_supplier_number) gp_supplier_number
  FROM validate_supplier_address_report t
     , UNNEST(string_split(t.gp_supplier_number, '|')) u(gp_supplier_number)
 WHERE NULLIF(TRIM(u.gp_supplier_number), '') IS NOT NULL
       ) v
        ON TRIM(b.vendor_id) = v.gp_supplier_number
 ORDER BY
       country_code
     , vendor_id
;
