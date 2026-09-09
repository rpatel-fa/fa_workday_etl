DROP TABLE IF EXISTS workday_supplier_company_restrictions
;
CREATE TABLE workday_supplier_company_restrictions
    AS
SELECT
       t.supplier_id                   supplier_id
     , t.nrm_supplier_name             supplier_name
     , t.nrm_agg_business_unit         restricted_to_companies
  FROM 
       src_fin_supplier                t
 WHERE 1 = 2
;
