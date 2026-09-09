DROP TABLE IF EXISTS workday_customer_company_restrictions
;
CREATE TABLE workday_customer_company_restrictions
    AS
SELECT
       c.customer_id                   customer_id
     , c.nrm_customer_name             customer_name
     , c.nrm_agg_business_unit         restricted_to_companies
  FROM 
       src_fin_customer                c
 WHERE 1 = 2
;
