DROP TABLE IF EXISTS workday_supplier_currencies
;
CREATE TABLE workday_supplier_currencies 
    AS
SELECT
       t.supplier_id                   supplier_id
     , t.nrm_supplier_name             supplier_name
     , 'All Currencies'                accepted_currencies_plus
     , t.nrm_currency_code             default_currency
  FROM src_fin_supplier                t
 ORDER BY 
       t.supplier_id
;
