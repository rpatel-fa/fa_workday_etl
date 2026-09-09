DROP TABLE IF EXISTS workday_customer_currencies
;
CREATE TABLE workday_customer_currencies
    AS
SELECT
       c.customer_id                                  customer_id
     , c.nrm_customer_name                            customer_name
     , 'All Currencies'                               accepted_currencies_plus
     , c.nrm_currency_code                            default_currency
  FROM src_fin_customer                c
 ORDER BY 
       customer_id
;
