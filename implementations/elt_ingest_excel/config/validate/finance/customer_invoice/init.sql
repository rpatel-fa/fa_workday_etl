DROP TABLE IF EXISTS validation_customer_invoice_issues
;
CREATE TABLE validation_customer_invoice_issues (
    table_name    VARCHAR
  , rule_type     VARCHAR
  , col           VARCHAR
  , invoice_id    VARCHAR
  , detail        VARCHAR
  , message       VARCHAR
)
;
