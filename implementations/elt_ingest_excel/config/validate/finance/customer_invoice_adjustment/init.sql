DROP TABLE IF EXISTS validation_customer_invoice_adjustment_issues
;
CREATE TABLE validation_customer_invoice_adjustment_issues (
    table_name    VARCHAR
  , rule_type     VARCHAR
  , col           VARCHAR
  , invoice_id    VARCHAR
  , detail        VARCHAR
  , message       VARCHAR
)
;
