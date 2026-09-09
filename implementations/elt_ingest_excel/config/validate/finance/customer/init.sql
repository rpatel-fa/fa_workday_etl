DROP TABLE IF EXISTS validation_customer_issues
;
CREATE TABLE validation_customer_issues (
    table_name  VARCHAR
  , rule_type   VARCHAR
  , col         VARCHAR
  , customer_id VARCHAR
  , detail      VARCHAR
  , message     VARCHAR
)
;
