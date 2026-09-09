DROP TABLE IF EXISTS validation_supplier_issues
;
CREATE TABLE validation_supplier_issues (
    table_name  VARCHAR
  , rule_type   VARCHAR
  , col         VARCHAR
  , supplier_id VARCHAR
  , detail      VARCHAR
  , message     VARCHAR
)
;
