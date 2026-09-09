DROP TABLE IF EXISTS reconciliation_supplier_issues
;
CREATE TABLE reconciliation_supplier_issues (
    table_name  VARCHAR
  , rule_type   VARCHAR
  , col         VARCHAR
  , supplier_id VARCHAR
  , detail      VARCHAR
  , message     VARCHAR
)
;
