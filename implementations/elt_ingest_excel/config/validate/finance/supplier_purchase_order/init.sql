DROP TABLE IF EXISTS validation_supplier_purchase_order_issues
;
CREATE TABLE validation_supplier_purchase_order_issues (
    table_name    VARCHAR
  , rule_type     VARCHAR
  , col           VARCHAR
  , po_id         VARCHAR
  , detail        VARCHAR
  , message       VARCHAR
)
;
