DROP TABLE IF EXISTS reconciliation_supplier_po_receipt_issues
;
CREATE TABLE reconciliation_supplier_po_receipt_issues (
    table_name  VARCHAR
  , rule_type   VARCHAR
  , col         VARCHAR
  , recon_key   VARCHAR
  , detail      VARCHAR
  , message     VARCHAR
)
;
