DROP TABLE IF EXISTS validate_customer_invoice_adjustment_missing_adjustment_id
;
-- EIB rows with no derivable Customer Invoice Adjustment ID (column_h). The id
-- is the "...: <id>" suffix of wd_adjustment_transaction_needed; a blank here
-- means the source worklist text was malformed and the row would submit without
-- a document id. Expected to be empty.
CREATE TABLE validate_customer_invoice_adjustment_missing_adjustment_id
    AS
SELECT t.column_l                                  company
     , t.column_aj                                 wd_customer_id
     , t.column_h                                  adjustment_id
     , t.column_fb                                 extended_amount
  FROM workday_customer_invoice_adjustment          t
 WHERE NULLIF(TRIM(t.column_h), '')                IS NULL
;
