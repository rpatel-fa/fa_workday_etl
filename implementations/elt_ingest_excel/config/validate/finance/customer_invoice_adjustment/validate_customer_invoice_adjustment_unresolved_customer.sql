DROP TABLE IF EXISTS validate_customer_invoice_adjustment_unresolved_customer
;
-- Adjustment rows whose Workday customer id does not resolve in the Workday
-- customer extract. Expected to be empty (the source is pre-resolved); any row
-- here is a customer that would fail to load and must be triaged before submit.
CREATE TABLE validate_customer_invoice_adjustment_unresolved_customer
    AS
SELECT a.company
     , a.customer_number
     , a.customer_name
     , a.wd_customer_id
     , a.document_number
     , a.nrm_adjustment_id
  FROM src_fin_customer_invoice_adjustment          a
       LEFT OUTER JOIN
       fin_workday_extract_customer                 w
         ON TRIM(w.customer_id)                     = a.wd_customer_id
 WHERE w.customer_id                                IS NULL
;
