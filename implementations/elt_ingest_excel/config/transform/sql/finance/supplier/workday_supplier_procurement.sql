DROP TABLE IF EXISTS workday_supplier_procurement
;
-- TEMPORARY: Business have said they won't populate this for now
-- When ready, remove WHERE FALSE and uncomment the real SELECT logic
CREATE TABLE workday_supplier_procurement
    AS
  WITH cte_supplier_email_valid
    AS (
SELECT
       t.supplier_id                                                   supplier_id
     , TRIM(u.email)                                                   email_address
  FROM src_fin_supplier                t
     , UNNEST(
           STRING_SPLIT(
               LOWER(
                   COALESCE(NULLIF(TRIM(t.nrm_agg_email_to_address), ''), '')
                || ';'
                || COALESCE(NULLIF(TRIM(t.nrm_agg_email_cc_address), ''), '')
                || ';'
                || COALESCE(NULLIF(TRIM(t.nrm_agg_email_bcc_address), ''), '')
               )
             , ';'
           )
       ) u(email)
 WHERE NULLIF(TRIM(u.email), '')                                       IS NOT NULL
   AND TRIM(u.email)                                                   ~ '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'
       )
     , cte_supplier_email_unique
    AS (
SELECT t.supplier_id                                                   supplier_id
     , t.email_address                                                 email_address
     , RANK() OVER (
           PARTITION BY t.email_address
               ORDER BY t.supplier_id
       )                                                               unique_email_rnk
  FROM cte_supplier_email_valid                                        t
       )
     , cte_supplier_has_email
    AS (
SELECT e.supplier_id                                                   supplier_id
  FROM cte_supplier_email_unique                                       e
 WHERE e.unique_email_rnk                                               = 1
 GROUP BY
       e.supplier_id
       )
SELECT
       t.supplier_id                                     supplier_id
     , t.nrm_supplier_name                               supplier_name
     , 'Email'                                           purchase_order_issue_option
     , CAST(NULL AS VARCHAR)                             invoice_any_supplier
     , CAST(NULL AS VARCHAR)                             shipping_method
     , CAST(NULL AS VARCHAR)                             acknowledgement_expected
     , CAST(NULL AS VARCHAR)                             enable_asn
     , CAST(NULL AS VARCHAR)                             asn_due_in_days
     , CAST(NULL AS VARCHAR)                             supplier_minimum_order_amount
     , CAST(NULL AS VARCHAR)                             minimum_order_amount_currency
     , CAST(NULL AS VARCHAR)                             text_for_default_supplier_payment_memo
     , CAST(NULL AS VARCHAR)                             use_supplier_reference_as_default_supplier_payment_memo
     , CAST(NULL AS VARCHAR)                             use_invoice_memo_as_default_supplier_payment_memo
     , CAST(NULL AS VARCHAR)                             use_supplier_connection_memo
     , CAST(NULL AS VARCHAR)                             enable_global_location_number
     , CAST(NULL AS VARCHAR)                             disable_change_order
     , CAST(NULL AS VARCHAR)                             multi_supplier_supplier_link_for_po_issue
     , CAST(NULL AS VARCHAR)                             procurement_credit_card
     , CAST(NULL AS VARCHAR)                             edit_portal_taxes
     , CAST(NULL AS VARCHAR)                             change_order_issue_option
  FROM src_fin_supplier                t
       INNER JOIN
       cte_supplier_has_email          e
          ON e.supplier_id             = t.supplier_id
 ORDER BY
       supplier_id
;
