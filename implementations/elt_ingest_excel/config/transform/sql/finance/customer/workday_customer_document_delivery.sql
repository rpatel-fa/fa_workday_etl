DROP TABLE IF EXISTS workday_customer_document_delivery
;
CREATE TABLE workday_customer_document_delivery
    AS
SELECT
       c.customer_id                                  customer_id
     , c.nrm_customer_name                            customer_name
     , CAST(NULL AS VARCHAR)                          always_separate_payments
     , CAST(NULL AS VARCHAR)                          remit_from_customer
     , CASE
         WHEN NULLIF(TRIM(c.email_to_address), '') IS NOT NULL
         THEN 'Email'
         ELSE 'Print'
       END                                            invoice_delivery_method
     , 'Bill-To Contact Emails on Invoice'            invoice_notification_email_recipients
     , CASE
         WHEN NULLIF(TRIM(c.email_to_address), '') IS NOT NULL
         THEN 'Email'
         ELSE 'Print'
       END                                            statement_delivery_method
     , 'Statement Contact Emails'                     statement_notification_email_recipients
     , CAST(NULL AS VARCHAR)                          dunning_delivery_method
     , CAST(NULL AS VARCHAR)                          dunning_letter_notification_email_recipients
     , CAST(NULL AS VARCHAR)                          electronic_invoicing_start_date
     , CAST(NULL AS VARCHAR)                          electronic_invoicing_government_issued_customer_id
     , CAST(NULL AS VARCHAR)                          electronic_invoicing_intermediary_vendor_id
     , CAST(NULL AS VARCHAR)                          electronic_invoicing_vendor_issued_customer_id
     , CAST(NULL AS VARCHAR)                          mandate_required
     , CAST(NULL AS VARCHAR)                          direct_debit_payment_type
     , CAST(NULL AS VARCHAR)                          default_mandate
  FROM src_fin_customer                c
 WHERE (
       c.nrm_agg_email_to_address  IS NOT NULL
    OR c.nrm_agg_email_cc_address  IS NOT NULL
    OR c.nrm_agg_email_bcc_address IS NOT NULL
       )
 ORDER BY 
       c.customer_id
;
