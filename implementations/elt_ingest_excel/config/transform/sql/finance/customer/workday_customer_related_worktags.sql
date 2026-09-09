DROP TABLE IF EXISTS workday_customer_related_worktags
;
CREATE TABLE workday_customer_related_worktags
    AS
SELECT
       c.customer_id                                  customer_id
     , c.nrm_customer_name                            customer_name
     , 'Customer'                                     worktag_type
     , CAST(NULL AS VARCHAR)                          required_on_transaction
     , CAST(NULL AS VARCHAR)                          required_on_transaction_for_validation
     , CAST(NULL AS VARCHAR)                          delete_default_value
     , CAST(NULL AS VARCHAR)                          worktag_id_type
     , CAST(NULL AS VARCHAR)                          default_worktag_id_value
     , CAST(NULL AS VARCHAR)                          external_supplier_invoice_source
     , CAST(NULL AS VARCHAR)                          replace_all_allowed_values
     , CAST(NULL AS VARCHAR)                          delete_allowed_values
     , CAST(NULL AS VARCHAR)                          allowed_worktag_id_value
  FROM src_fin_customer                c
 WHERE 1 = 2
;
