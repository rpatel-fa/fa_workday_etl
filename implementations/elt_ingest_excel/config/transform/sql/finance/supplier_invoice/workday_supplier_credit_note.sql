DROP TABLE IF EXISTS workday_supplier_credit_note
;
CREATE TABLE workday_supplier_credit_note
    AS
SELECT 
       ROW_NUMBER() OVER (
         ORDER BY 
               t.business_unit
             , t.wd_supplier_id
             , CAST(TRY_STRPTIME(NULLIF(TRIM(t.nrm_document_date), ''), '%d/%m/%Y') AS DATE)
             , t.document_number
                         )                         column_b    -- spreadsheet_key
     , 'Y'                                         column_c    -- add_only
     , 'Y'                                         column_e    -- auto_complete
     , 'Y'                                         column_i    -- submit
     , t.voucher_number                            column_k    -- supplier_invoice_number
     , t.nrm_company_id                            column_p    -- company
     , t.currency_id                               column_r    -- currency
     , t.wd_supplier_id                            column_s    -- supplier
     , CAST(NULL AS VARCHAR)                       column_x    -- adjustment_reason
     , t.nrm_document_date                         column_y    -- adjustment_date
     , t.nrm_transaction_amt                       column_ap   -- control_total_amount
     , '1'                                         column_cv   -- invoice_line_row_id
     , UPPER(TRIM(t.document_type))                column_da   -- invoice_line_item_description
     , CAST(NULL AS VARCHAR)                       column_dh   -- invoice_line_spend_category
     , t.nrm_transaction_amt                       column_eq   -- invoice_line_extended_amount
     , CAST(NULL AS VARCHAR)                       column_fb   -- memo
  FROM src_fin_supplier_invoice                    t
 WHERE UPPER(TRIM(t.document_type))               != 'INVOICE'
 ORDER BY 
       column_b
;
