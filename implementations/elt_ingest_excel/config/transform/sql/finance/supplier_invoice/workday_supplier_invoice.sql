DROP TABLE IF EXISTS workday_supplier_invoice
;
CREATE TABLE workday_supplier_invoice
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
     , t.nrm_company_id                            column_r    -- company
     , t.currency_id                               column_t    -- currency
     , t.wd_supplier_id                            column_u    -- supplier
     , t.nrm_document_date                         column_af   -- invoice_date
     , t.nrm_transaction_amt                       column_as   -- control_amount_total
     , t.voucher_number                            column_bg   -- supplier_invoice_number
     , CAST(NULL AS VARCHAR)                       column_bo   -- memo
     , '1'                                         column_ek   -- invoice_line_row_id
     , UPPER(TRIM(t.document_type))                column_ep   -- invoice_line_item_description
     , CAST(NULL AS VARCHAR)                       column_ew   -- invoice_line_spend_category
     , t.nrm_transaction_amt                       column_gf   -- invoice_line_extended_amount
  FROM src_fin_supplier_invoice                    t
 WHERE UPPER(TRIM(t.document_type))                = 'INVOICE'
 ORDER BY 
       column_b
;
