DROP TABLE IF EXISTS workday_customer_invoice
;
CREATE TABLE workday_customer_invoice
    AS
SELECT 
       ROW_NUMBER() OVER (
         ORDER BY 
               t.business_unit
             , t.wd_customer_id
             , CAST(TRY_STRPTIME(NULLIF(TRIM(t.nrm_document_date), ''), '%d/%m/%Y') AS DATE)
             , t.document_number
                         )                         column_b    -- spreadsheet_key
     , 'Y'                                         column_e    -- auto_complete
     , t.document_number                           column_h    -- customer_invoice_number (legacy GP)
     , 'Y'                                         column_i    -- submit
     , t.nrm_company_id                            column_k    -- company
     , t.currency_id                               column_cp   -- currency
     , t.wd_customer_id                            column_cq   -- customer
     , t.nrm_document_date                         column_gf   -- invoice_date
     , 'Unpaid'                                    column_gy   -- payment_status
     , 'Approved'                                  column_gz   -- document_status
     , t.nrm_transaction_amt                       column_ha   -- amount_due
     , 'Standard'                                  column_hc   -- customer_invoice_type
     , '1'                                         column_ih   -- invoice_line_row_id
     , 'AR Conversion'                             column_ke   -- revenue_category
     , 'Open AR Conversion'                        column_kf   -- invoice_line_item_description
     , t.nrm_transaction_amt                       column_mj   -- extended_amount
     , 'Open AR Conversion'                        column_mt   -- memo
  FROM src_fin_customer_invoice                    t
 WHERE UPPER(TRIM(t.document_type))              LIKE '%INVOICE%'
 ORDER BY 
       column_b
;
