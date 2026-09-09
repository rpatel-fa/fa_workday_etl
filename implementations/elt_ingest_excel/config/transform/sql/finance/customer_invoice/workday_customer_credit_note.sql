DROP TABLE IF EXISTS workday_customer_credit_note
;
CREATE TABLE workday_customer_credit_note
    AS
SELECT 
       ROW_NUMBER() OVER (
         ORDER BY 
               t.business_unit
             , t.wd_customer_id
             , CAST(TRY_STRPTIME(NULLIF(TRIM(t.nrm_document_date), ''), '%d/%m/%Y') AS DATE)
             , t.document_number
                         )                         column_b    -- spreadsheet_key
     , 'Y'                                         column_c    -- add_only
     , 'Y'                                         column_e    -- auto_complete     
     , t.document_number                           column_h    -- customer_invoice_number (legacy GP)
     , 'Y'                                         column_i    -- submit
     , t.nrm_company_id                            column_l    -- company
     , t.wd_customer_id                            column_aj   -- customer_id
     , t.wd_customer_id                            column_ak   -- sold_to_customer_id
     , 'Conversion'                                column_bg   -- adjustment_reason
     , 'N'                                         column_bh   -- increase_amount_due
     , t.currency_id                               column_bk   -- currency
     , t.nrm_document_date                         column_bl   -- invoice_date
     , 'Unpaid'                                    column_bz   -- payment_status
     , 'Approved'                                  column_ca   -- document_status
     , 'Standard'                                  column_cf   -- customer_invoice_type
     , 'Open AR Adjustment Conversion'             column_cq   -- memo
     , '1'                                         column_dh   -- invoice_line_row_id
     , 'AR Conversion'                             column_dz   -- revenue_category
     , 'Open AR Adjustment Conversion'             column_ea   -- line_item_description
     , t.nrm_transaction_amt                       column_fb   -- extended_amount
     , t.document_number                           column_fc   -- line_memo
  FROM src_fin_customer_invoice                    t
 WHERE UPPER(TRIM(t.document_type))            NOT LIKE '%INVOICE%'
 ORDER BY 
       column_b
;
