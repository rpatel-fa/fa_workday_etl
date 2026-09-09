DROP TABLE IF EXISTS workday_customer_invoice_adjustment
;
-- Customer Invoice Adjustment EIB (DC_Customer_Invoice_Adjustment template).
--
-- Column map mirrors workday_customer_credit_note.sql (same template), with two
-- differences driven by this source's shape:
--   column_h  = the adjustment's OWN id (template col "Customer Invoice
--               Adjustment ID"), derived as the "...: <id>" suffix, e.g.
--               CIN0112387A - NOT the base invoice number.
--   column_bh = Increase Amount Due, derived from the signed adjustment amount
--               (> 0 increases amount due -> 'Y'; otherwise 'N'). The old
--               credit-note EIB hardcoded 'N' because GP only ever produced
--               credits; this file carries genuine DEBIT (increase) rows too.
--   column_fb = Extended Amount is ALWAYS positive; direction is carried by
--               column_bh, so the source minus sign is stripped upstream.
CREATE TABLE workday_customer_invoice_adjustment
    AS
SELECT
       ROW_NUMBER() OVER (
         ORDER BY
               t.company
             , t.wd_customer_id
             , CAST(TRY_STRPTIME(NULLIF(TRIM(t.nrm_document_date), ''), '%d/%m/%Y') AS DATE)
             , t.nrm_adjustment_id
                         )                         column_b    -- spreadsheet_key
     , 'Y'                                         column_c    -- add_only
     , 'Y'                                         column_e    -- auto_complete
     , t.nrm_adjustment_id                         column_h    -- customer_invoice_adjustment_id
     , 'Y'                                         column_i    -- submit
     , t.nrm_company_id                            column_l    -- company
     , t.wd_customer_id                            column_aj   -- customer_id
     , t.wd_customer_id                            column_ak   -- sold_to_customer_id
     , 'Conversion'                                column_bg   -- adjustment_reason
     , CASE WHEN t.nrm_adjustment_amt > 0
            THEN 'Y' ELSE 'N' END                  column_bh   -- increase_amount_due
     , t.nrm_currency_id                           column_bk   -- currency
     , t.nrm_document_date                         column_bl   -- adjustment_date
     , 'Unpaid'                                    column_bz   -- payment_status
     , 'Approved'                                  column_ca   -- document_status
     , 'Standard'                                  column_cf   -- customer_invoice_type
     , 'Open AR Adjustment Conversion'             column_cq   -- memo
     , '1'                                         column_dh   -- invoice_line_row_id
     , 'AR Conversion'                             column_dz   -- revenue_category
     , 'Open AR Adjustment Conversion'             column_ea   -- line_item_description
     , t.nrm_extended_amount                       column_fb   -- extended_amount
     , t.nrm_adjustment_id                         column_fc   -- line_memo
  FROM src_fin_customer_invoice_adjustment         t
 ORDER BY
       column_b
;
