DROP TABLE IF EXISTS workday_fin_gl_opening_balance
;
CREATE TABLE workday_fin_gl_opening_balance
    AS
  WITH cte_src
    AS (
SELECT
       t.*
     , CAST(COALESCE(NULLIF(TRIM(t.o_bal), ''), '0') AS DOUBLE)
                                                         opening_balance_amt
     , r.*
  FROM src_fin_gl                                        t
       LEFT OUTER JOIN
       src_fin_gl_ref                                    r
          ON t.row_id                                    = r.row_id
 WHERE t.business_unit                                  <> 'CONS'          
       )
SELECT
       ROW_NUMBER() OVER (
         ORDER BY 
               t.business_unit
             , t.account
                         )                               column_b    -- spreadsheet_key
     , t.nrm_company_id                                  column_f    -- company
     , 'GBP'                                             column_g    -- currency
     , t.nrm_company_id                                  column_i    -- company
     , 'Actuals'                                         column_j    -- ledger_type      
     , '2026'                                            column_k    -- fiscal_year
     , 'YE_Jul'                                          column_l    -- fiscal_schedule
     , t.account                                         column_m    -- journal_entry_memo
     , 'Conversion'                                      column_n    -- journal_source
     , '1'                                               column_s    -- row_id
     , CASE
         WHEN t.spend_category_ref_id        IS NOT NULL THEN t.spend_category_workday_ledger_code
         WHEN t.revenue_category_ref_id      IS NOT NULL THEN t.revenue_category_workday_ledger_code
         WHEN t.spend_category_grant_ref_id  IS NOT NULL THEN t.spend_category_grant_workday_ledger_code
         WHEN t.ledger_far_ref_id            IS NOT NULL THEN t.ledger_far_workday_ledger_code
         WHEN t.ledger_vat_dimension_ref_id  IS NOT NULL THEN t.ledger_vat_dimension_workday_ledger_code
         WHEN t.bank_account_ref_id          IS NOT NULL THEN t.bank_account_workday_ledger_code
         ELSE t.ledger_ref_id         
       END                                               column_v    -- ledger_account
     , 'FA_Child_Account_Set'                            column_w    -- account_set
     , CASE
         WHEN t.opening_balance_amt > 0
           THEN ROUND(t.opening_balance_amt, 2)
         ELSE NULL
       END                                               column_z    -- debit_amount
     , CASE
         WHEN t.opening_balance_amt < 0
           THEN ROUND(ABS(t.opening_balance_amt), 2)
         ELSE NULL
       END                                               column_aa   -- credit_amount
     , t.account                                         column_aj   -- memo
     , t.reporting_type                                  column_am   -- reporting_type
     , t.cost_centre_ref_id                              column_an   -- cost_centre
  FROM cte_src                                           t
 WHERE ABS(t.opening_balance_amt) > 0.00001
 ORDER BY
       column_b
;
