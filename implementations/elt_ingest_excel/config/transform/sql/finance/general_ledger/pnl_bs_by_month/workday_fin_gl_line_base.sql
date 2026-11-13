DROP TABLE IF EXISTS workday_fin_gl_line_base
;
CREATE TABLE workday_fin_gl_line_base
    AS
  WITH cte_base
    AS (
SELECT * 
  FROM src_fin_gl                                     t
 WHERE business_unit                                 <> 'CONS'
       )
     , cte_unpivot
    AS (
UNPIVOT cte_base
   ON o_bal, august, september, october, november, december, january, february, march, april, may, june, july
INTO NAME month_name VALUE period_amt_raw
       )
     , cte_src
    AS (
  WITH cte_amt
    AS (
SELECT
       u.*
     , COALESCE(NULLIF(TRIM(u.period_amt_raw), ''), '0')  period_amt_str
     , h.header_key
     , h.accounting_date
     , h.period_label
     , r.*
  FROM cte_unpivot                                       u
       INNER JOIN
       src_fin_gl_header                                 h
          ON u.business_unit                             = h.business_unit
         AND LOWER(u.month_name)                         = LOWER(h.month_name)
       LEFT OUTER JOIN
       src_fin_gl_ref                                    r
          ON r.row_id                                    = u.row_id
       )
SELECT
       t.*
     , CAST(t.period_amt_str AS DECIMAL(20,5))           period_amt_dec
  FROM cte_amt                                           t
 WHERE CAST(t.period_amt_str AS DECIMAL(20,5)) <> 0
       )
SELECT
       t.header_key                                      column_b    -- header_key
     , ROW_NUMBER() OVER (
         PARTITION BY
               t.business_unit
             , t.period_label
         ORDER BY
               t.account
                         )                               column_c    -- line_key
     , CASE
         WHEN t.account = '1193-409007-PWLF0006-0000'
           THEN '7350'
         WHEN t.account = '9020-802130-00000000-0000'
           THEN '1600'
         WHEN t.account = '9020-940600-00000000-0000'
           THEN '2474'
         WHEN t.account LIKE '1209-%'
           THEN '5200'
         ELSE COALESCE (
                t.spend_category_override_workday_ledger_code
              , t.spend_category_grant_override_workday_ledger_code
              , t.spend_category_workday_ledger_code
              , t.revenue_category_workday_ledger_code
              , t.spend_category_grant_workday_ledger_code
              , t.ledger_far_workday_ledger_code
              , t.ledger_vat_dimension_workday_ledger_code
              , t.bank_account_workday_ledger_code
              , t.ledger_ref_id
              )
       END                                               column_f    -- ledger_account
     , 'FA_Child_Account_Set'                            column_g    -- account_set
     , CASE
         WHEN t.period_amt_dec > 0
           THEN t.period_amt_dec
         ELSE NULL
       END                                               column_j    -- debit_amount
     , CASE
         WHEN t.period_amt_dec < 0
           THEN ABS(t.period_amt_dec)
         ELSE NULL
       END                                               column_k    -- credit_amount
     , 'GBP'                                             column_l    -- currency
     , t.account                                         column_p    -- memo
     , t.business_unit || '_' || t.account || '_' || t.period_label
                                                         column_q    -- external_reference_id
     , COALESCE(
              t.spend_category_override_ref_id
            , t.spend_category_grant_override_ref_id
            , t.spend_category_ref_id
            , t.spend_category_grant_ref_id
       )                                                  column_s    -- spend_category
     , t.revenue_category_ref_id                         column_t    -- revenue_category
     , NULL /* t.bank_account_ref_id_resolved */         column_u    -- bank_account (suppressed 20260907)
     , t.reporting_type                                  column_v    -- reporting_type
     , t.cost_centre_ref_id                              column_w    -- cost_centre
     , t.event_ref_id                                    column_x    -- event
     , t.camp_ref_id                                     column_y    -- camp
     , t.fa_learning_course_ref_id                       column_z    -- fa_learning_course
     , t.commercial_partner_ref_id                       column_aa   -- commercial_partner
     , t.grant_ref_id                                    column_ab   -- grant
     , t.program_ref_id                                  column_ac   -- programme
     , t.project_ref_id                                column_ad   -- project
     , NULL /* t.ledger_interco_affiliate_flag */        column_ae   -- intercompany_affiliate (suppressed per Workday decision; restore t.ledger_interco_affiliate_flag to reverse)
  FROM cte_src                                           t
;
