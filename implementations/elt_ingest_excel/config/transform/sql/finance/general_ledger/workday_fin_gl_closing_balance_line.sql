DROP TABLE IF EXISTS workday_fin_gl_closing_balance_line
;
CREATE TABLE workday_fin_gl_closing_balance_line
    AS
  WITH cte_header_key
    AS (
-- Year-end header per company, shared with workday_fin_gl_closing_balance_header
-- so line.column_b and header.column_b are one source of truth, not two.
SELECT
       t.business_unit
     , t.header_key
  FROM src_fin_gl_header                                 t
 WHERE t.period_label = (SELECT MAX(period_label) FROM src_fin_gl_header)
       )
     , cte_agg
    AS (
-- Roll the resolved monthly-movement lines up across all 13 periods (o_bal is
-- one of the unpivoted periods in line_base, so opening is already included).
-- Ledger account and every worktag are constant per source account, so this
-- nets each account's full-year position at line_base's resolution grain.
-- Each row is tagged with its Ledger Account Type (from the workday ledger code);
-- cte_bs below keeps only the balance-sheet types (Asset/Liability/Equity).
SELECT
       SPLIT_PART(l.column_q, '_', 1)                    business_unit
     , l.column_p                                        account
     , l.column_f                                        ledger_account
     , l.column_g                                        account_set
     , l.column_l                                        currency
     , l.column_s                                        spend_category
     , l.column_t                                        revenue_category
     , l.column_u                                        bank_account
     , l.column_v                                        reporting_type
     , l.column_w                                        cost_centre
     , l.column_x                                        event_worktag
     , l.column_y                                        camp
     , l.column_z                                        fa_learning_course
     , l.column_aa                                       commercial_partner
     , l.column_ab                                       grant_worktag
     , l.column_ac                                       programme
     , l.column_ad                                       project
     , l.column_ae                                       intercompany_affiliate
       -- Ledger Account Type, derived from the workday ledger code:
       --   1xxx = Asset, 2xxx = Liability, 3xxx = Equity  (balance sheet)
       --   4xxx = Income, 5xxx-8xxx = Expense             (P&L)
     , CASE
         WHEN TRY_CAST(l.column_f AS INTEGER) BETWEEN 1000 AND 1999 THEN 'Asset'
         WHEN TRY_CAST(l.column_f AS INTEGER) BETWEEN 2000 AND 2999 THEN 'Liability'
         WHEN TRY_CAST(l.column_f AS INTEGER) BETWEEN 3000 AND 3999 THEN 'Equity'
         WHEN TRY_CAST(l.column_f AS INTEGER) BETWEEN 4000 AND 4999 THEN 'Income'
         ELSE 'Expense'
       END                                               ledger_account_type
     , SUM(COALESCE(l.column_j, 0) - COALESCE(l.column_k, 0))  net_amt
  FROM workday_fin_gl_line_base                          l
 GROUP BY ALL
       )
     , cte_bs
    AS (
-- Balance-sheet accounts only: keep Asset / Liability / Equity; drop Income and
-- Expense (P&L). Their net result is rolled into Retained Earnings by cte_plug.
SELECT *
  FROM cte_agg
 WHERE ledger_account_type                               IN ('Asset', 'Liability', 'Equity')
   AND ABS(net_amt)                                      > 0.00001
       )
     , cte_plug
    AS (
-- Current-year P&L result posted to Retained Earnings (3300) so each company's
-- closing balance sheet balances. Plug = -(sum of BS net), i.e. a profit
-- (BS net debit) becomes a credit to equity. Equity house convention:
-- cost centre CC16005 / reporting type 'FA Operations'.
SELECT
       t.business_unit
     , -SUM(t.net_amt)                                   net_amt
  FROM cte_bs                                            t
 GROUP BY
       t.business_unit
HAVING ABS(SUM(t.net_amt))                               > 0.00001
       )
     , cte_lines
    AS (
SELECT
       t.business_unit
     , t.ledger_account
     , t.account_set
     , t.currency
     , t.net_amt
     , t.account                                         memo
     , t.business_unit || '_' || t.account || '_CLOSE_202607'  ext_ref
     , t.spend_category
     , t.revenue_category
     , t.bank_account
     , t.reporting_type
     , t.cost_centre
     , t.event_worktag
     , t.camp
     , t.fa_learning_course
     , t.commercial_partner
     , t.grant_worktag
     , t.programme
     , t.project
     , t.intercompany_affiliate
  FROM cte_bs                                            t
 UNION ALL
SELECT
       p.business_unit
     , '3300'                                            ledger_account
     , 'FA_Child_Account_Set'                            account_set
     , 'GBP'                                             currency
     , p.net_amt
     , 'Current Year Earnings - Conversion Close'        memo
     , p.business_unit || '_CLOSE_202607_CYE'            ext_ref
     , NULL                                              spend_category
     , NULL                                              revenue_category
     , NULL                                              bank_account
     , 'FA Operations'                                   reporting_type
     , 'CC16005'                                         cost_centre
     , NULL                                              event_worktag
     , NULL                                              camp
     , NULL                                              fa_learning_course
     , NULL                                              commercial_partner
     , NULL                                              grant_worktag
     , NULL                                              programme
     , NULL                                              project
     , NULL                                              intercompany_affiliate
  FROM cte_plug                                          p
       )
SELECT
       h.header_key                                      column_b    -- header_key
     , ROW_NUMBER() OVER (
         PARTITION BY
               h.header_key
         ORDER BY
               x.ledger_account
             , x.ext_ref
                         )                               column_c    -- line_key
     , x.ledger_account                                  column_f    -- ledger_account
     , x.account_set                                     column_g    -- account_set
     , CASE
         WHEN x.net_amt > 0
           THEN ROUND(x.net_amt, 2)
         ELSE NULL
       END                                               column_j    -- debit_amount
     , CASE
         WHEN x.net_amt < 0
           THEN ROUND(ABS(x.net_amt), 2)
         ELSE NULL
       END                                               column_k    -- credit_amount
     , x.currency                                        column_l    -- currency
     , x.memo                                            column_p    -- memo
     , x.ext_ref                                         column_q    -- external_reference_id
     , x.spend_category                                  column_s    -- spend_category
     , x.revenue_category                                column_t    -- revenue_category
     , x.bank_account                                    column_u    -- bank_account
     , x.reporting_type                                  column_v    -- reporting_type
     , x.cost_centre                                     column_w    -- cost_centre
     , x.event_worktag                                   column_x    -- event
     , x.camp                                            column_y    -- camp
     , x.fa_learning_course                              column_z    -- fa_learning_course
     , x.commercial_partner                              column_aa   -- commercial_partner
     , x.grant_worktag                                   column_ab   -- grant
     , x.programme                                       column_ac   -- programme
     , x.project                                         column_ad   -- project
     , NULL /* x.intercompany_affiliate */               column_ae   -- intercompany_affiliate (suppressed per Workday decision; restore x.intercompany_affiliate to reverse)
  FROM cte_lines                                         x
       INNER JOIN
       cte_header_key                                    h
          ON x.business_unit                             = h.business_unit
 ORDER BY
       column_b
     , column_c
;
