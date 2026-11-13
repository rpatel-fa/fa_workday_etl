DROP TABLE IF EXISTS workday_fin_gl_closing_balance_pnl_line
;
CREATE TABLE workday_fin_gl_closing_balance_pnl_line
    AS
  WITH cte_header_key
    AS (
-- Year-end header per company, shared with workday_fin_gl_closing_balance_pnl_header
-- so line.column_b and header.column_b are one source of truth, not two.
SELECT
       t.business_unit
     , t.header_key
  FROM src_fin_gl_header                                 t
 WHERE t.period_label = (SELECT MAX(period_label) FROM src_fin_gl_header)
       )
     , cte_sales_item
    AS (
-- Sales Item (EIB column U on the P&L template) per company + GP account.
-- Resolved in src_fin_gl_ref but not carried on workday_fin_gl_line_base (the
-- monthly journals select line_base.* and their template has Bank Account in U),
-- so it is joined in here. (business_unit, account) is unique in src_fin_gl.
SELECT
       g.business_unit
     , g.account
     , r.sales_item_ref_id
  FROM src_fin_gl                                        g
       INNER JOIN
       src_fin_gl_ref                                    r
          ON r.row_id                                    = g.row_id
 WHERE r.sales_item_ref_id                               IS NOT NULL
       )
     , cte_agg
    AS (
-- Roll the resolved monthly-movement lines up across all 13 periods (o_bal is
-- one of the unpivoted periods in line_base, so opening is already included),
-- netting each account's full-year position at line_base's resolution grain.
-- P&L only: GP codes NOT beginning with 9 (9xxx GP codes are balance sheet and
-- are carried by workday_fin_gl_closing_balance_line).
SELECT
       SPLIT_PART(l.column_q, '_', 1)                    business_unit
     , l.column_p                                        account
     , l.column_f                                        ledger_account
     , l.column_g                                        account_set
     , l.column_l                                        currency
     , l.column_s                                        spend_category
     , l.column_t                                        revenue_category
     , l.column_v                                        reporting_type
     , l.column_w                                        cost_centre
     , l.column_x                                        event_worktag
     , l.column_y                                        camp
     , l.column_z                                        fa_learning_course
     , l.column_aa                                       commercial_partner
     , l.column_ab                                       grant_worktag
     , l.column_ac                                       programme
     , l.column_ad                                       project
     , SUM(COALESCE(l.column_j, 0) - COALESCE(l.column_k, 0))  net_amt
  FROM workday_fin_gl_line_base                          l
 WHERE LEFT(TRIM(l.column_p), 1)                         <> '9'
 GROUP BY ALL
HAVING ABS(SUM(COALESCE(l.column_j, 0) - COALESCE(l.column_k, 0))) > 0.00001
       )
     , cte_plug
    AS (
-- Balancing line per company: a P&L-only journal nets to the year's result, so
-- post -(sum of P&L net) to Retained Earnings (3300). This is the mirror of the
-- 3300 plug in workday_fin_gl_closing_balance_line (same amounts, opposite
-- side), so loading both journals nets 3300 to zero. Same equity house
-- convention: cost centre CC16005 / reporting type 'FA Operations'.
SELECT
       t.business_unit
     , -SUM(t.net_amt)                                   net_amt
  FROM cte_agg                                           t
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
     , si.sales_item_ref_id                              sales_item
     , t.reporting_type
     , t.cost_centre
     , t.event_worktag
     , t.camp
     , t.fa_learning_course
     , t.commercial_partner
     , t.grant_worktag
     , t.programme
     , t.project
  FROM cte_agg                                           t
       LEFT OUTER JOIN
       cte_sales_item                                    si
          ON si.business_unit                            = t.business_unit
         AND si.account                                  = t.account
 UNION ALL
SELECT
       p.business_unit
     , '3300'                                            ledger_account
     , 'FA_Child_Account_Set'                            account_set
     , 'GBP'                                             currency
     , p.net_amt
     , 'Current Year Earnings - Conversion Close (P&L)'  memo
     , p.business_unit || '_CLOSE_202607_PNL_CYE'        ext_ref
     , NULL                                              spend_category
     , NULL                                              revenue_category
     , NULL                                              sales_item
     , 'FA Operations'                                   reporting_type
     , 'CC16005'                                         cost_centre
     , NULL                                              event_worktag
     , NULL                                              camp
     , NULL                                              fa_learning_course
     , NULL                                              commercial_partner
     , NULL                                              grant_worktag
     , NULL                                              programme
     , NULL                                              project
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
     , x.sales_item                                      column_u    -- sales_item (P&L template; was bank_account)
     , x.reporting_type                                  column_v    -- reporting_type
     , x.cost_centre                                     column_w    -- cost_centre
     , x.event_worktag                                   column_x    -- event
     , x.camp                                            column_y    -- camp
     , x.fa_learning_course                              column_z    -- fa_learning_course
     , x.commercial_partner                              column_aa   -- commercial_partner
     , x.grant_worktag                                   column_ab   -- grant
     , x.programme                                       column_ac   -- programme
     , x.project                                         column_ad   -- project
     , NULL                                              column_ae   -- intercompany_affiliate (suppressed per Workday decision)
  FROM cte_lines                                         x
       INNER JOIN
       cte_header_key                                    h
          ON x.business_unit                             = h.business_unit
 ORDER BY
       column_b
     , column_c
;
