DROP TABLE IF EXISTS workday_fin_gl_line_pnl_base
;
CREATE TABLE workday_fin_gl_line_pnl_base
    AS
  WITH cte_pnl
    AS (
-- P&L-only copy of ../pnl_bs_by_month/workday_fin_gl_line_base: GP codes NOT beginning with 9
-- (9xxx GP codes are balance sheet).
SELECT
       l.*
     , SPLIT_PART(l.column_q, '_', 1)                    business_unit
     , RIGHT(l.column_q, 6)                              period_label
  FROM workday_fin_gl_line_base                          l
 WHERE LEFT(TRIM(l.column_p), 1)                         <> '9'
       )
     , cte_sales_item
    AS (
-- Sales Item replaces Bank Account in EIB column U on the P&L template.
-- (business_unit, account) is unique in src_fin_gl.
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
     , cte_plug
    AS (
-- Balancing line per company per period: a P&L-only journal nets to that
-- period's result, so post -(P&L net) to Retained Earnings (3300). Same
-- convention as the closing-balance plug: CC16005 / 'FA Operations'.
SELECT
       ANY_VALUE(t.column_b)                             column_b
     , t.business_unit
     , t.period_label
     , -SUM(COALESCE(t.column_j, 0) - COALESCE(t.column_k, 0))  net_amt
  FROM cte_pnl                                           t
 GROUP BY
       t.business_unit
     , t.period_label
HAVING ABS(SUM(COALESCE(t.column_j, 0) - COALESCE(t.column_k, 0))) > 0.00001
       )
     , cte_lines
    AS (
SELECT
       t.column_b
     , t.business_unit
     , t.period_label
     , t.column_f
     , t.column_g
     , t.column_j
     , t.column_k
     , t.column_l
     , t.column_p
     , t.column_q
     , t.column_s
     , t.column_t
     , si.sales_item_ref_id                              column_u
     , t.column_v
     , t.column_w
     , t.column_x
     , t.column_y
     , t.column_z
     , t.column_aa
     , t.column_ab
     , t.column_ac
     , t.column_ad
     , t.column_ae
  FROM cte_pnl                                           t
       LEFT OUTER JOIN
       cte_sales_item                                    si
          ON si.business_unit                            = t.business_unit
         AND si.account                                  = t.column_p
 UNION ALL
SELECT
       p.column_b
     , p.business_unit
     , p.period_label
     , '3300'                                            column_f
     , 'FA_Child_Account_Set'                            column_g
     , CASE WHEN p.net_amt > 0 THEN p.net_amt END        column_j
     , CASE WHEN p.net_amt < 0 THEN ABS(p.net_amt) END   column_k
     , 'GBP'                                             column_l
     , 'P&L Balancing - Conversion'                      column_p
     , p.business_unit || '_PNL_BALANCE_' || p.period_label
                                                         column_q
     , NULL                                              column_s
     , NULL                                              column_t
     , NULL                                              column_u
     , 'FA Operations'                                   column_v
     , 'CC16005'                                         column_w
     , NULL                                              column_x
     , NULL                                              column_y
     , NULL                                              column_z
     , NULL                                              column_aa
     , NULL                                              column_ab
     , NULL                                              column_ac
     , NULL                                              column_ad
     , NULL                                              column_ae
  FROM cte_plug                                          p
       )
SELECT
       x.column_b                                        column_b    -- header_key
     , ROW_NUMBER() OVER (
         PARTITION BY
               x.business_unit
             , x.period_label
         ORDER BY
               x.column_p
                         )                               column_c    -- line_key
     , x.column_f                                        column_f    -- ledger_account
     , x.column_g                                        column_g    -- account_set
     , x.column_j                                        column_j    -- debit_amount
     , x.column_k                                        column_k    -- credit_amount
     , x.column_l                                        column_l    -- currency
     , x.column_p                                        column_p    -- memo
     , x.column_q                                        column_q    -- external_reference_id
     , x.column_s                                        column_s    -- spend_category
     , x.column_t                                        column_t    -- revenue_category
     , x.column_u                                        column_u    -- sales_item (P&L template; was bank_account)
     , x.column_v                                        column_v    -- reporting_type
     , x.column_w                                        column_w    -- cost_centre
     , x.column_x                                        column_x    -- event
     , x.column_y                                        column_y    -- camp
     , x.column_z                                        column_z    -- fa_learning_course
     , x.column_aa                                       column_aa   -- commercial_partner
     , x.column_ab                                       column_ab   -- grant
     , x.column_ac                                       column_ac   -- programme
     , x.column_ad                                       column_ad   -- project
     , x.column_ae                                       column_ae   -- intercompany_affiliate
  FROM cte_lines                                         x
;
