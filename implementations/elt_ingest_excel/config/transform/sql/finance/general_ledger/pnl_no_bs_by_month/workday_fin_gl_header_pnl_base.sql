DROP TABLE IF EXISTS workday_fin_gl_header_pnl_base
;
CREATE TABLE workday_fin_gl_header_pnl_base
    AS
-- P&L-only copy of ../pnl_bs_by_month/workday_fin_gl_header_base.sql. Same header per company
-- per period; journal id / external reference carry PNL so they cannot collide
-- with the P&L+BS journals. Only company/periods that have P&L lines get a
-- header (e.g. at o_bal only WNSL carries a P&L balance), so no empty journals.
SELECT
       t.header_key                                      column_b    -- header_key
     , 'Y'                                               column_f    -- auto_complete
     , t.business_unit || '-TB-PNL-' || t.period_label   column_i    -- accounting_journal_id
     , 'Y'                                               column_j    -- submit
     , t.company                                         column_p    -- company
     , t.currency                                        column_q    -- currency
     , t.ledger_type                                     column_r    -- ledger_type
     , t.accounting_date                                 column_t    -- accounting_date
     , t.journal_source                                  column_u    -- journal_source
     , t.business_unit || '_GL_TRIAL_BALANCE_PNL_' || t.period_label
                                                         column_aa   -- external_reference_id
  FROM src_fin_gl_header                                 t
 WHERE EXISTS (
       SELECT 1
         FROM workday_fin_gl_line_pnl_base               l
        WHERE l.column_b                                 = t.header_key
          AND RIGHT(l.column_q, 6)                       = t.period_label
       )
;
