DROP TABLE IF EXISTS workday_fin_gl_closing_balance_pnl_header
;
CREATE TABLE workday_fin_gl_closing_balance_pnl_header
    AS
  WITH cte_year_end
    AS (
-- P&L counterpart of workday_fin_gl_closing_balance_header. Same year-end
-- header per company (final GL period) from src_fin_gl_header; only the journal
-- id and external reference differ so the two closing journals cannot collide.
SELECT t.*
  FROM src_fin_gl_header                                 t
 WHERE t.period_label = (SELECT MAX(period_label) FROM src_fin_gl_header)
       )
SELECT
       t.header_key                                      column_b    -- header_key
     , 'Y'                                               column_f    -- auto_complete
     , t.business_unit || '-TB-CLOSE-PNL-' || t.period_label
                                                         column_i    -- accounting_journal_id
     , 'Y'                                               column_j    -- submit
     , t.company                                         column_p    -- company
     , t.currency                                        column_q    -- currency
     , t.ledger_type                                     column_r    -- ledger_type
     , t.accounting_date                                 column_t    -- accounting_date
     , t.journal_source                                  column_u    -- journal_source
     , t.business_unit || '_GL_CLOSING_BALANCE_PNL_' || t.period_label
                                                         column_aa   -- external_reference_id
  FROM cte_year_end                                      t
 ORDER BY
       column_b
;
