DROP TABLE IF EXISTS workday_fin_gl_header_base
;
CREATE TABLE workday_fin_gl_header_base
    AS
SELECT
       t.header_key                                      column_b    -- header_key
     , 'Y'                                               column_f    -- auto_complete
     , t.accounting_journal_id                           column_i    -- accounting_journal_id
     , 'Y'                                               column_j    -- submit
     , t.company                                         column_p    -- company
     , t.currency                                        column_q    -- currency
     , t.ledger_type                                     column_r    -- ledger_type
     , t.accounting_date                                 column_t    -- accounting_date
     , t.journal_source                                  column_u    -- journal_source
     , t.external_reference_id                           column_aa   -- external_reference_id
  FROM src_fin_gl_header                                 t
;
