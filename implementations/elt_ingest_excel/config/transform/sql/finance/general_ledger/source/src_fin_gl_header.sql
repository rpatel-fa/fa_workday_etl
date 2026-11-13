DROP TABLE IF EXISTS src_fin_gl_header
;
CREATE TABLE src_fin_gl_header
    AS
  WITH cte_period
    AS (
SELECT *
  FROM (
       VALUES
         ('o_bal',     '2025-07-31', '202507')
       , ('August',    '2025-08-31', '202508')
       , ('September', '2025-09-30', '202509')
       , ('October',   '2025-10-31', '202510')
       , ('November',  '2025-11-30', '202511')
       , ('December',  '2025-12-31', '202512')
       , ('January',   '2026-01-31', '202601')
       , ('February',  '2026-02-28', '202602')
       , ('March',     '2026-03-31', '202603')
       , ('April',     '2026-04-30', '202604')
       , ('May',       '2026-05-31', '202605')
       , ('June',      '2026-06-30', '202606')
       , ('July',      '2026-07-31', '202607')
            ) AS t(month_name, accounting_date, period_label)
       )
     , cte_business_unit
    AS (
SELECT DISTINCT
       t.business_unit
     , t.nrm_company_id
  FROM src_fin_gl                                        t
 WHERE t.business_unit                                  <> 'CONS'
       )
SELECT
       ROW_NUMBER() OVER (
         PARTITION BY
               p.period_label
         ORDER BY
               bu.business_unit
                         )                               header_key
     , bu.business_unit
     , bu.nrm_company_id                                 company
     , 'GBP'                                             currency
     , 'Actuals'                                         ledger_type
     , p.accounting_date                                 accounting_date
     , 'Conversion'                                      journal_source
     , bu.business_unit || '-TB-' || p.period_label      accounting_journal_id
     , bu.business_unit || '_GL_TRIAL_BALANCE_' || p.period_label
                                                         external_reference_id
     , p.month_name
     , p.period_label
  FROM cte_business_unit                                 bu
       CROSS JOIN
       cte_period                                        p
 ORDER BY
       business_unit
     , period_label
;
