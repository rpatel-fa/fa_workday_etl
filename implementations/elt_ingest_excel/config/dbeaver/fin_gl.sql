SELECT * FROM fin_general_ledger_cons;
SELECT * FROM fin_general_ledger_fa;
SELECT * FROM fin_general_ledger_nfc;
SELECT * FROM fin_general_ledger_wnsl;
SELECT * FROM fin_ref_wd_fdm_gl_mapping;
SELECT * FROM src_fin_gl_base ORDER BY account;
SELECT * FROM src_fin_gl;
SELECT * FROM src_fin_gl_header ORDER BY period_label, header_key;
SELECT * FROM src_fin_gl_ref;
SELECT * FROM workday_fin_gl_line_base;

SELECT business_unit, account FROM src_fin_gl               GROUP BY business_unit, account HAVING count(*)>1;
SELECT row_id                 FROM src_fin_gl_ref           GROUP BY row_id HAVING count(*)>1;
SELECT column_b, column_q     FROM workday_fin_gl_line_base GROUP BY column_b, column_q HAVING count(*)>1;

SELECT * FROM src_fin_gl_ref;

SELECT * FROM fin_ref_wd_fdm_gl_mapping WHERE core_vs_prism = 'Prism';

SELECT * FROM fin_ref_wd_fdm_gl_mapping WHERE workday_ref_id = 'CC21001';

SELECT * FROM workday_fin_gl_line_base where CAST(COALESCE(column_j, 0) AS double) < 0.000001 AND CAST(COALESCE(column_k, 0) AS double) < 0.000001

SELECT * FROM workday_fin_gl_line_base        WHERE column_f IS  NULL ORDER BY column_p;
SELECT DISTINCT column_p FROM workday_fin_gl_line_base WHERE column_w IS  NULL ORDER BY column_p;


SELECT * FROM workday_fin_gl_line_base  WHERE column_p = '9020-802130-00000000-0000' ORDER BY column_q;
SELECT * FROM src_fin_gl_ref            WHERE account =  '9020-802130-00000000-0000';
SELECT * FROM fin_ref_wd_fdm_gl_mapping WHERE old_code IN ('9020', '802130', '00000000', '0000');



SELECT column_p, column_w FROM workday_fin_gl_line_base WHERE column_w = 'CC21001';



