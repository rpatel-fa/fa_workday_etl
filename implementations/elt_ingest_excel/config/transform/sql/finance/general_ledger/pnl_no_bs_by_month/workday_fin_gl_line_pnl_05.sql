DROP TABLE IF EXISTS workday_fin_gl_line_pnl_05
;
CREATE TABLE workday_fin_gl_line_pnl_05
    AS
SELECT
       t.*
  FROM workday_fin_gl_line_pnl_base        t
 WHERE t.column_q                      LIKE '%202512'
 ORDER BY
       t.column_b
     , t.column_c
;
