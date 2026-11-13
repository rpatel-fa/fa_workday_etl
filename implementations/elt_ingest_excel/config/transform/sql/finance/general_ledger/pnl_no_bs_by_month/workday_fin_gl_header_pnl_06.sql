DROP TABLE IF EXISTS workday_fin_gl_header_pnl_06
;
CREATE TABLE workday_fin_gl_header_pnl_06
    AS
SELECT
       t.*
  FROM workday_fin_gl_header_pnl_base      t
 WHERE t.column_t                      LIKE '2026-01%'
 ORDER BY 
       t.column_b
;
