DROP TABLE IF EXISTS workday_fin_gl_header_pnl_05
;
CREATE TABLE workday_fin_gl_header_pnl_05
    AS
SELECT
       t.*
  FROM workday_fin_gl_header_pnl_base      t
 WHERE t.column_t                      LIKE '2025-12%'
 ORDER BY 
       t.column_b
;
