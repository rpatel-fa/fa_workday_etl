DROP TABLE IF EXISTS workday_fin_gl_header_pnl_04
;
CREATE TABLE workday_fin_gl_header_pnl_04
    AS
SELECT
       t.*
  FROM workday_fin_gl_header_pnl_base      t
 WHERE t.column_t                      LIKE '2025-11%'
 ORDER BY 
       t.column_b
;
