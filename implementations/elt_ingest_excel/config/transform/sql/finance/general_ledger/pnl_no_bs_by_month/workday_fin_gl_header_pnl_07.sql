DROP TABLE IF EXISTS workday_fin_gl_header_pnl_07
;
CREATE TABLE workday_fin_gl_header_pnl_07
    AS
SELECT
       t.*
  FROM workday_fin_gl_header_pnl_base      t
 WHERE t.column_t                      LIKE '2026-02%'
 ORDER BY 
       t.column_b
;
