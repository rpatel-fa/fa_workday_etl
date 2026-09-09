DROP TABLE IF EXISTS workday_fin_gl_header_07
;
CREATE TABLE workday_fin_gl_header_07
    AS
SELECT
       t.*
  FROM workday_fin_gl_header_base      t
 WHERE t.column_t                      LIKE '2026-02%'
 ORDER BY 
       t.column_b
;
