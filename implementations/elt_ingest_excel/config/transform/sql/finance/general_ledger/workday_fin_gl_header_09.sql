DROP TABLE IF EXISTS workday_fin_gl_header_09
;
CREATE TABLE workday_fin_gl_header_09
    AS
SELECT
       t.*
  FROM workday_fin_gl_header_base      t
 WHERE t.column_t                      LIKE '2026-04%'
 ORDER BY 
       t.column_b
;
