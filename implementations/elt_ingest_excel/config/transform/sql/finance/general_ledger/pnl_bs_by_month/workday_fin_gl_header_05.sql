DROP TABLE IF EXISTS workday_fin_gl_header_05
;
CREATE TABLE workday_fin_gl_header_05
    AS
SELECT
       t.*
  FROM workday_fin_gl_header_base      t
 WHERE t.column_t                      LIKE '2025-12%'
 ORDER BY 
       t.column_b
;
