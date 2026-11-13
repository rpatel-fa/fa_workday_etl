DROP TABLE IF EXISTS workday_fin_gl_header_03
;
CREATE TABLE workday_fin_gl_header_03
    AS
SELECT
       t.*
  FROM workday_fin_gl_header_base      t
 WHERE t.column_t                      LIKE '2025-10%'
 ORDER BY 
       t.column_b
;
