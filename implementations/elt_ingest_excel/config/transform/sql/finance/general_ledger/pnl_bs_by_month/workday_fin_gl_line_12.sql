DROP TABLE IF EXISTS workday_fin_gl_line_12
;
CREATE TABLE workday_fin_gl_line_12
    AS
SELECT
       t.*
  FROM workday_fin_gl_line_base        t
 WHERE t.column_q                      LIKE '%202607'
 ORDER BY
       t.column_b
     , t.column_c
;
