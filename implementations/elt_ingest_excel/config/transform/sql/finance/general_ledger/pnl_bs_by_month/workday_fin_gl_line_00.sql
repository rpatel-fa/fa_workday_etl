DROP TABLE IF EXISTS workday_fin_gl_line_00
;
CREATE TABLE workday_fin_gl_line_00
    AS
SELECT
       t.*
  FROM workday_fin_gl_line_base        t
 WHERE t.column_q                      LIKE '%202507'
 ORDER BY
       t.column_b
     , t.column_c
;
