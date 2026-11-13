DROP TABLE IF EXISTS workday_fin_gl_line_10
;
CREATE TABLE workday_fin_gl_line_10
    AS
SELECT
       t.*
  FROM workday_fin_gl_line_base        t
 WHERE t.column_q                      LIKE '%202605'
 ORDER BY
       t.column_b
     , t.column_c
;
