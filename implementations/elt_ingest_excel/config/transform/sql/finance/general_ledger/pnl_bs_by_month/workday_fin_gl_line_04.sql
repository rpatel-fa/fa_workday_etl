DROP TABLE IF EXISTS workday_fin_gl_line_04
;
CREATE TABLE workday_fin_gl_line_04
    AS
SELECT
       t.*
  FROM workday_fin_gl_line_base        t
 WHERE t.column_q                      LIKE '%202511'
 ORDER BY
       t.column_b
     , t.column_c
;
