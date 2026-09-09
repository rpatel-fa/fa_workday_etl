DROP TABLE IF EXISTS workday_fin_gl_header_12
;
CREATE TABLE workday_fin_gl_header_12
    AS
SELECT
       t.*
  FROM workday_fin_gl_header_base      t
 WHERE t.column_t                      LIKE '2026-07%'
 ORDER BY
       t.column_b
;
