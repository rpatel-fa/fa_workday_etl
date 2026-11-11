-- TODO(hcm/learn): PLACEHOLDER — add row filters and dedup keys once the
-- source workbook columns are confirmed. Currently a pass-through of the raw
-- ingest table.
DROP TABLE IF EXISTS src_hcm_learn_base
;
CREATE TABLE src_hcm_learn_base
    AS
  WITH cte_hcm_learn_base
    AS (
SELECT DISTINCT
       t.*
  FROM hcm_learn                       t
 WHERE
       1 = 1
       )
SELECT t.*
  FROM cte_hcm_learn_base              t
;
