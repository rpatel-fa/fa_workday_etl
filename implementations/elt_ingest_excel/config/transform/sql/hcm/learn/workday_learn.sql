-- TODO(hcm/learn): PLACEHOLDER — the Workday Learning target templates are not
-- yet specified. Split this into one workday_learn_<sheet> table per 3+ HCM
-- Learning template sheet (mirroring workday_contingent_worker_*) once the
-- target workbook and column mappings are confirmed. Currently a pass-through
-- of src_hcm_learn.
DROP TABLE IF EXISTS workday_learn
;
CREATE TABLE workday_learn
    AS
SELECT
       t.*
  FROM src_hcm_learn                   t
;
