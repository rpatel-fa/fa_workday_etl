-- TODO(hcm/learn): PLACEHOLDER — add nrm_* normalised columns (TRIM/NULLIF/
-- case normalisation, date parsing) once the source columns are confirmed.
-- Currently a pass-through of src_hcm_learn_base.
DROP TABLE IF EXISTS src_hcm_learn_nrm
;
CREATE TABLE src_hcm_learn_nrm
    AS
  WITH cte_hcm_learn_nrm
    AS (
SELECT
       t.*
  FROM src_hcm_learn_base              t
       )
SELECT t.*
  FROM cte_hcm_learn_nrm               t
;
