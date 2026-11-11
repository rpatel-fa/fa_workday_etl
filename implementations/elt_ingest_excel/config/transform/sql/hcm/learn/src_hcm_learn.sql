-- TODO(hcm/learn): PLACEHOLDER — assign the business key (e.g. learn_id) and
-- join any reference tables once the mapping is confirmed. Currently a
-- pass-through of src_hcm_learn_nrm.
DROP TABLE IF EXISTS src_hcm_learn
;
CREATE TABLE src_hcm_learn
    AS
SELECT
       t.*
  FROM src_hcm_learn_nrm               t
;
