DROP TABLE IF EXISTS validation_learn_issues
;
CREATE TABLE validation_learn_issues (
    table_name            VARCHAR
  , rule_type             VARCHAR
  , col                   VARCHAR
  , learn_id              VARCHAR
  , detail                VARCHAR
  , message               VARCHAR
)
;
