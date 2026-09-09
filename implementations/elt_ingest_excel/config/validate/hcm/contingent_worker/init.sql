DROP TABLE IF EXISTS validation_contingent_worker_issues
;
CREATE TABLE validation_contingent_worker_issues (
    table_name            VARCHAR
  , rule_type             VARCHAR
  , col                   VARCHAR
  , contingent_worker_id  VARCHAR
  , detail                VARCHAR
  , message               VARCHAR
)
;
