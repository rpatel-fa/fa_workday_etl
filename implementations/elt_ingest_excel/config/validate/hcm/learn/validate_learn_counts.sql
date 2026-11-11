DROP TABLE IF EXISTS validation_learn_counts
;
CREATE TABLE validation_learn_counts
    AS
SELECT 'hcm_learn'                                     AS table_name, COUNT(*)::BIGINT AS row_count FROM hcm_learn
UNION ALL
SELECT 'src_hcm_learn_base'                            AS table_name, COUNT(*)::BIGINT AS row_count FROM src_hcm_learn_base
UNION ALL
SELECT 'src_hcm_learn_nrm'                             AS table_name, COUNT(*)::BIGINT AS row_count FROM src_hcm_learn_nrm
UNION ALL
SELECT 'src_hcm_learn'                                 AS table_name, COUNT(*)::BIGINT AS row_count FROM src_hcm_learn
UNION ALL
SELECT 'workday_learn'                                 AS table_name, COUNT(*)::BIGINT AS row_count FROM workday_learn
;

INSERT INTO validation_learn_issues
SELECT
       c.table_name                                                     AS table_name
     , 'INFO'                                                           AS rule_type
     , 'row_count'                                                      AS col
     , CAST(NULL AS VARCHAR)                                            AS learn_id
     , CAST(c.row_count AS VARCHAR)                                     AS detail
     , 'Row count for table.'                                           AS message
  FROM validation_learn_counts c
;
