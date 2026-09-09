DROP TABLE IF EXISTS validation_contingent_worker_counts
;
CREATE TABLE validation_contingent_worker_counts
    AS
SELECT 'hcm_okta_cw'                                  AS table_name, COUNT(*)::BIGINT AS row_count FROM hcm_okta_cw
UNION ALL
SELECT 'hcm_okta_cw_line_manager_update'              AS table_name, COUNT(*)::BIGINT AS row_count FROM hcm_okta_cw_line_manager_update
UNION ALL
SELECT 'src_hcm_contingent_worker_base'               AS table_name, COUNT(*)::BIGINT AS row_count FROM src_hcm_contingent_worker_base
UNION ALL
SELECT 'src_hcm_contingent_worker_nrm'                AS table_name, COUNT(*)::BIGINT AS row_count FROM src_hcm_contingent_worker_nrm
UNION ALL
SELECT 'src_hcm_contingent_worker_match'              AS table_name, COUNT(*)::BIGINT AS row_count FROM src_hcm_contingent_worker_match
UNION ALL
SELECT 'src_hcm_contingent_worker'                    AS table_name, COUNT(*)::BIGINT AS row_count FROM src_hcm_contingent_worker
UNION ALL
SELECT 'src_hcm_contingent_worker_mgr'                AS table_name, COUNT(*)::BIGINT AS row_count FROM src_hcm_contingent_worker_mgr
UNION ALL
SELECT 'workday_contingent_worker_name'               AS table_name, COUNT(*)::BIGINT AS row_count FROM workday_contingent_worker_name
UNION ALL
SELECT 'workday_contingent_worker_professional_ee'    AS table_name, COUNT(*)::BIGINT AS row_count FROM workday_contingent_worker_professional_ee
UNION ALL
SELECT 'workday_contingent_worker_manager'            AS table_name, COUNT(*)::BIGINT AS row_count FROM workday_contingent_worker_manager
UNION ALL
SELECT 'workday_contingent_worker_email'              AS table_name, COUNT(*)::BIGINT AS row_count FROM workday_contingent_worker_email
;

INSERT INTO validation_contingent_worker_issues
SELECT
       c.table_name                                                     AS table_name
     , 'INFO'                                                           AS rule_type
     , 'row_count'                                                      AS col
     , CAST(NULL AS VARCHAR)                                            AS contingent_worker_id
     , CAST(c.row_count AS VARCHAR)                                     AS detail
     , 'Row count for table.'                                           AS message
  FROM validation_contingent_worker_counts c
;
