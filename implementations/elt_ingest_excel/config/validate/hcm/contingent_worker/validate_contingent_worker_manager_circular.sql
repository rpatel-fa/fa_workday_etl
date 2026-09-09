DROP TABLE IF EXISTS validation_contingent_worker_manager_circular
;
CREATE TABLE validation_contingent_worker_manager_circular
    AS
WITH RECURSIVE walk AS (
   SELECT
      m.contingent_worker_id                 AS start_worker_id
    , m.contingent_worker_id                 AS current_worker_id
    , m.mgr_src_cw_contingent_worker_id      AS next_worker_id
    , [m.contingent_worker_id]::VARCHAR[]    AS path
    , 0                                      AS depth
    , FALSE                                  AS is_cycle
   FROM src_hcm_contingent_worker_mgr m
   WHERE m.mgr_src_cw_contingent_worker_id IS NOT NULL
   UNION ALL
   SELECT
      w.start_worker_id
    , m.contingent_worker_id                 AS current_worker_id
    , m.mgr_src_cw_contingent_worker_id      AS next_worker_id
    , list_append(w.path, m.contingent_worker_id) AS path
    , w.depth + 1                            AS depth
    , list_contains(w.path, m.contingent_worker_id) AS is_cycle
   FROM walk w
   JOIN src_hcm_contingent_worker_mgr m
     ON w.next_worker_id = m.contingent_worker_id
   WHERE w.next_worker_id IS NOT NULL
     AND NOT w.is_cycle
     AND w.depth < 500
)
SELECT
   start_worker_id
 , current_worker_id             AS repeated_worker_id
 , depth
 , array_to_string(path, ' -> ') AS traversal_path
  FROM walk
 WHERE is_cycle
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_mgr'       table_name
     , 'CYCLE'                               rule_type
     , 'mgr_src_cw_contingent_worker_id'     col
     , start_worker_id                       contingent_worker_id
     , traversal_path                        detail
     , 'Circular manager relationship detected (within src contingent workers).' message
  FROM validation_contingent_worker_manager_circular
;
