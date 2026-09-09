INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_mgr'                                table_name
     , 'COVERAGE'                                                     rule_type
     , 'contingent_worker_id'                                         col
     , cw.contingent_worker_id                                        contingent_worker_id
     , CAST(NULL AS VARCHAR)                                          detail
     , 'contingent_worker_id from src_hcm_contingent_worker is missing from src_hcm_contingent_worker_mgr.' message
  FROM src_hcm_contingent_worker cw
       LEFT OUTER JOIN
       src_hcm_contingent_worker_mgr m
          ON m.contingent_worker_id = cw.contingent_worker_id
 WHERE m.contingent_worker_id IS NULL
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_mgr'                                table_name
     , 'COVERAGE'                                                     rule_type
     , 'contingent_worker_id'                                         col
     , m.contingent_worker_id                                         contingent_worker_id
     , CAST(NULL AS VARCHAR)                                          detail
     , 'contingent_worker_id exists in src_hcm_contingent_worker_mgr but not in src_hcm_contingent_worker.' message
  FROM src_hcm_contingent_worker_mgr m
       LEFT OUTER JOIN
       src_hcm_contingent_worker cw
          ON cw.contingent_worker_id = m.contingent_worker_id
 WHERE cw.contingent_worker_id IS NULL
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_mgr'                                table_name
     , 'OVERLAP'                                                      rule_type
     , 'manager_source'                                               col
     , m.contingent_worker_id                                         contingent_worker_id
     , CAST(
          (CASE WHEN m.mgr_emp_employee_id IS NOT NULL THEN 'Employee|' ELSE '' END) ||
          (CASE WHEN m.mgr_src_cw_contingent_worker_id IS NOT NULL THEN 'SrcCW|' ELSE '' END)
          AS VARCHAR
       )                                                             detail
     , 'More than one manager source matched for this contingent_worker_id.' message
  FROM src_hcm_contingent_worker_mgr m
 WHERE (CASE WHEN m.mgr_emp_employee_id IS NOT NULL THEN 1 ELSE 0 END
      + CASE WHEN m.mgr_src_cw_contingent_worker_id IS NOT NULL THEN 1 ELSE 0 END
      ) > 1
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_mgr'                                table_name
     , 'MISSING'                                                      rule_type
     , 'resolved_manager_id'                                          col
     , m.contingent_worker_id                                         contingent_worker_id
     , CAST(
          COALESCE(m.nrm_managerid, '') || '|' ||
          COALESCE(m.nrm_manager_name, '') || '|' ||
          COALESCE(m.nrm_manager_email_address, '') || '|' ||
          COALESCE(m.nrm_manager_email, '')
          AS VARCHAR
       )                                                             detail
     , 'Manager fields exist but resolved_manager_id is NULL.'        message
  FROM src_hcm_contingent_worker_mgr m
 WHERE m.resolved_manager_id IS NULL
   AND (
          m.nrm_managerid IS NOT NULL
       OR m.nrm_manager_name IS NOT NULL
       OR m.nrm_manager_email_address IS NOT NULL
       OR m.nrm_manager_email IS NOT NULL
       )
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_mgr'                                table_name
     , 'MISSING'                                                      rule_type
     , 'employee_or_src_manager'                                      col
     , m.contingent_worker_id                                         contingent_worker_id
     , CAST(
          COALESCE(m.nrm_managerid, '') || '|' ||
          COALESCE(m.nrm_manager_name, '') || '|' ||
          COALESCE(m.nrm_manager_email_address, '') || '|' ||
          COALESCE(m.nrm_manager_email, '')
          AS VARCHAR
       )                                                             detail
     , 'Manager fields exist but no match was found in hcm_workday_worker_all (Employee) or src_hcm_contingent_worker.' message
  FROM src_hcm_contingent_worker_mgr m
 WHERE (
          m.nrm_managerid IS NOT NULL
       OR m.nrm_manager_name IS NOT NULL
       OR m.nrm_manager_email_address IS NOT NULL
       OR m.nrm_manager_email IS NOT NULL
       )
   AND m.mgr_emp_employee_id IS NULL
   AND m.mgr_src_cw_contingent_worker_id IS NULL
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_mgr'                                table_name
     , 'MISSING'                                                      rule_type
     , 'manager_fields'                                               col
     , m.contingent_worker_id                                         contingent_worker_id
     , CAST(NULL AS VARCHAR)                                          detail
     , 'All manager fields are NULL/blank; no manager supplied for this contingent worker.' message
  FROM src_hcm_contingent_worker_mgr m
 WHERE m.resolved_manager_id IS NULL
   AND m.nrm_managerid IS NULL
   AND m.nrm_manager_name IS NULL
   AND m.nrm_manager_email_address IS NULL
   AND m.nrm_manager_email IS NULL
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_mgr'                                table_name
     , 'CONSISTENCY'                                                  rule_type
     , 'resolved_manager_source|resolved_manager_id'                  col
     , m.contingent_worker_id                                         contingent_worker_id
     , CAST(m.resolved_manager_source || '|' || m.resolved_manager_id AS VARCHAR) detail
     , 'resolved_manager_source is NULL but resolved_manager_id is not NULL.' message
  FROM src_hcm_contingent_worker_mgr m
 WHERE m.resolved_manager_source IS NULL
   AND m.resolved_manager_id IS NOT NULL
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_mgr'                                table_name
     , 'CONSISTENCY'                                                  rule_type
     , 'resolved_manager_source|resolved_manager_id'                  col
     , m.contingent_worker_id                                         contingent_worker_id
     , CAST(m.resolved_manager_source || '|' || COALESCE(m.resolved_manager_id, '') AS VARCHAR) detail
     , 'resolved_manager_source is Employee but mgr_emp_employee_id is NULL or does not match resolved_manager_id.' message
  FROM src_hcm_contingent_worker_mgr m
 WHERE m.resolved_manager_source = 'Employee'
   AND (
          m.mgr_emp_employee_id IS NULL
       OR m.resolved_manager_id <> m.mgr_emp_employee_id
       )
;

INSERT INTO validation_contingent_worker_issues
SELECT
       'src_hcm_contingent_worker_mgr'                                table_name
     , 'CONSISTENCY'                                                  rule_type
     , 'resolved_manager_source|resolved_manager_id'                  col
     , m.contingent_worker_id                                         contingent_worker_id
     , CAST(m.resolved_manager_source || '|' || COALESCE(m.resolved_manager_id, '') AS VARCHAR) detail
     , 'resolved_manager_source is Contingent Worker (src) but mgr_src_cw_contingent_worker_id is NULL or does not match resolved_manager_id.' message
  FROM src_hcm_contingent_worker_mgr m
 WHERE m.resolved_manager_source = 'Contingent Worker (src)'
   AND (
          m.mgr_src_cw_contingent_worker_id IS NULL
       OR m.resolved_manager_id <> m.mgr_src_cw_contingent_worker_id
       )
;
