-- ============================================================================
-- VALIDATION RESULTS TABLE
-- ============================================================================
DROP TABLE IF EXISTS validation_contingent_worker_result
;
CREATE TABLE validation_contingent_worker_result (
    validation_type VARCHAR
  , table_name      VARCHAR
  , data_count      BIGINT
)
;

-- ============================================================================
-- INGESTION SOURCE COUNTS (Raw tables from Okta extract)
-- ============================================================================
INSERT INTO validation_contingent_worker_result
SELECT 'ingestion_raw', 'hcm_okta_cw', COUNT(*) FROM hcm_okta_cw
;

INSERT INTO validation_contingent_worker_result
SELECT 'ingestion_raw', 'hcm_okta_cw_line_manager_update', COUNT(*) FROM hcm_okta_cw_line_manager_update
;

-- ============================================================================
-- STAGING TABLE COUNTS
-- ============================================================================
INSERT INTO validation_contingent_worker_result
SELECT 'staging', 'src_hcm_contingent_worker_base', COUNT(*) FROM src_hcm_contingent_worker_base
;

INSERT INTO validation_contingent_worker_result
SELECT 'staging', 'src_hcm_contingent_worker_nrm', COUNT(*) FROM src_hcm_contingent_worker_nrm
;

INSERT INTO validation_contingent_worker_result
SELECT 'staging', 'src_hcm_contingent_worker_match', COUNT(*) FROM src_hcm_contingent_worker_match
;

INSERT INTO validation_contingent_worker_result
SELECT 'staging', 'src_hcm_contingent_worker', COUNT(*) FROM src_hcm_contingent_worker
;

INSERT INTO validation_contingent_worker_result
SELECT 'staging', 'src_hcm_contingent_worker_mgr', COUNT(*) FROM src_hcm_contingent_worker_mgr
;

-- ============================================================================
-- WORKDAY OUTPUT TABLE COUNTS
-- ============================================================================
INSERT INTO validation_contingent_worker_result
SELECT 'record_count', 'workday_contingent_worker_name', COUNT(*) FROM workday_contingent_worker_name
;
INSERT INTO validation_contingent_worker_result
SELECT 'record_count', 'workday_contingent_worker_email', COUNT(*) FROM workday_contingent_worker_email
;
INSERT INTO validation_contingent_worker_result
SELECT 'record_count', 'workday_contingent_worker_manager', COUNT(*) FROM workday_contingent_worker_manager
;
INSERT INTO validation_contingent_worker_result
SELECT 'record_count', 'workday_contingent_worker_professional_ee', COUNT(*) FROM workday_contingent_worker_professional_ee
;

-- ============================================================================
-- DATA QUALITY CHECKS
-- ============================================================================

-- Missing required worker name fields
INSERT INTO validation_contingent_worker_result
SELECT 'data_quality', 'missing_required_worker_name', COUNT(*)
  FROM workday_contingent_worker_name
 WHERE worker_id IS NULL OR first_name IS NULL OR last_name IS NULL
;

-- Missing required email addresses (should have at least primary)
INSERT INTO validation_contingent_worker_result
SELECT 'data_quality', 'missing_primary_email', COUNT(*)
  FROM workday_contingent_worker_email
 WHERE email_address IS NULL
;

-- Invalid email format
INSERT INTO validation_contingent_worker_result
SELECT 'data_quality', 'invalid_email_addresses', COUNT(*)
  FROM src_hcm_contingent_worker
 WHERE nrm_primary_email IS NOT NULL
   AND nrm_primary_email !~ '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'
;

-- Workers without manager assignment
INSERT INTO validation_contingent_worker_result
SELECT 'data_quality', 'workers_without_manager', COUNT(*)
  FROM workday_contingent_worker_manager
 WHERE worker_id_of_direct_manager IS NULL
;

-- Duplicate worker IDs (should be 0)
INSERT INTO validation_contingent_worker_result
SELECT 'data_quality', 'duplicate_worker_ids', COUNT(*)
  FROM (
         SELECT worker_id, COUNT(*) AS cnt
           FROM workday_contingent_worker_name
          GROUP BY worker_id
         HAVING COUNT(*) > 1
       )
;

-- Duplicate emails marked as primary (should be 0)
INSERT INTO validation_contingent_worker_result
SELECT 'data_quality', 'multiple_primary_emails', COUNT(*)
  FROM workday_contingent_worker_email
 WHERE is_primary = 'Y'
 GROUP BY worker_id, email_type
HAVING COUNT(*) > 1
;

-- Workers with future activation dates (should be reviewed)
INSERT INTO validation_contingent_worker_result
SELECT 'data_quality', 'future_activation_dates', COUNT(*)
  FROM workday_contingent_worker_professional_ee
 WHERE hire_date_contract_start_date > CURRENT_DATE
;

-- Workers with past contract end dates (may be inactive)
INSERT INTO validation_contingent_worker_result
SELECT 'data_quality', 'past_contract_end_dates', COUNT(*)
  FROM workday_contingent_worker_professional_ee
 WHERE employment_contract_end_date_fixed_term < CURRENT_DATE
;

-- Invalid location (not in canonical list)
INSERT INTO validation_contingent_worker_result
SELECT 'data_quality', 'invalid_locations', COUNT(*)
  FROM workday_contingent_worker_professional_ee
 WHERE primary_position NOT IN (
       'Homebased - SGP'
     , 'Homebased - Wembley'
     , 'St George''s Park'
     , 'Wembley Stadium'
     )
;

-- Missing worker type
INSERT INTO validation_contingent_worker_result
SELECT 'data_quality', 'missing_worker_type', COUNT(*)
  FROM workday_contingent_worker_name
 WHERE worker_type IS NULL OR worker_type != 'Contingent Worker'
;

-- Invalid hours per week (should be > 0 and <= 168)
INSERT INTO validation_contingent_worker_result
SELECT 'data_quality', 'invalid_weekly_hours', COUNT(*)
  FROM workday_contingent_worker_professional_ee
 WHERE scheduled_weekly_hours IS NOT NULL
   AND (scheduled_weekly_hours::DOUBLE <= 0 OR scheduled_weekly_hours::DOUBLE > 168)
;

-- ============================================================================
-- REFERENTIAL INTEGRITY CHECKS
-- ============================================================================

-- Email records without matching name records
INSERT INTO validation_contingent_worker_result
SELECT 'referential_integrity', 'orphan_email_records', COUNT(*)
  FROM workday_contingent_worker_email e
  LEFT JOIN workday_contingent_worker_name n ON e.worker_id = n.worker_id
 WHERE n.worker_id IS NULL
;

-- Manager records without matching name records
INSERT INTO validation_contingent_worker_result
SELECT 'referential_integrity', 'orphan_manager_records', COUNT(*)
  FROM workday_contingent_worker_manager m
  LEFT JOIN workday_contingent_worker_name n ON m.worker_id = n.worker_id
 WHERE n.worker_id IS NULL
;

-- Professional EE records without matching name records
INSERT INTO validation_contingent_worker_result
SELECT 'referential_integrity', 'orphan_professional_ee_records', COUNT(*)
  FROM workday_contingent_worker_professional_ee p
  LEFT JOIN workday_contingent_worker_name n ON p.worker_id = n.worker_id
 WHERE n.worker_id IS NULL
;

-- Managers that don't exist in employee ID mapping
INSERT INTO validation_contingent_worker_result
SELECT 'referential_integrity', 'managers_not_found', COUNT(*)
  FROM workday_contingent_worker_manager m
  LEFT JOIN hcm_workday_worker_all w
    ON w.worker_type = 'Employee'
   AND NULLIF(TRIM(w.employee_id), '') = NULLIF(TRIM(m.worker_id_of_direct_manager), '')
  LEFT JOIN src_hcm_contingent_worker cw
    ON m.worker_id_of_direct_manager = cw.contingent_worker_id
 WHERE w.employee_id IS NULL
   AND cw.contingent_worker_id IS NULL
   AND m.worker_id_of_direct_manager IS NOT NULL
;

-- ============================================================================
-- LOCATION DISTRIBUTION
-- ============================================================================
DROP TABLE IF EXISTS validation_contingent_worker_location_dist
;
CREATE TABLE validation_contingent_worker_location_dist AS
SELECT 
       primary_position AS location
     , COUNT(*) AS worker_count
     , COUNT(*) * 100.0 / SUM(COUNT(*)) OVER () AS percentage
  FROM workday_contingent_worker_professional_ee
 GROUP BY primary_position
 ORDER BY worker_count DESC
;

-- ============================================================================
-- WORKER TYPE DISTRIBUTION
-- ============================================================================
DROP TABLE IF EXISTS validation_contingent_worker_type_dist
;
CREATE TABLE validation_contingent_worker_type_dist AS
SELECT 
       employee_contingent_worker_type AS worker_type
     , COUNT(*) AS worker_count
     , COUNT(*) * 100.0 / SUM(COUNT(*)) OVER () AS percentage
  FROM workday_contingent_worker_professional_ee
 GROUP BY employee_contingent_worker_type
 ORDER BY worker_count DESC
;

-- ============================================================================
-- DEPARTMENT DISTRIBUTION
-- ============================================================================
DROP TABLE IF EXISTS validation_contingent_worker_department_dist
;
CREATE TABLE validation_contingent_worker_department_dist AS
SELECT 
       applicant_source_name AS department
     , COUNT(*) AS worker_count
     , COUNT(*) * 100.0 / SUM(COUNT(*)) OVER () AS percentage
  FROM workday_contingent_worker_professional_ee
 GROUP BY applicant_source_name
 ORDER BY worker_count DESC
;

-- ============================================================================
-- INGESTION RECONCILIATION SUMMARY
-- This compares row counts from raw ingested data vs processed tables
-- Note: Counts may differ due to DISTINCT and deduplication logic
-- ============================================================================
DROP TABLE IF EXISTS validation_contingent_worker_reconciliation
;
CREATE TABLE validation_contingent_worker_reconciliation AS
WITH ingestion_counts AS (
    SELECT COUNT(*) AS okta_source_rows FROM hcm_okta_cw
),
base_counts AS (
    SELECT COUNT(*) AS base_rows FROM src_hcm_contingent_worker_base
),
workday_counts AS (
    SELECT COUNT(*) AS workday_name_rows FROM workday_contingent_worker_name
)
SELECT 
       i.okta_source_rows
     , b.base_rows AS final_staging_rows
     , w.workday_name_rows
     , b.base_rows - w.workday_name_rows AS staging_to_workday_diff
     , CASE
         WHEN w.workday_name_rows > 0 THEN 'OK - Data loaded'
         WHEN i.okta_source_rows = 0 THEN 'OK - No source data'
         ELSE 'ERROR - Missing data'
       END AS status
  FROM ingestion_counts i
     , base_counts b
     , workday_counts w
;

-- ============================================================================
-- DATA QUALITY SUMMARY REPORT
-- ============================================================================
DROP TABLE IF EXISTS validation_contingent_worker_summary
;
CREATE TABLE validation_contingent_worker_summary AS
WITH validation_totals AS (
    SELECT 
           validation_type
         , SUM(data_count) AS total_issues
         , COUNT(*) AS check_count
      FROM validation_contingent_worker_result
     WHERE validation_type IN ('data_quality', 'referential_integrity')
     GROUP BY validation_type
)
SELECT 
       validation_type
     , total_issues
     , check_count
     , CASE
         WHEN total_issues = 0 THEN 'PASS'
         ELSE 'FAIL - Review required'
       END AS status
  FROM validation_totals
 ORDER BY 
       CASE validation_type
         WHEN 'data_quality' THEN 1
         WHEN 'referential_integrity' THEN 2
         ELSE 4
       END
;

-- ============================================================================
-- OUTPUT VALIDATION RESULTS
-- ============================================================================
-- Display summary
SELECT * FROM validation_contingent_worker_summary;

-- Display reconciliation
SELECT * FROM validation_contingent_worker_reconciliation;

-- Display detailed validation results with issues
SELECT * FROM validation_contingent_worker_result
WHERE data_count > 0
ORDER BY 
  CASE validation_type
    WHEN 'data_quality' THEN 1
    WHEN 'referential_integrity' THEN 2
    WHEN 'staging' THEN 4
    WHEN 'record_count' THEN 5
    ELSE 6
  END
, validation_type
, table_name
;

-- Display location distribution
SELECT * FROM validation_contingent_worker_location_dist;

-- Display worker type distribution
SELECT * FROM validation_contingent_worker_type_dist;

-- Display department distribution (top 20)
SELECT * FROM validation_contingent_worker_department_dist
LIMIT 20
;
