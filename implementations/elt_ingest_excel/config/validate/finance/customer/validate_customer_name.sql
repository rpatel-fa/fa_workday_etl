-- ============================================================================
-- VALIDATE: Customer Name
-- Source table: workday_customer_name
-- VBA rules: Duplicatecolumn, Missing, Match, EnsureDataMultipleColumns
-- ============================================================================

-- ----------------------------------------------------------------------------
-- DUPLICATE: Customer ID must be unique
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_name'                                   table_name
     , 'DUPLICATE'                                              rule_type
     , 'customer_id'                                            col
     , n.customer_id                                            customer_id
     , n.customer_id                                            detail
     , 'Customer ID must be unique and cannot be duplicated.'   message
  FROM workday_customer_name                                    n
 WHERE n.customer_id IN
       (
         SELECT d.customer_id
           FROM workday_customer_name                           d
          GROUP BY d.customer_id
         HAVING COUNT(*) > 1
       )
;

-- ----------------------------------------------------------------------------
-- MISSING: Customer ID is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_name'                                           table_name
     , 'MISSING'                                                         rule_type
     , 'customer_id'                                                     col
     , NULL                                                              customer_id
     , NULL                                                              detail
     , 'Customer ID is a required column and must be populated.'         message
  FROM workday_customer_name                                             n
 WHERE n.customer_id IS NULL
    OR TRIM(n.customer_id) = ''
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: Customer Name must be unique
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_name'                                   table_name
     , 'DUPLICATE'                                              rule_type
     , 'customer_name'                                          col
     , n.customer_id                                            customer_id
     , n.customer_name                                          detail
     , 'Customer Name must be unique and cannot be duplicated.' message
  FROM workday_customer_name                                    n
 WHERE n.customer_name IN
       (
         SELECT d.customer_name
           FROM workday_customer_name                           d
          GROUP BY d.customer_name
         HAVING COUNT(*) > 1
       )
;

-- ----------------------------------------------------------------------------
-- MISSING: Customer Name is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_name'                                           table_name
     , 'MISSING'                                                         rule_type
     , 'customer_name'                                                   col
     , n.customer_id                                                     customer_id
     , NULL                                                              detail
     , 'Customer Name is a required column and must be populated.'       message
  FROM workday_customer_name                                             n
 WHERE n.customer_name IS NULL
    OR TRIM(n.customer_name) = ''
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: Reference ID must be unique
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_name'                                   table_name
     , 'DUPLICATE'                                              rule_type
     , 'reference_id'                                           col
     , n.customer_id                                            customer_id
     , n.reference_id                                           detail
     , 'Reference ID must be unique and cannot be duplicated.'  message
  FROM workday_customer_name                                    n
 WHERE n.reference_id IN
       (
         SELECT d.reference_id
           FROM workday_customer_name                           d
          GROUP BY d.reference_id
         HAVING COUNT(*) > 1
       )
;

-- ----------------------------------------------------------------------------
-- MISSING: Reference ID is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_name'                                           table_name
     , 'MISSING'                                                         rule_type
     , 'reference_id'                                                    col
     , n.customer_id                                                     customer_id
     , NULL                                                              detail
     , 'Reference ID is a required column and must be populated.'        message
  FROM workday_customer_name                                             n
 WHERE n.reference_id IS NULL
    OR TRIM(n.reference_id) = ''
;

-- ----------------------------------------------------------------------------
-- MISSING: Customer Category is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_name'                                           table_name
     , 'MISSING'                                                         rule_type
     , 'customer_category'                                               col
     , n.customer_id                                                     customer_id
     , NULL                                                              detail
     , 'Customer Category is a required column and must be populated.'   message
  FROM workday_customer_name                                             n
 WHERE n.customer_category IS NULL
    OR TRIM(n.customer_category) = ''
;

-- ----------------------------------------------------------------------------
-- MATCH: Customer Category must be a valid value
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_name'                                            table_name
     , 'MATCH'                                                           rule_type
     , 'customer_category'                                               col
     , n.customer_id                                                     customer_id
     , n.customer_category                                               detail
     , 'Please select a customer category from the drop-down list.'      message
  FROM workday_customer_name                                             n
  LEFT JOIN
       (
         SELECT DISTINCT
               r.target_value                                           customer_category
           FROM ref_customer_category                                    r
       )                                                                 rc
   ON UPPER(TRIM(rc.customer_category))                                 = UPPER(TRIM(n.customer_category))
 WHERE n.customer_category IS NOT NULL
   AND TRIM(n.customer_category) != ''
   AND rc.customer_category IS NULL
;

-- ----------------------------------------------------------------------------
-- MATCH: Worktag Only must be Yes or No (when populated)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_name'                                      table_name
     , 'MATCH'                                                     rule_type
     , 'worktag_only'                                              col
     , n.customer_id                                               customer_id
     , n.worktag_only                                              detail
     , 'Worktag Only must be either "Yes" or "No".'                message
  FROM workday_customer_name                                       n
 WHERE n.worktag_only IS NOT NULL
   AND TRIM(n.worktag_only) != ''
   AND n.worktag_only NOT IN ('Yes', 'No')
;
