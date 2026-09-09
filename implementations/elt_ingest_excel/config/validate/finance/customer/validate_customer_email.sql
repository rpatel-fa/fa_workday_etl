-- ============================================================================
-- VALIDATE: Customer Email
-- Source table: workday_customer_email
-- VBA rules: Match (cross-tab), Duplicatecolumn, Missing, Match (Yes/No),
--            OneSelectionPerInstance, Email (format validation)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Customer ID must exist in Customer Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_email'                                       table_name
     , 'CROSS_TAB'                                                  rule_type
     , 'customer_id'                                                col
     , e.customer_id                                                customer_id
     , e.customer_id                                                detail
     , 'Customer ID must be listed first on the Customer Name tab.' message
  FROM workday_customer_email                                       e
  LEFT OUTER JOIN
       workday_customer_name                                        n
         ON  n.customer_id                                          = e.customer_id
 WHERE n.customer_id IS NULL
;

-- ----------------------------------------------------------------------------
-- MISSING: Email ID is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_email'                                      table_name
     , 'MISSING'                                                    rule_type
     , 'email_id'                                                   col
     , e.customer_id                                                customer_id
     , NULL                                                         detail
     , 'Email ID is a required column.'                             message
  FROM workday_customer_email                                       e
 WHERE e.email_id IS NULL
    OR TRIM(e.email_id) = ''
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: Email ID must be unique
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_email'                                           table_name
     , 'DUPLICATE'                                                       rule_type
     , 'email_id'                                                        col
     , e.customer_id                                                     customer_id
     , e.email_id                                                        detail
     , 'Email ID must be unique and cannot be duplicated.'               message
  FROM workday_customer_email                                            e
 WHERE e.email_id IN
       (
         SELECT d.email_id
           FROM workday_customer_email                                   d
          GROUP BY d.email_id
         HAVING COUNT(*) > 1
       )
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: Email Address must be unique
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_email'                                              table_name
     , 'DUPLICATE'                                                       rule_type
     , 'email_address'                                                   col
     , e.customer_id                                                     customer_id
     , e.email_address                                                   detail
     , 'Email Address must be unique.'                                   message
  FROM workday_customer_email                                            e
 WHERE e.email_address IN
       (
         SELECT d.email_address
           FROM workday_customer_email                                   d
          GROUP BY d.email_address
         HAVING COUNT(*) > 1
       )
;

-- ----------------------------------------------------------------------------
-- MISSING: Public flag is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_email'                                       table_name
     , 'MISSING'                                                         rule_type
     , 'is_public'                                                        col
     , e.customer_id                                                      customer_id
     , NULL                                                               detail
     , 'The public column is a required column.'                          message
  FROM workday_customer_email                                             e
 WHERE e.is_public IS NULL
    OR TRIM(e.is_public) = ''
;

-- ----------------------------------------------------------------------------
-- MATCH: Public flag must be Yes or No
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_email'                                                  table_name
     , 'MATCH'                                                                 rule_type
     , 'is_public'                                                              col
     , e.customer_id                                                            customer_id
     , e.is_public                                                              detail
     , 'The Email Public column must be populated with either "Yes" or "No".'   message
  FROM workday_customer_email                                                   e
 WHERE e.is_public IS NOT NULL
   AND TRIM(e.is_public) != ''
   AND e.is_public NOT IN ('Yes', 'No')
;

-- ----------------------------------------------------------------------------
-- MISSING: Primary flag is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_email'                                       table_name
     , 'MISSING'                                                         rule_type
     , 'primary_flag'                                                    col
     , e.customer_id                                                     customer_id
     , NULL                                                              detail
     , 'The primary column is a required column.'                        message
  FROM workday_customer_email                                            e
 WHERE e.primary_flag IS NULL
    OR TRIM(e.primary_flag) = ''
;

-- ----------------------------------------------------------------------------
-- MATCH: Primary flag must be Yes or No
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_email'                                                  table_name
     , 'MATCH'                                                                 rule_type
     , 'primary_flag'                                                           col
     , e.customer_id                                                            customer_id
     , e.primary_flag                                                           detail
     , 'The Primary column must be populated with either "Yes" or "No".'        message
  FROM workday_customer_email                                                   e
 WHERE e.primary_flag IS NOT NULL
   AND TRIM(e.primary_flag) != ''
   AND e.primary_flag NOT IN ('Yes', 'No')
;

-- ----------------------------------------------------------------------------
-- ONE_PRIMARY: Each Customer ID + Use For group must have exactly one primary
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_email'                                                  table_name
     , 'ONE_PRIMARY'                                                            rule_type
     , 'primary_flag'                                                           col
     , grp.customer_id                                                          customer_id
     , 'use_for=' || COALESCE(grp.use_for, 'ALL') || ' primary_count=' || grp.primary_count detail
     , 'One business email must be marked as primary and one only.'             message
  FROM
       (
         SELECT
                e.customer_id                                                   customer_id
              , e.use_for                                                       use_for
              , COUNT(*) FILTER (WHERE e.primary_flag = 'Yes')                 primary_count
           FROM workday_customer_email                                          e
          GROUP BY e.customer_id, e.use_for
         HAVING COUNT(*) FILTER (WHERE e.primary_flag = 'Yes') != 1
       )                                                                        grp
;

-- ----------------------------------------------------------------------------
-- FORMAT: Email Address must be a valid email format
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_email'                                                  table_name
     , 'FORMAT'                                                                 rule_type
     , 'email_address'                                                          col
     , e.customer_id                                                            customer_id
     , e.email_address                                                          detail
     , 'The Email Address must be valid (no spaces, must include "@", valid domain).' message
  FROM workday_customer_email                                                   e
 WHERE e.email_address IS NOT NULL
   AND TRIM(e.email_address) != ''
   AND NOT regexp_matches(
             e.email_address
           , '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'
       )
;
