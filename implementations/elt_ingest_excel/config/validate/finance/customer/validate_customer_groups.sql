-- ============================================================================
-- VALIDATE: Customer Groups
-- Source table: workday_customer_groups
-- VBA rules: Match (cross-tab Customer ID), Match (group name), CheckDuplicateCombinationColumns
-- ============================================================================

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Customer ID must exist in Customer Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_groups'                                       table_name
     , 'CROSS_TAB'                                                    rule_type
     , 'customer_id'                                                  col
     , g.customer_id                                                  customer_id
     , g.customer_id                                                  detail
     , 'Customer ID must be listed first on the Customer Name tab.'   message
  FROM workday_customer_groups                                        g
  LEFT OUTER JOIN
       workday_customer_name                                          n
         ON  n.customer_id                                            = g.customer_id
 WHERE n.customer_id IS NULL
;

-- ----------------------------------------------------------------------------
-- MATCH: Customer Group must be a valid named-range value
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_groups'                                           table_name
     , 'MATCH'                                                            rule_type
     , 'customer_group'                                                   col
     , g.customer_id                                                      customer_id
     , g.customer_group                                                   detail
     , 'The Customer Group name must match the drop-down values or configured additional group mappings.' message
  FROM workday_customer_groups                                            g
 WHERE UPPER(TRIM(regexp_replace(regexp_replace(g.customer_group, '\\s*-\\s*', '-', 'g'), '\\s+', ' ', 'g'))) NOT IN
       (
         SELECT UPPER(TRIM(regexp_replace(regexp_replace(a.customer_group, '\\s*-\\s*', '-', 'g'), '\\s+', ' ', 'g')))
           FROM
                (
                  SELECT 'Customers 0 - 9'                                    customer_group
                  UNION ALL
                  SELECT 'Customers A - L'                                    customer_group
                  UNION ALL
                  SELECT 'Customers M - Z'                                    customer_group
                  UNION ALL
                  SELECT DISTINCT
                         r.target_value                                     customer_group
                    FROM ref_workday_group_additional                       r
                   WHERE r.source_type                                      = 'customer'
                )                                                          a
       )
;

-- ----------------------------------------------------------------------------
-- DUPLICATE_COMBO: Customer ID + Customer Group combination must be unique
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_groups'                                             table_name
     , 'DUPLICATE_COMBO'                                                    rule_type
     , 'customer_id + customer_group'                                       col
     , g.customer_id                                                        customer_id
     , g.customer_group                                                     detail
     , 'The Customer Group must be unique for each Customer ID.'            message
  FROM workday_customer_groups                                              g
 WHERE (g.customer_id, g.customer_group) IN
       (
         SELECT d.customer_id, d.customer_group
           FROM workday_customer_groups                                      d
          GROUP BY d.customer_id, d.customer_group
         HAVING COUNT(*) > 1
       )
;
