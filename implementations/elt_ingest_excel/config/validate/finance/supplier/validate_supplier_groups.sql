-- ============================================================================
-- VALIDATE: Supplier Groups
-- Source table: workday_supplier_groups
-- VBA rules: Match (cross-tab Supplier ID), Match (group name), CheckDuplicateCombinationColumns
-- ============================================================================

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Supplier ID must exist in Supplier Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_groups'                                       table_name
     , 'CROSS_TAB'                                                    rule_type
     , 'supplier_id'                                                  col
     , g.supplier_id                                                  supplier_id
     , g.supplier_id                                                  detail
     , 'Supplier ID must be listed first on the Supplier Name tab.'   message
  FROM workday_supplier_groups                                        g
  LEFT OUTER JOIN
       workday_supplier_name                                          n
         ON  n.supplier_id                                            = g.supplier_id
 WHERE n.supplier_id IS NULL
;

-- ----------------------------------------------------------------------------
-- MATCH: Supplier Group must be a valid named-range value
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_groups'                                           table_name
     , 'MATCH'                                                            rule_type
     , 'supplier_group'                                                   col
     , g.supplier_id                                                      supplier_id
     , g.supplier_group                                                   detail
     , 'The Supplier Group name must match the drop-down values or configured additional group mappings.' message
  FROM workday_supplier_groups                                            g
 WHERE UPPER(TRIM(regexp_replace(regexp_replace(g.supplier_group, '\\s*-\\s*', '-', 'g'), '\\s+', ' ', 'g'))) NOT IN
       (
         SELECT UPPER(TRIM(regexp_replace(regexp_replace(a.supplier_group, '\\s*-\\s*', '-', 'g'), '\\s+', ' ', 'g')))
           FROM
                (
                  SELECT 'Suppliers 0 - 9'                                    supplier_group
                  UNION ALL
                  SELECT 'Suppliers A - L'                                    supplier_group
                  UNION ALL
                  SELECT 'Suppliers M - Z'                                    supplier_group
                  UNION ALL
                  SELECT DISTINCT
                         r.target_value                                     supplier_group
                    FROM ref_workday_group_additional                       r
                   WHERE r.source_type                                      = 'supplier'
                )                                                          a
       )
;

-- ----------------------------------------------------------------------------
-- DUPLICATE_COMBO: Supplier ID + Supplier Group combination must be unique
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_groups'                                             table_name
     , 'DUPLICATE_COMBO'                                                    rule_type
     , 'supplier_id + supplier_group'                                       col
     , g.supplier_id                                                        supplier_id
     , g.supplier_group                                                     detail
     , 'The Supplier Group must be unique for each Supplier ID.'            message
  FROM workday_supplier_groups                                              g
 WHERE (g.supplier_id, g.supplier_group) IN
       (
         SELECT d.supplier_id, d.supplier_group
           FROM workday_supplier_groups                                      d
          GROUP BY d.supplier_id, d.supplier_group
         HAVING COUNT(*) > 1
       )
;
