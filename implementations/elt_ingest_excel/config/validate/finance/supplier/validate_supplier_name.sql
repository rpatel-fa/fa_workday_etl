-- ============================================================================
-- VALIDATE: Supplier Name
-- Source table: workday_supplier_name
-- VBA rules: Duplicatecolumn, Missing, Match, EnsureDataMultipleColumns
-- ============================================================================

-- ----------------------------------------------------------------------------
-- DUPLICATE: Supplier ID must be unique
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_name'                                   table_name
     , 'DUPLICATE'                                              rule_type
     , 'supplier_id'                                            col
     , n.supplier_id                                            supplier_id
     , n.supplier_id                                            detail
     , 'Supplier ID must be unique and cannot be duplicated.'   message
  FROM workday_supplier_name                                    n
 WHERE n.supplier_id IN
       (
         SELECT d.supplier_id
           FROM workday_supplier_name                           d
          GROUP BY d.supplier_id
         HAVING COUNT(*) > 1
       )
;

-- ----------------------------------------------------------------------------
-- MISSING: Supplier ID is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_name'                                           table_name
     , 'MISSING'                                                         rule_type
     , 'supplier_id'                                                     col
     , NULL                                                              supplier_id
     , NULL                                                              detail
     , 'Supplier ID is a required column and must be populated.'         message
  FROM workday_supplier_name                                             n
 WHERE n.supplier_id IS NULL
    OR TRIM(n.supplier_id) = ''
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: Supplier Name shared across suppliers that also share a bank account
--
-- A duplicate supplier_name is NOT an error on its own: Workday identifies
-- suppliers by reference ID, not name, and the Level-2 payee split
-- (src_fin_supplier_rnk.payee_key) deliberately emits the same normalised name
-- under distinct suppliers when they are genuinely different payees (different
-- bank accounts). Flagging every duplicate name here would report those intended
-- splits as false positives.
--
-- The real invariant that survives the split: two distinct suppliers should not
-- share a name AND a bank account - that combination is a genuine duplicate the
-- split failed to merge (e.g. a shared secondary account the representative-account
-- key didn't cluster on). Only those are flagged. Supplier ID and Reference ID
-- uniqueness remain unconditionally checked in the blocks above/below.
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_name'                                                              table_name
     , 'DUPLICATE'                                                                          rule_type
     , 'supplier_name'                                                                      col
     , n.supplier_id                                                                        supplier_id
     , n.supplier_name                                                                      detail
     , 'Supplier Name duplicated across suppliers sharing a bank account - likely an unmerged duplicate.' message
  FROM workday_supplier_name                                                                n
 WHERE n.supplier_name IN
       (
         SELECT s.supplier_name
           FROM workday_supplier_settlement_account                                        s
          WHERE s.bank_account_number                                                       IS NOT NULL
          GROUP BY s.supplier_name
                 , s.bank_account_number
         HAVING COUNT(DISTINCT s.supplier_id)                                               > 1
       )
;

-- ----------------------------------------------------------------------------
-- MISSING: Supplier Name is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_name'                                           table_name
     , 'MISSING'                                                         rule_type
     , 'supplier_name'                                                   col
     , n.supplier_id                                                     supplier_id
     , NULL                                                              detail
     , 'Supplier Name is a required column and must be populated.'       message
  FROM workday_supplier_name                                             n
 WHERE n.supplier_name IS NULL
    OR TRIM(n.supplier_name) = ''
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: Reference ID must be unique
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_name'                                   table_name
     , 'DUPLICATE'                                              rule_type
     , 'reference_id'                                           col
     , n.supplier_id                                            supplier_id
     , n.reference_id                                           detail
     , 'Reference ID must be unique and cannot be duplicated.'  message
  FROM workday_supplier_name                                    n
 WHERE n.reference_id IN
       (
         SELECT d.reference_id
           FROM workday_supplier_name                           d
          GROUP BY d.reference_id
         HAVING COUNT(*) > 1
       )
;

-- ----------------------------------------------------------------------------
-- MISSING: Reference ID is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_name'                                           table_name
     , 'MISSING'                                                         rule_type
     , 'reference_id'                                                    col
     , n.supplier_id                                                     supplier_id
     , NULL                                                              detail
     , 'Reference ID is a required column and must be populated.'        message
  FROM workday_supplier_name                                             n
 WHERE n.reference_id IS NULL
    OR TRIM(n.reference_id) = ''
;

-- ----------------------------------------------------------------------------
-- MISSING: Supplier Category is required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_name'                                           table_name
     , 'MISSING'                                                         rule_type
     , 'supplier_category'                                               col
     , n.supplier_id                                                     supplier_id
     , NULL                                                              detail
     , 'Supplier Category is a required column and must be populated.'   message
  FROM workday_supplier_name                                             n
 WHERE n.supplier_category IS NULL
    OR TRIM(n.supplier_category) = ''
;

-- ----------------------------------------------------------------------------
-- MATCH: Supplier Category must be a valid value
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_name'                                            table_name
     , 'MATCH'                                                           rule_type
     , 'supplier_category'                                               col
     , n.supplier_id                                                     supplier_id
     , n.supplier_category                                               detail
     , 'Please select a supplier category from the drop-down list.'      message
  FROM workday_supplier_name                                             n
  LEFT JOIN
       (
         SELECT DISTINCT
                r.supplier_category                                      supplier_category
           FROM ref_supplier_category                                    r
       )                                                                 r
    ON UPPER(TRIM(r.supplier_category))                                  = UPPER(TRIM(n.supplier_category))
 WHERE n.supplier_category IS NOT NULL
   AND TRIM(n.supplier_category) != ''
   AND r.supplier_category IS NULL
;

-- ----------------------------------------------------------------------------
-- MATCH: Worktag Only must be Yes or No (when populated)
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_name'                                      table_name
     , 'MATCH'                                                     rule_type
     , 'worktag_only'                                              col
     , n.supplier_id                                               supplier_id
     , n.worktag_only                                              detail
     , 'Worktag Only must be either "Yes" or "No".'                message
  FROM workday_supplier_name                                       n
 WHERE n.worktag_only IS NOT NULL
   AND TRIM(n.worktag_only) != ''
   AND n.worktag_only NOT IN ('Yes', 'No')
;

-- ----------------------------------------------------------------------------
-- ENSURE_DATA: If supplier_change_source populated, supplier_source required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_name'                                              table_name
     , 'ENSURE_DATA'                                                        rule_type
     , 'supplier_source'                                                    col
     , n.supplier_id                                                        supplier_id
     , n.supplier_change_source                                             detail
     , 'If Supplier Change Source is populated, Supplier Source must also be populated.' message
  FROM workday_supplier_name                                                n
 WHERE n.supplier_change_source IS NOT NULL
   AND TRIM(n.supplier_change_source) != ''
   AND (n.supplier_source IS NULL OR TRIM(n.supplier_source) = '')
;

-- ----------------------------------------------------------------------------
-- ENSURE_DATA: If external_entity_id populated, supplier_source required
-- ----------------------------------------------------------------------------
INSERT INTO validation_supplier_issues
SELECT
       'workday_supplier_name'                                                   table_name
     , 'ENSURE_DATA'                                                             rule_type
     , 'supplier_source'                                                         col
     , n.supplier_id                                                             supplier_id
     , n.external_entity_id                                                      detail
     , 'If External Entity ID is populated, Supplier Source must also be populated.' message
  FROM workday_supplier_name                                                     n
 WHERE n.external_entity_id IS NOT NULL
   AND TRIM(n.external_entity_id) != ''
   AND (n.supplier_source IS NULL OR TRIM(n.supplier_source) = '')
;
