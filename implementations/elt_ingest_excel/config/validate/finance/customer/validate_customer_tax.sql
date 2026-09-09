-- ============================================================================
-- VALIDATE: Customer Tax
-- Source table: workday_customer_tax
-- VBA rules: Match (cross-tab), Duplicatecolumn, EnsureDataMultipleColumns,
--            Match (Yes/No)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- CROSS_TAB: Customer ID must exist in Customer Name tab
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_tax'                                           table_name
     , 'CROSS_TAB'                                                     rule_type
     , 'customer_id'                                                   col
     , t.customer_id                                                   customer_id
     , t.customer_id                                                   detail
     , 'Customer ID must be listed first on the Customer Name tab.'    message
  FROM workday_customer_tax                                            t
  LEFT OUTER JOIN
       workday_customer_name                                           n
         ON  n.customer_id                                             = t.customer_id
 WHERE n.customer_id IS NULL
;

-- ----------------------------------------------------------------------------
-- DUPLICATE: Tax ID must be unique (when populated)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_tax'                                                  table_name
     , 'DUPLICATE'                                                            rule_type
     , 'tax_id'                                                               col
     , t.customer_id                                                          customer_id
     , t.tax_id                                                               detail
     , 'The provided Tax ID must be unique.'                                  message
  FROM workday_customer_tax                                                   t
 WHERE t.tax_id IS NOT NULL
   AND t.tax_id IN
       (
         SELECT d.tax_id
           FROM workday_customer_tax                                          d
          WHERE d.tax_id IS NOT NULL
          GROUP BY d.tax_id
         HAVING COUNT(*) > 1
       )
;

-- ----------------------------------------------------------------------------
-- MATCH: Primary Tax ID must be Yes or No (when populated)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_tax'                                                table_name
     , 'MATCH'                                                              rule_type
     , 'primary_tax_id'                                                     col
     , t.customer_id                                                        customer_id
     , t.primary_tax_id                                                     detail
     , 'The Primary Tax ID column must be populated with either "Yes" or "No".' message
  FROM workday_customer_tax                                                 t
 WHERE t.primary_tax_id IS NOT NULL
   AND TRIM(t.primary_tax_id) != ''
   AND t.primary_tax_id NOT IN ('Yes', 'No')
;

-- ----------------------------------------------------------------------------
-- MATCH: Transaction Tax ID must be Yes or No (when populated)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_tax'                                                  table_name
     , 'MATCH'                                                                rule_type
     , 'transaction_tax_id'                                                   col
     , t.customer_id                                                          customer_id
     , t.transaction_tax_id                                                   detail
     , 'The Transaction Tax ID column must be populated with either "Yes" or "No".' message
  FROM workday_customer_tax                                                   t
 WHERE t.transaction_tax_id IS NOT NULL
   AND TRIM(t.transaction_tax_id) != ''
   AND t.transaction_tax_id NOT IN ('Yes', 'No')
;

-- ----------------------------------------------------------------------------
-- ENSURE_DATA: If any tax ID field populated, all required fields must be set
-- Required group: tax_id, tax_id_type, primary_tax_id, tax_id_country
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_tax'                                                            table_name
     , 'ENSURE_DATA'                                                                    rule_type
     , 'tax_id / tax_id_type / primary_tax_id / tax_id_country'                        col
     , t.customer_id                                                                    customer_id
     , COALESCE(t.tax_id, t.tax_id_type, t.primary_tax_id, t.tax_id_country)           detail
     , 'If any of Tax ID, Tax ID Type, Primary Tax ID, Tax ID Country is populated, all must be populated.' message
  FROM workday_customer_tax                                                             t
 WHERE (
           t.tax_id         IS NOT NULL
        OR t.tax_id_type    IS NOT NULL
        OR t.primary_tax_id IS NOT NULL
        OR t.tax_id_country IS NOT NULL
       )
   AND NOT
       (
           t.tax_id         IS NOT NULL AND TRIM(t.tax_id)         != ''
       AND t.tax_id_type    IS NOT NULL AND TRIM(t.tax_id_type)    != ''
       AND t.primary_tax_id IS NOT NULL AND TRIM(t.primary_tax_id) != ''
       AND t.tax_id_country IS NOT NULL AND TRIM(t.tax_id_country) != ''
       )
;

-- ----------------------------------------------------------------------------
-- MATCH: Tax ID Type must be valid for the customer's country
-- ref: ref_workday_validate_customer_tax_id_type (Tax ID Types sheet)
-- Matches on either instance (display name) or reference_id (Workday ID)
-- ----------------------------------------------------------------------------
INSERT INTO validation_customer_issues
SELECT
       'workday_customer_tax'                                                          table_name
     , 'MATCH'                                                                       rule_type
     , 'tax_id_type'                                                                 col
     , t.customer_id                                                                 customer_id
     , t.tax_id_type || ' (' || COALESCE(t.tax_id_country, 'no country') || ')'      detail
     , 'Tax ID Type "' || t.tax_id_type || '" is not a valid type for '
       || COALESCE(t.tax_id_country, 'the supplied country') || '.'                  message
  FROM workday_customer_tax                                                            t
  LEFT JOIN
       (
         SELECT LOWER(TRIM(country))      AS country
              , LOWER(TRIM(instance))     AS tax_id_value
           FROM ref_workday_validate_customer_tax_id_type
          WHERE NULLIF(TRIM(instance), '') IS NOT NULL
         UNION
         SELECT LOWER(TRIM(country))
              , LOWER(TRIM(reference_id))
           FROM ref_workday_validate_customer_tax_id_type
          WHERE NULLIF(TRIM(reference_id), '') IS NOT NULL
       )                                                                               r
    ON r.country                                                                       = LOWER(TRIM(t.tax_id_country))
   AND r.tax_id_value                                                                  = LOWER(TRIM(t.tax_id_type))
 WHERE t.tax_id_type IS NOT NULL
   AND TRIM(t.tax_id_type) != ''
   AND t.tax_id_country IS NOT NULL
   AND TRIM(t.tax_id_country) != ''
   AND r.tax_id_value IS NULL
;
