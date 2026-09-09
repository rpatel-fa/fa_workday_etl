DROP TABLE IF EXISTS reconcile_customer_ccs_expected_customer
;
CREATE TABLE reconcile_customer_ccs_expected_customer
    AS
SELECT n.customer_id                                                                 expected_customer_id
     , n.customer_name                                                               expected_customer_name
     , n.reference_id                                                                expected_reference_id
     , n.customer_category                                                           expected_customer_category
     , p.payment_terms                                                               expected_payment_terms
     , p.default_payment_type                                                        expected_default_payment_type
     , c.default_currency                                                            expected_default_currency
     , t.tax_id_country                                                              expected_tax_id_country
     , t.tax_id_type                                                                 expected_tax_id_type
     , t.tax_id                                                                      expected_tax_id
     , t.tax_status_country                                                          expected_tax_status_country
     , t.transaction_tax_status                                                      expected_transaction_tax_status
  FROM workday_customer_name                                                         n
       LEFT OUTER JOIN
       workday_customer_payment                                                      p
          ON p.customer_id                                                           = n.customer_id
       LEFT OUTER JOIN
       workday_customer_currencies                                                   c
          ON c.customer_id                                                           = n.customer_id
       LEFT OUTER JOIN
       workday_customer_tax                                                          t
          ON t.customer_id                                                           = n.customer_id
;

DROP TABLE IF EXISTS reconcile_customer_ccs_missing_in_workday
;
CREATE TABLE reconcile_customer_ccs_missing_in_workday
    AS
SELECT s.customer_number
     , s.source_row_count
     , s.source_business_units
     , s.source_customer_names
     , s.source_countries
     , m.expected_customer_id
     , m.expected_reference_id
  FROM reconcile_customer_ccs_source_profile                                         s
       LEFT OUTER JOIN
       (
         SELECT DISTINCT
                customer_number
              , expected_customer_id
              , expected_reference_id
           FROM reconcile_customer_ccs_source_merge_bridge
       )                                                                             m
          ON m.customer_number                                                       = s.customer_number
       LEFT OUTER JOIN
       (
         SELECT DISTINCT customer_number
           FROM reconcile_customer_ccs_workday_reference_bridge
       )                                                                             w
          ON w.customer_number                                                       = s.customer_number
 WHERE w.customer_number                                                             IS NULL
;

DROP TABLE IF EXISTS reconcile_customer_ccs_extra_in_workday
;
CREATE TABLE reconcile_customer_ccs_extra_in_workday
    AS
SELECT w.customer_number
     , w.workday_customer_id
     , w.workday_customer_name
     , w.workday_reference_id
  FROM reconcile_customer_ccs_workday_reference_bridge                               w
       LEFT OUTER JOIN
       reconcile_customer_ccs_source_profile                                         s
          ON s.customer_number                                                       = w.customer_number
 WHERE s.customer_number                                                             IS NULL
;

DROP TABLE IF EXISTS reconcile_customer_ccs_workday_dummy_row
;
CREATE TABLE reconcile_customer_ccs_workday_dummy_row
    AS
SELECT w.*
  FROM src_fin_workday_extract_ccs_customer                                          w
 WHERE w.customer_id NOT IN (SELECT customer_id FROM src_fin_customer)
;

DROP TABLE IF EXISTS reconcile_customer_ccs_source_multi_workday
;
CREATE TABLE reconcile_customer_ccs_source_multi_workday
    AS
SELECT w.customer_number
     , COUNT(DISTINCT w.workday_customer_id)                                         workday_customer_count
     , STRING_AGG(DISTINCT w.workday_customer_id,    '|' ORDER BY w.workday_customer_id)
                                                                                    workday_customer_ids
     , STRING_AGG(DISTINCT w.workday_reference_id,  '|' ORDER BY w.workday_reference_id)
                                                                                    workday_reference_ids
  FROM reconcile_customer_ccs_workday_reference_bridge                               w
 GROUP BY
       w.customer_number
HAVING COUNT(DISTINCT w.workday_customer_id)                                         > 1
;

DROP TABLE IF EXISTS reconcile_customer_ccs_merged_split_workday
;
CREATE TABLE reconcile_customer_ccs_merged_split_workday
    AS
SELECT m.expected_customer_id
     , m.expected_customer_name
     , m.expected_reference_id
     , COUNT(DISTINCT m.customer_number)                                             source_customer_number_count
     , COUNT(DISTINCT w.workday_customer_id)                                         workday_customer_count
     , STRING_AGG(DISTINCT w.workday_customer_id,    '|' ORDER BY w.workday_customer_id)
                                                                                    workday_customer_ids
     , STRING_AGG(DISTINCT w.workday_reference_id,  '|' ORDER BY w.workday_reference_id)
                                                                                    workday_reference_ids
  FROM reconcile_customer_ccs_source_merge_bridge                                    m
       LEFT OUTER JOIN
       reconcile_customer_ccs_workday_reference_bridge                               w
          ON w.customer_number                                                       = m.customer_number
 GROUP BY
       m.expected_customer_id
     , m.expected_customer_name
     , m.expected_reference_id
HAVING COUNT(DISTINCT w.workday_customer_id)                                         > 1
;

DROP TABLE IF EXISTS reconcile_customer_ccs_workday_overlap_merged
;
CREATE TABLE reconcile_customer_ccs_workday_overlap_merged
    AS
SELECT w.workday_customer_id
     , COUNT(DISTINCT m.expected_customer_id)                                        expected_customer_count
     , STRING_AGG(DISTINCT m.expected_customer_id,   '|' ORDER BY m.expected_customer_id)
                                                                                    expected_customer_ids
     , STRING_AGG(DISTINCT m.expected_reference_id, '|' ORDER BY m.expected_reference_id)
                                                                                    expected_reference_ids
  FROM reconcile_customer_ccs_workday_reference_bridge                               w
       LEFT OUTER JOIN
       reconcile_customer_ccs_source_merge_bridge                                    m
          ON m.customer_number                                                       = w.customer_number
 WHERE m.expected_customer_id                                                        IS NOT NULL
 GROUP BY
       w.workday_customer_id
HAVING COUNT(DISTINCT m.expected_customer_id)                                        > 1
;

INSERT INTO reconciliation_customer_ccs_issues
SELECT '2_population_exceptions'                                                     section
     , 'MISSING_IN_WORKDAY'                                                          issue_type
     , t.customer_number
     , t.expected_customer_id
     , CAST(NULL AS VARCHAR)                                                         workday_customer_id
     , t.expected_reference_id
     , CAST(NULL AS VARCHAR)                                                         workday_reference_id
     , CAST(NULL AS VARCHAR)                                                         attribute_name
     , CAST(NULL AS VARCHAR)                                                         expected_value
     , CAST(NULL AS VARCHAR)                                                         actual_value
     , 'Source GP customer_number is not present in exploded Workday reference_id'   details
  FROM reconcile_customer_ccs_missing_in_workday                                     t
;

INSERT INTO reconciliation_customer_ccs_issues
SELECT '2_population_exceptions'
     , 'EXTRA_IN_WORKDAY'
     , t.customer_number
     , CAST(NULL AS VARCHAR)
     , t.workday_customer_id
     , CAST(NULL AS VARCHAR)
     , t.workday_reference_id
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , 'Exploded Workday reference_id is not present in GP source'
  FROM reconcile_customer_ccs_extra_in_workday                                       t
;

INSERT INTO reconciliation_customer_ccs_issues
SELECT '2_population_exceptions'
     , 'WORKDAY_DUMMY_ROW'
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , t.customer_id
     , CAST(NULL AS VARCHAR)
     , t.reference_id
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , 'Workday row exists without a matching transformed customer_id'
  FROM reconcile_customer_ccs_workday_dummy_row                                      t
;

INSERT INTO reconciliation_customer_ccs_issues
SELECT '2_population_exceptions'
     , 'SOURCE_NUMBER_LINKED_TO_MULTIPLE_WORKDAY_ROWS'
     , t.customer_number
     , CAST(NULL AS VARCHAR)
     , t.workday_customer_ids
     , CAST(NULL AS VARCHAR)
     , t.workday_reference_ids
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , 'One source customer_number expands to more than one Workday customer row'
  FROM reconcile_customer_ccs_source_multi_workday                                   t
;

INSERT INTO reconciliation_customer_ccs_issues
SELECT '2_population_exceptions'
     , 'MERGED_CUSTOMER_SPLIT_ACROSS_WORKDAY_ROWS'
     , CAST(NULL AS VARCHAR)
     , t.expected_customer_id
     , t.workday_customer_ids
     , t.expected_reference_id
     , t.workday_reference_ids
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , 'One merged customer maps to more than one Workday customer row'
  FROM reconcile_customer_ccs_merged_split_workday                                   t
;

INSERT INTO reconciliation_customer_ccs_issues
SELECT '2_population_exceptions'
     , 'WORKDAY_ROW_LINKS_MULTIPLE_MERGED_CUSTOMERS'
     , CAST(NULL AS VARCHAR)
     , t.expected_customer_ids
     , t.workday_customer_id
     , t.expected_reference_ids
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , CAST(NULL AS VARCHAR)
     , 'One Workday customer row overlaps more than one merged customer'
  FROM reconcile_customer_ccs_workday_overlap_merged                                 t
;
