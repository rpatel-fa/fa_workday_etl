DROP TABLE IF EXISTS reconcile_customer_ccs_source_profile
;
CREATE TABLE reconcile_customer_ccs_source_profile
    AS
SELECT TRIM(t.customer_number)                                                       customer_number
     , COUNT(*)                                                                       source_row_count
     , STRING_AGG(DISTINCT TRIM(t.business_unit),  '|' ORDER BY TRIM(t.business_unit))
                                                                                    source_business_units
     , STRING_AGG(DISTINCT TRIM(t.customer_name),  '|' ORDER BY TRIM(t.customer_name))
                                                                                    source_customer_names
     , STRING_AGG(DISTINCT TRIM(t.country),        '|' ORDER BY TRIM(t.country))    source_countries
  FROM src_fin_customer_base                                                         t
 WHERE NULLIF(TRIM(t.customer_number), '')                                           IS NOT NULL
 GROUP BY
       TRIM(t.customer_number)
;

DROP TABLE IF EXISTS reconcile_customer_ccs_source_merge_bridge
;
CREATE TABLE reconcile_customer_ccs_source_merge_bridge
    AS
SELECT DISTINCT
       TRIM(u.customer_number)                                                       customer_number
     , c.customer_id                                                                 expected_customer_id
     , c.nrm_customer_name                                                           expected_customer_name
     , c.nrm_agg_customer_number                                                     expected_reference_id
  FROM src_fin_customer                                                              c
     , UNNEST(STRING_SPLIT(c.nrm_agg_customer_number, '|'))                          u(customer_number)
 WHERE NULLIF(TRIM(u.customer_number), '')                                           IS NOT NULL
;

DROP TABLE IF EXISTS reconcile_customer_ccs_workday_reference_bridge
;
CREATE TABLE reconcile_customer_ccs_workday_reference_bridge
    AS
SELECT DISTINCT
       TRIM(u.customer_number)                                                       customer_number
     , w.customer_id                                                                 workday_customer_id
     , w.customer                                                                    workday_customer_name
     , w.reference_id                                                                workday_reference_id
  FROM src_fin_workday_extract_ccs_customer                                          w
     , UNNEST(STRING_SPLIT(w.reference_id, '|'))                                     u(customer_number)
 WHERE NULLIF(TRIM(u.customer_number), '')                                           IS NOT NULL
;

DROP TABLE IF EXISTS reconcile_customer_ccs_bridge
;
CREATE TABLE reconcile_customer_ccs_bridge
    AS
SELECT COALESCE(s.customer_number, m.customer_number, w.customer_number)             customer_number
     , s.source_row_count
     , s.source_business_units
     , s.source_customer_names
     , s.source_countries
     , m.expected_customer_id
     , m.expected_customer_name
     , m.expected_reference_id
     , w.workday_customer_id
     , w.workday_customer_name
     , w.workday_reference_id
     , CASE WHEN s.customer_number IS NOT NULL THEN 'Y' ELSE 'N' END                 in_source_gp
     , CASE WHEN m.customer_number IS NOT NULL THEN 'Y' ELSE 'N' END                 in_merged_customer
     , CASE WHEN w.customer_number IS NOT NULL THEN 'Y' ELSE 'N' END                 in_workday
  FROM reconcile_customer_ccs_source_profile                                         s
       FULL OUTER JOIN
       reconcile_customer_ccs_source_merge_bridge                                    m
          ON m.customer_number                                                       = s.customer_number
       FULL OUTER JOIN
       reconcile_customer_ccs_workday_reference_bridge                               w
          ON w.customer_number                                                       = COALESCE(s.customer_number, m.customer_number)
;
