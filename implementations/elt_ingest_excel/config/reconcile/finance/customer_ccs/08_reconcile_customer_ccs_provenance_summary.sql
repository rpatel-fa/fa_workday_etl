DROP TABLE IF EXISTS reconcile_customer_ccs_provenance_case_summary
;
CREATE TABLE reconcile_customer_ccs_provenance_case_summary
    AS
  WITH cte_semantic_equivalent
    AS (
SELECT 'semantic_wording_variation'                                               provenance_bucket
     , 'systematic'                                                                issue_scope
     , 'transaction_tax_status'                                                    attribute_name
     , t.expected_value
     , t.actual_value
     , COUNT(*)                                                                    source_customer_number_count
     , COUNT(DISTINCT t.expected_customer_id)                                      merged_customer_count
     , COUNT(DISTINCT d.customer_id)                                               matched_customer_id_count
     , 'Source and transform agree; loaded Workday wording differs semantically'
                                                                                    provenance_explanation
     , 'Treat as reporting equivalence unless a counterexample appears'
                                                                                    suggested_operator_action
  FROM reconcile_customer_ccs_tax_analysis                                         t
       LEFT OUTER JOIN
       reconcile_customer_ccs_transform_extract_tax_diff_detail                    d
          ON d.customer_id                                                         = t.expected_customer_id
         AND d.comparison_scope                                                    = 'MATCHED_CUSTOMER_ID'
         AND d.expected_transaction_tax_status                                     = t.expected_value
         AND d.actual_transaction_tax_status                                       = t.actual_value
 WHERE t.attribute_name                                                            = 'transaction_tax_status'
   AND t.comparison_class                                                          = 'SEMANTIC_EQUIVALENT'
 GROUP BY
       t.expected_value
     , t.actual_value
       )
     , cte_company_number_extract_omission
    AS (
SELECT 'extract_omits_company_number_payload'                                      provenance_bucket
     , 'systematic'                                                                issue_scope
     , 'company_number_tax_payload'                                                attribute_name
     , 'United Kingdom / Company Number payload populated in transform'            expected_value
     , 'Blank tax-id fields and flags in loaded extract'                           actual_value
     , COUNT(DISTINCT t.customer_number)                                           source_customer_number_count
     , COUNT(DISTINCT t.expected_customer_id)                                      merged_customer_count
     , COUNT(DISTINCT d.customer_id)                                               matched_customer_id_count
     , 'Source and transform agree on Company Number payload; extract does not return tax-id fields or flags'
                                                                                    provenance_explanation
     , 'Treat as extract limitation, not a transform defect'
                                                                                    suggested_operator_action
  FROM reconcile_customer_ccs_tax_analysis                                         t
       LEFT OUTER JOIN
       reconcile_customer_ccs_transform_extract_tax_diff_detail                    d
          ON d.customer_id                                                         = t.expected_customer_id
         AND d.comparison_scope                                                    = 'MATCHED_CUSTOMER_ID'
         AND d.expected_tax_id_type                                                = 'Company Number'
         AND d.actual_tax_id_type                                                  = ''
 WHERE t.comparison_class                                                          = 'EXTRACT_OMITS_COMPANY_NUMBER'
       )
     , cte_transform_default_not_echoed
    AS (
SELECT 'extract_does_not_echo_transform_default_tax_code'                          provenance_bucket
     , 'systematic'                                                                issue_scope
     , 'tax_code'                                                                  attribute_name
     , expected_value
     , actual_value
     , CAST(NULL AS BIGINT)                                                        source_customer_number_count
     , CAST(NULL AS BIGINT)                                                        merged_customer_count
     , row_count                                                                   matched_customer_id_count
     , 'Transform sends a default tax code for every customer; loaded extract returns blank'
                                                                                    provenance_explanation
     , 'Treat as extract-side representation difference unless Workday confirms the field should round-trip'
                                                                                    suggested_operator_action
  FROM reconcile_customer_ccs_transform_extract_tax_diff_pattern_summary
 WHERE attribute_name                                                              = 'tax_code'
       )
     , cte_review_pack
    AS (
SELECT CASE
         WHEN o.suggested_bucket = 'accepted_workday_normalisation'
         THEN 'accepted_workday_normalisation'
         WHEN o.suggested_bucket = 'workday_target_state_review'
         THEN 'workday_target_state_review'
         WHEN o.suggested_bucket = 'ireland_sz_rule_gap'
         THEN 'ireland_sz_rule_gap_review'
         ELSE 'manual_review'
       END                                                                         provenance_bucket
     , CASE
         WHEN o.suggested_bucket = 'accepted_workday_normalisation'
         THEN 'accepted'
         ELSE 'operator_review'
       END                                                                         issue_scope
     , o.attribute_name
     , o.expected_value
     , o.actual_value
     , SUM(o.source_customer_number_count)                                         source_customer_number_count
     , COUNT(DISTINCT o.expected_customer_id)                                      merged_customer_count
     , COUNT(DISTINCT d.customer_id)                                               matched_customer_id_count
     , MAX(o.suggested_interpretation)                                             provenance_explanation
     , CASE o.suggested_bucket
         WHEN 'accepted_workday_normalisation'
         THEN 'Record as accepted Workday handling; no transform change required'
         WHEN 'workday_target_state_review'
         THEN 'Confirm whether UK should remain the tax country; if yes, raise as a Workday issue rather than a transform defect'
         WHEN 'ireland_sz_rule_gap'
         THEN 'Confirm whether this is a real rule gap or an accepted Workday representation'
         ELSE 'Review manually'
       END                                                                         suggested_operator_action
  FROM reconcile_customer_ccs_tax_review_operator_pack                             o
       LEFT OUTER JOIN
       reconcile_customer_ccs_transform_extract_tax_diff_detail                    d
          ON d.customer_id                                                         = o.expected_customer_id
         AND d.comparison_scope                                                    = 'MATCHED_CUSTOMER_ID'
         AND CASE
               WHEN o.attribute_name = 'tax_status_country'
               THEN d.expected_tax_status_country
               ELSE d.expected_transaction_tax_status
             END                                                                   = o.expected_value
         AND CASE
               WHEN o.attribute_name = 'tax_status_country'
               THEN d.actual_tax_status_country
               ELSE d.actual_transaction_tax_status
             END                                                                   = o.actual_value
 GROUP BY
       o.suggested_bucket
     , o.attribute_name
     , o.expected_value
     , o.actual_value
       )
SELECT ROW_NUMBER() OVER(
           ORDER BY CASE provenance_bucket
                      WHEN 'extract_omits_company_number_payload'         THEN 1
                      WHEN 'extract_does_not_echo_transform_default_tax_code'
                                                                        THEN 2
                      WHEN 'semantic_wording_variation'                  THEN 3
                      WHEN 'workday_target_state_review'                 THEN 4
                      WHEN 'ireland_sz_rule_gap_review'                 THEN 5
                      WHEN 'accepted_workday_normalisation'             THEN 6
                      ELSE 7
                    END
                  , attribute_name
                  , expected_value
                  , actual_value
       )                                                                            review_order
     , provenance_bucket
     , issue_scope
     , attribute_name
     , expected_value
     , actual_value
     , source_customer_number_count
     , merged_customer_count
     , matched_customer_id_count
     , provenance_explanation
     , suggested_operator_action
  FROM (
        SELECT *
          FROM cte_company_number_extract_omission

         UNION ALL
        SELECT *
          FROM cte_transform_default_not_echoed

         UNION ALL
        SELECT *
          FROM cte_semantic_equivalent

         UNION ALL
        SELECT *
          FROM cte_review_pack
       )                                                                          t
;

DROP TABLE IF EXISTS reconcile_customer_ccs_provenance_bucket_summary
;
CREATE TABLE reconcile_customer_ccs_provenance_bucket_summary
    AS
SELECT provenance_bucket
     , issue_scope
     , COUNT(*)                                                                    case_count
     , COALESCE(SUM(source_customer_number_count), 0)                              source_customer_number_count
     , COALESCE(SUM(merged_customer_count), 0)                                     merged_customer_count
     , COALESCE(SUM(matched_customer_id_count), 0)                                 matched_customer_id_count
  FROM reconcile_customer_ccs_provenance_case_summary
 GROUP BY
       provenance_bucket
     , issue_scope
 ORDER BY
       issue_scope
     , provenance_bucket
;
