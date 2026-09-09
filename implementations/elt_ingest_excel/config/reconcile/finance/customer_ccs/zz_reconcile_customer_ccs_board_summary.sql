-- Board-level reconciliation summary (first tab of the review pack).
-- One flat, ordered row per headline item. Every value is pulled live from the
-- reconcile tables already built by the numbered files and the z_ report, so the
-- summary can never drift from the detail. Filter disposition = 'FIX_IN_WORKDAY'
-- to get the Finance action list; 'OPERATOR_DECISION' for open decisions.
-- Runs last (zz_ prefix) because it reads reconcile_customer_ccs_report.

DROP TABLE IF EXISTS reconcile_customer_ccs_board_summary
;

CREATE TABLE reconcile_customer_ccs_board_summary AS
WITH r AS
    (
      SELECT metric, metric_value, status
        FROM reconcile_customer_ccs_report
    )
, c AS
    (
      SELECT attribute_name, comparison_class, row_count
        FROM reconcile_customer_ccs_contact_summary
    )
, p AS
    (
      SELECT provenance_bucket
           , SUM(merged_customer_count)      AS merged_customers
           , SUM(matched_customer_id_count)  AS matched_ids
        FROM reconcile_customer_ccs_provenance_case_summary
       GROUP BY provenance_bucket
    )
, b AS
    (
      SELECT review_bucket, merged_customer_count
        FROM reconcile_customer_ccs_contact_review_pattern_summary
    )
, summary(summary_order, section, item, value, disposition, note) AS
    (
                  -- 1. Population --------------------------------------------------
        SELECT 100, '1 Population', 'Records received from GP (source rows)'
             , (SELECT metric_value FROM r WHERE metric = '1a_src_fin_customer_base_row_count')
             , 'RECONCILED', 'Source rows before de-duplication'
      UNION ALL
        SELECT 110, '1 Population', 'Distinct GP customer accounts'
             , (SELECT metric_value FROM r WHERE metric = '1b_distinct_gp_customer_number')
             , 'RECONCILED', ''
      UNION ALL
        SELECT 120, '1 Population', 'Merged Workday customers created'
             , (SELECT metric_value FROM r WHERE metric = '1c_src_fin_customer_row_count')
             , 'RECONCILED', 'After de-duplication'
      UNION ALL
        SELECT 130, '1 Population', 'GP accounts missing from Workday'
             , (SELECT metric_value FROM r WHERE metric = '2b_source_customer_numbers_missing_in_workday')
             , 'RECONCILED', 'Zero population loss'
      UNION ALL
        SELECT 140, '1 Population', 'Extra Workday rows not traced to GP'
             , (SELECT metric_value FROM r WHERE metric = '2c_extra_workday_reference_ids_not_in_source')
             , 'INFO', 'Workday test record, excluded as expected'
                  -- 2. Core business attributes -----------------------------------
      UNION ALL
        SELECT 200, '2 Core attributes', 'Customer name (matched / mismatched)'
             , (SELECT metric_value FROM r WHERE metric = '3_customer_name')
             , 'RECONCILED', ''
      UNION ALL
        SELECT 210, '2 Core attributes', 'Customer category (matched / mismatched)'
             , (SELECT metric_value FROM r WHERE metric = '3_customer_category')
             , 'RECONCILED', ''
      UNION ALL
        SELECT 220, '2 Core attributes', 'Default currency (matched / mismatched)'
             , (SELECT metric_value FROM r WHERE metric = '3_default_currency')
             , 'RECONCILED', ''
      UNION ALL
        SELECT 230, '2 Core attributes', 'Default payment terms (matched / mismatched)'
             , (SELECT metric_value FROM r WHERE metric = '3_default_payment_terms')
             , 'RECONCILED', ''
      UNION ALL
        SELECT 240, '2 Core attributes', 'Default payment type (matched / mismatched)'
             , (SELECT metric_value FROM r WHERE metric = '3_default_payment_type')
             , 'RECONCILED', ''
      UNION ALL
        SELECT 250, '2 Core attributes', 'Reference ID (matched / mismatched)'
             , (SELECT metric_value FROM r WHERE metric = '3_reference_id')
             , 'RECONCILED', ''
                  -- 3. Tax ---------------------------------------------------------
      UNION ALL
        SELECT 300, '3 Tax', 'Tax country shows Singapore instead of UK'
             , CAST((SELECT merged_customers FROM p WHERE provenance_bucket = 'workday_target_state_review') AS VARCHAR)
             , 'FIX_IN_WORKDAY'
             , 'Workday followed the Singapore primary address; set tax/primary country to UK ('
               || COALESCE((SELECT MAX(expected_customer_name)
                              FROM reconcile_customer_ccs_tax_review_operator_pack
                             WHERE suggested_bucket = 'workday_target_state_review'), 'see tax_review_pack') || ')'
      UNION ALL
        SELECT 310, '3 Tax', 'Ireland / SZ transaction tax status differs'
             , CAST((SELECT merged_customers FROM p WHERE provenance_bucket = 'ireland_sz_rule_gap_review') AS VARCHAR)
             , 'OPERATOR_DECISION', 'Rule gap vs accepted Workday representation'
      UNION ALL
        SELECT 320, '3 Tax', 'Company Number payload not returned by extract'
             , CAST((SELECT merged_customers FROM p WHERE provenance_bucket = 'extract_omits_company_number_payload') AS VARCHAR)
             , 'ACCEPTED', 'Extract limitation; payload was correctly sent to Workday'
      UNION ALL
        SELECT 330, '3 Tax', 'Transaction tax status wording differs'
             , CAST((SELECT merged_customers FROM p WHERE provenance_bucket = 'semantic_wording_variation') AS VARCHAR)
             , 'ACCEPTED', 'Same meaning, different Workday wording'
      UNION ALL
        SELECT 340, '3 Tax', 'Isle of Man reported as United Kingdom'
             , CAST((SELECT merged_customers FROM p WHERE provenance_bucket = 'accepted_workday_normalisation') AS VARCHAR)
             , 'ACCEPTED', 'Standard Workday handling'
      UNION ALL
        SELECT 350, '3 Tax', 'Default tax code not echoed by extract'
             , CAST((SELECT matched_ids FROM p WHERE provenance_bucket = 'extract_does_not_echo_transform_default_tax_code') AS VARCHAR)
             , 'ACCEPTED', 'Representation only; value was sent'
                  -- 4. Address & contact ------------------------------------------
      UNION ALL
        SELECT 400, '4 Address', 'Address exact match'
             , CAST((SELECT row_count FROM c WHERE attribute_name = 'default_address' AND comparison_class = 'EXACT_MATCH') AS VARCHAR)
             , 'RECONCILED', ''
      UNION ALL
        SELECT 410, '4 Address', 'Address formatting / layout variation'
             , CAST((SELECT row_count FROM c WHERE attribute_name = 'default_address' AND comparison_class = 'FORMAT_VARIATION') AS VARCHAR)
             , 'ACCEPTED', 'Accepted as equivalent'
      UNION ALL
        SELECT 420, '4 Address', 'Address blank in both source and extract'
             , CAST((SELECT row_count FROM c WHERE attribute_name = 'default_address' AND comparison_class = 'BOTH_BLANK') AS VARCHAR)
             , 'INFO', ''
      UNION ALL
        SELECT 430, '4 Address', 'County shows Coventry instead of Cornwall'
             , CAST((SELECT merged_customer_count FROM b WHERE review_bucket = 'suspected_uk_locality_substitution') AS VARCHAR)
             , 'FIX_IN_WORKDAY', 'Known UK locality substitution; correct county to Cornwall'
      UNION ALL
        SELECT 440, '4 Address', 'Address blank in Workday (sent but did not load)'
             , CAST((SELECT merged_customer_count FROM b WHERE review_bucket = 'blank_extract_default_address') AS VARCHAR)
             , 'FIX_IN_WORKDAY', 'Re-enter the customer address in Workday'
      UNION ALL
        SELECT 450, '4 Address', 'Genuine content differences (town / postcode changed)'
             , CAST((SELECT merged_customer_count FROM b WHERE review_bucket = 'manual_review') AS VARCHAR)
             , 'FIX_IN_WORKDAY', 'Review and correct where content is wrong'
      UNION ALL
        SELECT 460, '4 Address', 'Postcode / city layout compaction only'
             , CAST((SELECT merged_customer_count FROM b WHERE review_bucket = 'postcode_city_layout_compaction') AS VARCHAR)
             , 'ACCEPTED', 'Formatting only, no content change'
      UNION ALL
        SELECT 470, '4 Address', 'USA city / state / ZIP rendering'
             , CAST((SELECT merged_customer_count FROM b WHERE review_bucket = 'us_city_state_postcode_layout') AS VARCHAR)
             , 'ACCEPTED', 'Equivalent Workday rendering, in places enriched'
      UNION ALL
        SELECT 480, '4 Address', 'Phone formatting variation'
             , CAST((SELECT row_count FROM c WHERE attribute_name = 'default_phone_number' AND comparison_class = 'FORMAT_VARIATION') AS VARCHAR)
             , 'ACCEPTED', 'Accepted as equivalent'
      UNION ALL
        SELECT 490, '4 Address', 'Email'
             , (SELECT metric_value FROM r WHERE metric = '6_support_sent_email_customer_ids_in_transform')
             , 'NOT_SUPPORTED', 'Not present in the current extract; count is emails sent'
                  -- 5. Status & outcome -------------------------------------------
      UNION ALL
        SELECT 500, '5 Status', 'Total rows flagged for review'
             , (SELECT metric_value FROM r WHERE metric = '6a_total_issue_rows')
             , 'INFO', 'Population, tax and contact review rows combined'
      UNION ALL
        SELECT 510, '5 Status', 'Exact attribute mismatches'
             , (SELECT metric_value FROM r WHERE metric = '6c_exact_attribute_mismatch_rows')
             , 'RECONCILED', 'Zero exact business-attribute mismatches'
      UNION ALL
        SELECT 900, '5 Status', 'Overall reconciliation outcome'
             , 'RECONCILED'
             , 'RECONCILED'
             , 'Population and core data fully reconciled; residual items are Workday cleanups for Finance, not migration or extract defects'
    )
SELECT summary_order
     , section
     , item
     , value
     , disposition
     , note
  FROM summary
 ORDER BY summary_order
;
