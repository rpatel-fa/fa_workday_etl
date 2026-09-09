DROP TABLE IF EXISTS reconcile_customer_ccs_closure_decision_pack
;
CREATE TABLE reconcile_customer_ccs_closure_decision_pack
    AS
SELECT ROW_NUMBER() OVER(
           ORDER BY CASE t.issue_bucket
                      WHEN 'workday_target_state_review'                    THEN 1
                      WHEN 'ireland_sz_rule_gap_review'                     THEN 2
                      WHEN 'accepted_workday_normalisation'                 THEN 3
                      WHEN 'default_address_value_mismatch'                 THEN 4
                      WHEN 'address_review_bucket:postcode_city_layout_compaction'
                                                                          THEN 5
                      WHEN 'address_review_bucket:us_city_state_postcode_layout'
                                                                          THEN 6
                      WHEN 'address_review_bucket:country_specific_region_layout'
                                                                          THEN 7
                      WHEN 'address_review_bucket:blank_extract_default_address'
                                                                          THEN 8
                      WHEN 'address_review_bucket:suspected_uk_locality_substitution'
                                                                          THEN 9
                      WHEN 'address_review_bucket:manual_review'           THEN 10
                      ELSE 99
                    END
                  , t.closure_order
       )                                                                            decision_order
     , t.issue_area
     , t.issue_bucket
     , t.issue_scope
     , t.sent_vs_extract
     , t.source_customer_number_count
     , t.merged_customer_count
     , t.matched_customer_id_count
     , t.evidence_table
     , CASE t.issue_bucket
         WHEN 'workday_target_state_review'
         THEN 'If Workday is deriving tax country from the Singapore primary address, should this customer still be treated as UK for tax and raised as a Workday issue?'
         WHEN 'ireland_sz_rule_gap_review'
         THEN 'Should this Ireland SZ tax-status pattern be accepted as Workday representation, or treated as a transform rule gap?'
         WHEN 'accepted_workday_normalisation'
         THEN 'Record Isle of Man -> United Kingdom as accepted Workday handling?'
         WHEN 'default_address_value_mismatch'
         THEN 'Do any remaining residual address buckets deserve another equivalence rule pass before defect treatment?'
         WHEN 'address_review_bucket:postcode_city_layout_compaction'
         THEN 'Are postcode/city compaction variants an accepted formatting equivalence for this extract?'
         WHEN 'address_review_bucket:us_city_state_postcode_layout'
         THEN 'Are US city/state/postcode layout variants an accepted formatting equivalence for this extract?'
         WHEN 'address_review_bucket:country_specific_region_layout'
         THEN 'Do these residual country-specific region layouts match approved local formatting rules?'
         WHEN 'address_review_bucket:blank_extract_default_address'
         THEN 'Should a blank extract default address be treated as a true defect or an expected Workday/extract behaviour for these rows?'
         WHEN 'address_review_bucket:suspected_uk_locality_substitution'
         THEN 'Do these UK locality substitutions indicate a true address defect rather than layout drift?'
         WHEN 'address_review_bucket:manual_review'
         THEN 'Do these rows represent genuine content changes that should remain as defects after all approved equivalence rules are applied?'
         ELSE 'Review remaining decision bucket'
       END                                                                         decision_question
     , CASE t.issue_bucket
         WHEN 'workday_target_state_review'
         THEN 'ESCALATE_WORKDAY_TARGET_STATE'
         WHEN 'ireland_sz_rule_gap_review'
         THEN 'DECIDE_EQUIVALENCE_OR_RULE_GAP'
         WHEN 'accepted_workday_normalisation'
         THEN 'RECORD_ACCEPTED_WORKDAY_HANDLING'
         WHEN 'default_address_value_mismatch'
         THEN 'USE_BUCKET_ROWS_BELOW_FOR_DECISION'
         WHEN 'address_review_bucket:postcode_city_layout_compaction'
         THEN 'ACCEPT_EQUIVALENCE_CANDIDATE'
         WHEN 'address_review_bucket:us_city_state_postcode_layout'
         THEN 'ACCEPT_EQUIVALENCE_CANDIDATE'
         WHEN 'address_review_bucket:country_specific_region_layout'
         THEN 'ACCEPT_EQUIVALENCE_CANDIDATE'
         WHEN 'address_review_bucket:blank_extract_default_address'
         THEN 'TRUE_DEFECT_CANDIDATE'
         WHEN 'address_review_bucket:suspected_uk_locality_substitution'
         THEN 'TRUE_DEFECT_CANDIDATE'
         WHEN 'address_review_bucket:manual_review'
         THEN 'TRUE_DEFECT_CANDIDATE'
         ELSE 'MANUAL_DECISION'
       END                                                                         recommended_resolution
     , CASE t.issue_bucket
         WHEN 'workday_target_state_review'
         THEN 'Check whether business ownership expects UK tax treatment; if yes, raise in Workday rather than changing the transform'
         WHEN 'ireland_sz_rule_gap_review'
         THEN 'Check whether Workday uses a different but acceptable tax representation for Ireland/SZ'
         WHEN 'accepted_workday_normalisation'
         THEN 'Accepted Workday handling; keep as an audit note, not an open transform issue'
         WHEN 'default_address_value_mismatch'
         THEN 'Use this as the parent row for the address buckets below'
         WHEN 'address_review_bucket:postcode_city_layout_compaction'
         THEN 'Largest remaining address bucket; best candidate if you want one more burn-down rule'
         WHEN 'address_review_bucket:us_city_state_postcode_layout'
         THEN 'Good candidate for an explicit country-specific equivalence rule'
         WHEN 'address_review_bucket:country_specific_region_layout'
         THEN 'Small bucket; likely low-cost to resolve by policy'
         WHEN 'address_review_bucket:blank_extract_default_address'
         THEN 'Blank extract values are usually stronger evidence of a true non-round-trip difference'
         WHEN 'address_review_bucket:suspected_uk_locality_substitution'
         THEN 'Content looks changed, not merely reformatted'
         WHEN 'address_review_bucket:manual_review'
         THEN 'Residual bucket after all current heuristics; likely true defects unless a new pattern emerges'
         ELSE t.operator_notes
       END                                                                         reviewer_notes
  FROM reconcile_customer_ccs_closure_summary                                       t
 WHERE t.operator_action                                                            = 'DECIDE_OPERATOR_REVIEW'
 ORDER BY
       decision_order
;
