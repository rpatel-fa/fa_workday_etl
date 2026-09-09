-- Finance worklist. Turns Gayle's go-live manual-update list into an actionable,
-- extract-verified worklist and consolidates every open item into ONE tab so
-- Finance can work purely from the review pack.

-- Enrich the manual-update list with what the live Workday extract actually shows,
-- so an item already resolved in the tenant is not reworked (e.g. HAR050's bank
-- was on the removal list but loaded successfully and is present).
DROP TABLE IF EXISTS reconcile_supplier_ccs_manual_update_status
;
CREATE TABLE reconcile_supplier_ccs_manual_update_status
    AS
  WITH cte_exploded
    AS (
SELECT mu.seq
     , TRIM(u.supplier_number)                                                       supplier_number
  FROM reconcile_supplier_ccs_manual_update                                          mu
     , UNNEST(STRING_SPLIT(COALESCE(mu.workday_reference_id, ''), '|'))              u(supplier_number)
 WHERE mu.matched_to_extract                                                         = 'Y'
   AND NULLIF(TRIM(u.supplier_number), '')                                           IS NOT NULL
       )
     , cte_state
    AS (
SELECT x.seq
     , BOOL_OR(a.actual_address_present)                                             address_present
     , BOOL_OR(a.actual_bank_present)                                                bank_present
     , BOOL_OR(a.actual_payment_manual_only)                                         payment_manual_only
  FROM cte_exploded                                                                  x
       LEFT OUTER JOIN
       reconcile_supplier_ccs_actual_supplier                                        a
          ON a.supplier_number                                                       = x.supplier_number
 GROUP BY
       x.seq
       )
SELECT mu.exception_type
     , mu.supplier_name_raw
     , mu.workday_supplier_name
     , mu.workday_reference_id
     , mu.workday_supplier_id
     , mu.matched_to_extract
     , CASE
         WHEN mu.matched_to_extract = 'N'                                THEN 'Review — name not matched to extract'
         WHEN mu.exception_type = 'ADDRESS' AND COALESCE(s.address_present, FALSE)     THEN 'No — address already present in Workday'
         WHEN mu.exception_type = 'BANK'    AND COALESCE(s.bank_present, FALSE)        THEN 'No — bank already present in Workday'
         WHEN mu.exception_type = 'PAYMENT' AND NOT COALESCE(s.payment_manual_only, FALSE) THEN 'No — payment types already present in Workday'
         ELSE 'Yes'
       END                                                                           still_open_in_workday
     , CASE mu.exception_type
         WHEN 'ADDRESS' THEN 'Re-enter the supplier address in Workday Production'
         WHEN 'BANK'    THEN 'Re-enter the supplier bank / settlement account in Production'
         WHEN 'PAYMENT' THEN 'Restore the correct default / accepted payment types in Production'
         ELSE ''
       END                                                                           finance_action
  FROM reconcile_supplier_ccs_manual_update                                          mu
       LEFT OUTER JOIN
       cte_state                                                                     s
          ON s.seq                                                                   = mu.seq
 ORDER BY
       mu.exception_type
     , mu.supplier_name_raw
;

-- THE Finance worklist: every open item to fix in Workday, in one place.
DROP TABLE IF EXISTS reconcile_supplier_ccs_finance_action_list
;
CREATE TABLE reconcile_supplier_ccs_finance_action_list
    AS
SELECT * FROM (
  SELECT 1                                                                            sort_order
       , exception_type                                                               area
       , workday_reference_id                                                         reference_id
       , COALESCE(workday_supplier_name, supplier_name_raw)                           supplier_name
       , finance_action                                                               finance_action
       , 'Go-live manual-update list'                                                 item_source
    FROM reconcile_supplier_ccs_manual_update_status
   WHERE still_open_in_workday                                                        = 'Yes'
  UNION ALL
  SELECT 2
       , 'ADDRESS (residual)'
       , expected_reference_id
       , expected_supplier_name
       , 'Add / correct the supplier address — genuine source-data gap, not on the manual list'
       , 'Reconcile — attribute gap'
    FROM reconcile_supplier_ccs_attribute_gap
   WHERE gap_type                                                                     = 'ADDRESS'
     AND gap_status                                                                   = 'UNEXPLAINED'
  UNION ALL
  SELECT 3
       , 'REMOVE'
       , workday_reference_id
       , workday_supplier_name
       , 'Remove this test / non-GP supplier from the tenant'
       , 'Reconcile — dummy row'
    FROM reconcile_supplier_ccs_workday_dummy_row
) ORDER BY sort_order, area, supplier_name
;
