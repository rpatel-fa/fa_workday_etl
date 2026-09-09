DROP TABLE IF EXISTS reconcile_customer_invoice_unresolved_workday_update
;
-- "3+" follow-on: the data needed to build the Workday Customers CCS UPDATE.
--
-- For each Workday customer that needs reference_ids added because invoices are
-- unresolved on the normalisation join:
--   (a) every MATCH_ADD_REFERENCE_ID row (clean single-name-match)
--   (b) every AMBIGUOUS_NAME_MATCH row where gp_debtor_on_wd_this_ref = TRUE
--       -- i.e. the Simon Polito rule: pick the Workday duplicate whose
--          existing reference_id already hosts the same GP debtor record
--          that matched by name.
--
-- Output columns map directly to the CCS update sheet:
--   wd_customer_id            Workday customer_id (primary key for the update)
--   wd_customer_name          review label only
--   existing_reference_ids    current pipe-delimited reference_id tokens on the WD customer
--   new_reference_ids         pipe-delimited GP ids that SHOULD be added
--                             (already-excluded duplicates already in existing_reference_ids
--   combined_reference_ids    existing ∪ new, deduped, deterministic sort order
--
-- Token rules: empty / whitespace-only tokens are dropped from all aggregates,
-- combined excludes any token that is already in existing, and sorts alphabetically.
CREATE TABLE reconcile_customer_invoice_unresolved_workday_update
    AS
  WITH cte_action_rows
    AS (
SELECT wd_customer_id
     , wd_customer_name
     , wd_reference_id_current
     , gp_customer_id
  FROM reconcile_customer_invoice_unresolved_workday
 WHERE (
            match_status = 'MATCH_ADD_REFERENCE_ID'
        AND wd_customer_id IS NOT NULL
       )
    OR (
            match_status = 'AMBIGUOUS_NAME_MATCH'
        AND gp_debtor_on_wd_this_ref  = TRUE
        AND wd_customer_id IS NOT NULL
       )
       )
     , cte_all_tokens
    AS (
SELECT DISTINCT
       wd_customer_id
     , wd_customer_name
     , 'EXISTING'                                     tok_type
     , TRIM(UNNEST(STRING_SPLIT(wd_reference_id_current, '|')))   AS token
  FROM cte_action_rows
 WHERE NULLIF(TRIM(wd_reference_id_current), '')         IS NOT NULL
 UNION ALL
SELECT DISTINCT
       wd_customer_id
     , wd_customer_name
     , 'NEW_CAND'                                     tok_type
     , TRIM(UNNEST(STRING_SPLIT(gp_customer_id, '|')))         AS token
  FROM cte_action_rows
 WHERE NULLIF(TRIM(gp_customer_id), '')               IS NOT NULL
       )
     , cte_clean
    AS (
SELECT DISTINCT *
  FROM cte_all_tokens
 WHERE NULLIF(token, '')                                   IS NOT NULL
       )
SELECT c.wd_customer_id
     , MIN(c.wd_customer_name)                        wd_customer_name
     , STRING_AGG(CASE WHEN c.tok_type = 'EXISTING' THEN c.token END, '|'
                 ORDER BY CASE WHEN c.tok_type = 'EXISTING' THEN c.token END)
                                                        existing_reference_ids
     , STRING_AGG(CASE
           WHEN c.tok_type = 'NEW_CAND'
            AND NOT EXISTS (
                  SELECT 1
                    FROM cte_clean e
                   WHERE e.wd_customer_id = c.wd_customer_id
                     AND e.tok_type      = 'EXISTING'
                     AND e.token         = c.token
                )
           THEN c.token
       END, '|'
       ORDER BY CASE WHEN c.tok_type = 'NEW_CAND' THEN c.token END)
                                                        new_reference_ids
     , STRING_AGG(c.token, '|' ORDER BY c.token)       combined_reference_ids
  FROM cte_clean c
 GROUP BY c.wd_customer_id
 ORDER BY c.wd_customer_id
;