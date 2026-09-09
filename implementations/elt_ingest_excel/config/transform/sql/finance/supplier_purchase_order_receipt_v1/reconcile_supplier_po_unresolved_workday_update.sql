DROP TABLE IF EXISTS reconcile_supplier_po_unresolved_workday_update
;
-- "3+" follow-on: the data needed to build the Workday Suppliers CCS UPDATE.
--
-- Action rows = every MATCH_ADD_REFERENCE_ID row (clean single-name-match) PLUS
-- every AMBIGUOUS_NAME_MATCH row where gp_creditor_on_wd_this_ref = TRUE.
--
-- Output columns map to the CCS update sheet:
--   wd_supplier_id            Workday supplier_id (primary key for the update)
--   wd_supplier_name          review label only
--   existing_reference_ids    current pipe-delimited reference_id tokens on the WD supplier
--   new_reference_ids         pipe-delimited GP supplier codes to ADD (tokens not
--                             already in existing_reference_ids are excluded)
--   combined_reference_ids    existing ∪ new, deduped, deterministic sort order
--
-- Note on supplier tokens: the reference_id in fin_workday_extract_supplier can
-- carry BOTH a bare vendor id (e.g. YOU004) and a company-qualified vendor id
-- (e.g. YOU002-FA). Consumers of the CCS update that write back should write
-- BOTH shapes to ensure the two-candidate-key match on the next ingest works
-- regardless of how the token was originally encoded.
CREATE TABLE reconcile_supplier_po_unresolved_workday_update
    AS
  WITH cte_action_rows
    AS (
SELECT wd_supplier_id
     , wd_supplier_name
     , wd_reference_id_current
     , gp_supplier_id
  FROM reconcile_supplier_po_unresolved_workday
 WHERE (
            match_status = 'MATCH_ADD_REFERENCE_ID'
        AND wd_supplier_id IS NOT NULL
       )
    OR (
            match_status = 'AMBIGUOUS_NAME_MATCH'
        AND gp_creditor_on_wd_this_ref = TRUE
        AND wd_supplier_id IS NOT NULL
       )
       )
     , cte_all_tokens
    AS (
SELECT DISTINCT
       wd_supplier_id
     , wd_supplier_name
     , 'EXISTING'                                     tok_type
     , TRIM(UNNEST(STRING_SPLIT(wd_reference_id_current, '|')))   AS token
  FROM cte_action_rows
 WHERE NULLIF(TRIM(wd_reference_id_current), '')         IS NOT NULL
 UNION ALL
SELECT DISTINCT
       wd_supplier_id
     , wd_supplier_name
     , 'NEW_CAND'                                     tok_type
     , TRIM(UNNEST(STRING_SPLIT(gp_supplier_id, '|')))         AS token
  FROM cte_action_rows
 WHERE NULLIF(TRIM(gp_supplier_id), '')               IS NOT NULL
       )
     , cte_clean
    AS (
SELECT DISTINCT *
  FROM cte_all_tokens
 WHERE NULLIF(token, '')                                   IS NOT NULL
       )
SELECT c.wd_supplier_id
     , MIN(c.wd_supplier_name)                        wd_supplier_name
     , STRING_AGG(CASE WHEN c.tok_type = 'EXISTING' THEN c.token END, '|'
                 ORDER BY CASE WHEN c.tok_type = 'EXISTING' THEN c.token END)
                                                        existing_reference_ids
     , STRING_AGG(CASE
           WHEN c.tok_type = 'NEW_CAND'
            AND NOT EXISTS (
                  SELECT 1
                    FROM cte_clean e
                   WHERE e.wd_supplier_id = c.wd_supplier_id
                     AND e.tok_type      = 'EXISTING'
                     AND e.token         = c.token
                )
           THEN c.token
       END, '|'
       ORDER BY CASE WHEN c.tok_type = 'NEW_CAND' THEN c.token END)
                                                        new_reference_ids
     , STRING_AGG(c.token, '|' ORDER BY c.token)       combined_reference_ids
  FROM cte_clean c
 GROUP BY c.wd_supplier_id
 ORDER BY c.wd_supplier_id
;