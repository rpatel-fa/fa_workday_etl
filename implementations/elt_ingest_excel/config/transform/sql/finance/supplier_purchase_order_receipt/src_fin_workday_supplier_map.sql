DROP TABLE IF EXISTS src_fin_workday_supplier_map
;
-- Shared Workday supplier resolution map: reference token -> Workday supplier_id.
--
-- The Workday supplier extract carries the source vendor id in `reference_id`, in
-- two shapes depending on how suppliers were merged when the Workday population was
-- built:
--   - bare vendor id            e.g. `YOU004`      -> matched by vendor id alone
--   - vendor id + company code  e.g. `YOU002-FA`   -> matched by vendor id + company
-- The suffix is exactly the company code, identical to the source `supplierno`
-- prefix. The SAME vendor id can resolve to DIFFERENT Workday suppliers by company
-- (e.g. YOU002-FA = YouGov vs YOU002-WNSL = Your Smiles).
--
-- The reference token is preserved EXACTLY as it appears in the extract - it is
-- never split or parsed here, because the whole token is the key that maps to a
-- specific supplier_id. Consumers match by constructing the two candidate keys on
-- the SOURCE side: `vendor || '-' || company` (company-exact) first, then bare
-- `vendor` (fallback).
--
-- `reference_id` is pipe-delimited (a Workday supplier can carry several source
-- vendor ids) and can repeat across remit/bank/address rows, so this collapses to
-- one row per token. A single token resolving to multiple supplier_id values is a
-- separate conflict, surfaced by validate_workday_extract_supplier_reference_id_conflict.
CREATE TABLE src_fin_workday_supplier_map
    AS
  WITH cte_token
    AS (
SELECT TRIM(UNNEST(STRING_SPLIT(t.reference_id, '|')))            reference_id_token
     , t.supplier_id
     , t.supplier                                                 nrm_supplier_name
     , t.reference_id                                             nrm_agg_supplier_number
     , CASE
           WHEN COALESCE(
                    NULLIF(TRIM(t.default_po_issue_email), '')
                  , NULLIF(TRIM(t.primary_email_address), '')
                ) IS NOT NULL
           THEN 1
           ELSE 0
       END                                                        nrm_email_exists_ind
  FROM fin_workday_extract_supplier                               t
 WHERE NULLIF(TRIM(t.reference_id), '')                           IS NOT NULL
       )
SELECT reference_id_token
     , MIN(supplier_id)                                           supplier_id
     , MIN(nrm_supplier_name)                                     nrm_supplier_name
     , MIN(nrm_agg_supplier_number)                               nrm_agg_supplier_number
     , MAX(nrm_email_exists_ind)                                  nrm_email_exists_ind
  FROM cte_token
 GROUP BY reference_id_token
;
