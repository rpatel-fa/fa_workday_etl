DROP TABLE IF EXISTS src_fin_workday_customer_map
;
-- Shared Workday customer resolution map: reference token -> Workday customer_id.
--
-- The Workday customer extract carries the source GP customer id in `reference_id`.
-- Unlike the supplier extract, customer reference tokens are BARE (no company-code
-- suffix): the same source id resolves to a single Workday customer regardless of
-- company. Consumers therefore match on the bare `customer_id`; the company-exact
-- key (`customer || '-' || company`) is kept for parity with the supplier map but
-- currently resolves nothing.
--
-- The reference token is preserved EXACTLY as it appears in the extract - it is
-- never split or parsed here, because the whole token is the key that maps to a
-- specific customer_id.
--
-- `reference_id` is pipe-delimited (a Workday customer can carry several source
-- customer ids) and can repeat across rows, so this collapses to one row per token.
CREATE TABLE src_fin_workday_customer_map
    AS
  WITH cte_token
    AS (
SELECT TRIM(UNNEST(STRING_SPLIT(t.reference_id, '|')))            reference_id_token
     , t.customer_id
     , t.customer                                                 nrm_customer_name
     , t.reference_id                                             nrm_agg_customer_number
  FROM fin_workday_extract_customer                               t
 WHERE NULLIF(TRIM(t.reference_id), '')                           IS NOT NULL
       )
SELECT reference_id_token
     , MIN(customer_id)                                           customer_id
     , MIN(nrm_customer_name)                                     nrm_customer_name
     , MIN(nrm_agg_customer_number)                               nrm_agg_customer_number
  FROM cte_token
 GROUP BY reference_id_token
;
