-- ============================================================================
-- RECONCILE: Supplier vendor_id business-unit-collision review (row-level)
--
-- src_fin_supplier_raw.sql splits a GP vendor_id into separate supplier records
-- (nrm_supplier_number suffixed '-{business_unit}') whenever the fully-normalised
-- name disagrees across business units for that vendor_id - GP vendor codes are
-- scoped per company book and get reused for unrelated vendors. That name-based
-- signal is not perfect: a few splits are genuinely the same vendor entered with
-- a typo, an abbreviation, or a spacing/hyphen inconsistency rather than a real
-- identity collision (confirmed cases: 'ZENITH VEHICLE CONTRACT/S', 'PCL CO/ORPORATEWEAR').
-- No safe automated rule separates every one of these from a real collision
-- (e.g. 'READING CITY FC' vs 'READING FC' looks similar but may be two real clubs) -
-- this table surfaces every split ranked by name similarity so the highest-risk
-- false positives can be checked first, instead of reviewing all splits blind.
--
-- pair_similarity: highest pairwise jaro_winkler_similarity between any two of the
-- split names for that vendor_id. High (>0.85) = likely the same vendor, verify
-- before publishing as separate suppliers. Low = confidently distinct vendors.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_supplier_collision_review
;
CREATE TABLE reconcile_supplier_collision_review
    AS
  WITH cte_split_parts
    AS (
-- nrm_agg_supplier_number can be a pipe-list mixing suffixed and unsuffixed parts
-- (Level-2 name-based merging can combine a business-unit-collision split with an
-- unrelated, unsuffixed vendor_id under the same canonical name - e.g.
-- 'UNI012|UNI021-NFC'). Unnest first and test each part independently, rather than
-- pattern-matching the whole joined string, so a suffixed part is never missed
-- just because it shares a pipe-group with an unsuffixed one.
SELECT t.supplier_id                                                                           supplier_id
     , t.nrm_supplier_name                                                                     supplier_name
     , t.nrm_agg_business_unit                                                                 business_unit
     , TRIM(UNNEST(STRING_SPLIT(t.nrm_agg_supplier_number, '|')))                              part
  FROM src_fin_supplier                                                                        t
       )
     , cte_split_groups
    AS (
-- vendor_id itself is never hyphenated (only the '-{business_unit}' suffix
-- introduces one), so any part containing a hyphen was suffixed by the collision
-- split, and splitting on the first hyphen recovers the original vendor_id.
SELECT
       SPLIT_PART(t.part, '-', 1)                                                              gp_vendor_id
     , t.supplier_id                                                                           supplier_id
     , t.supplier_name                                                                         supplier_name
     , t.business_unit                                                                         business_unit
  FROM cte_split_parts                                                                         t
 WHERE t.part                                                                                  LIKE '%-%'
       )
     , cte_pairs
    AS (
SELECT a.gp_vendor_id                                                                          gp_vendor_id
     , jaro_winkler_similarity(a.supplier_name, b.supplier_name)                                pair_similarity
  FROM cte_split_groups                                                                        a
  JOIN cte_split_groups                                                                        b
    ON a.gp_vendor_id                                                                          = b.gp_vendor_id
   AND a.supplier_id                                                                            < b.supplier_id
       )
     , cte_max_similarity
    AS (
SELECT gp_vendor_id
     , MAX(pair_similarity)                                                                    max_pair_similarity
  FROM cte_pairs
 GROUP BY gp_vendor_id
       )
SELECT s.gp_vendor_id                                                                           gp_vendor_id
     , m.max_pair_similarity                                                                    max_pair_similarity
     , CASE WHEN m.max_pair_similarity > 0.85 THEN 'REVIEW - likely same vendor' ELSE 'OK - confidently distinct' END review_flag
     , s.supplier_id                                                                            supplier_id
     , s.supplier_name                                                                          supplier_name
     , s.business_unit                                                                          business_unit
  FROM cte_split_groups                                                                         s
  JOIN cte_max_similarity                                                                       m
    ON m.gp_vendor_id                                                                           = s.gp_vendor_id
 ORDER BY m.max_pair_similarity DESC
     , s.gp_vendor_id
     , s.supplier_id
;
