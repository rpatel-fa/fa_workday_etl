DROP TABLE IF EXISTS validate_workday_extract_supplier_reference_id_conflict
;
-- Workday supplier extract integrity.
--
-- Workday emits multiple rows per supplier (address / remit-to / bank-account
-- variants).  Those are benign, so the scan runs against the distinct
-- (reference_id, supplier_id) projection — only then does a multi-row conflict
-- mean the Workday master data is actually wrong.
--
-- Two conflict shapes are detected and both are emitted as rule_type=CONFLICT
-- into validation_supplier_purchase_order_receipt_issues:
--
--   reference_id     — the WHOLE pipe-delimited reference_id string (e.g.
--                      "YOU002-FA|YOU002") resolves to >1 Workday supplier_id.
--
--   source_supplier_no — a SINGLE unnested reference_id token (e.g. "YOU002-FA"
--                      or "YOU002") resolves to >1 Workday supplier_id.
--                      This is the GP supplier-id → Workday supplier sanity
--                      check: one unnested GP supplier-id = exactly one WD
--                      supplier per company token.
--
-- For the single-token (source_supplier_no) case we also carry the
-- canonical_supplier_id that src_fin_workday_supplier_map's MIN() would
-- silently pick, so reviewers can see what the collapse resolves to.
CREATE TABLE validate_workday_extract_supplier_reference_id_conflict
    AS
  WITH cte_distinct_extract
    AS (
SELECT DISTINCT
       t.reference_id
     , t.supplier_id
     , COALESCE(t.supplier, '')                                                    supplier
  FROM fin_workday_extract_supplier                                                t
 WHERE NULLIF(TRIM(t.reference_id), '')                                            IS NOT NULL
       )
     , cte_exact_reference_id
    AS (
SELECT 'reference_id'                                                              conflict_key_type
     , t.reference_id                                                              conflict_key
     , COUNT(*)                                                                    extract_row_count
     , COUNT(DISTINCT t.supplier_id)                                               supplier_id_count
     , STRING_AGG(DISTINCT t.supplier_id, '|' ORDER BY t.supplier_id)             supplier_ids
     , STRING_AGG(DISTINCT t.supplier, '|' ORDER BY t.supplier)                    supplier_names
     , STRING_AGG(DISTINCT t.reference_id, '|' ORDER BY t.reference_id)           reference_ids
     , CAST(NULL AS VARCHAR)                                                       canonical_supplier_id
  FROM cte_distinct_extract                                                        t
 GROUP BY
       t.reference_id
HAVING COUNT(DISTINCT t.supplier_id) > 1
       )
     , cte_exploded_token
    AS (
SELECT TRIM(u.source_supplier_no)                                                  source_supplier_no
     , COUNT(*)                                                                    extract_row_count
     , COUNT(DISTINCT t.supplier_id)                                               supplier_id_count
     , STRING_AGG(DISTINCT t.supplier_id, '|' ORDER BY t.supplier_id)             supplier_ids
     , STRING_AGG(DISTINCT t.supplier, '|' ORDER BY t.supplier)                    supplier_names
     , STRING_AGG(DISTINCT t.reference_id, '|' ORDER BY t.reference_id)           reference_ids
     , MIN(t.supplier_id)                                                          canonical_supplier_id
  FROM cte_distinct_extract                                                        t
     , UNNEST(STRING_SPLIT(t.reference_id, '|'))                                   u(source_supplier_no)
 WHERE NULLIF(TRIM(u.source_supplier_no), '')                                      IS NOT NULL
 GROUP BY
       TRIM(u.source_supplier_no)
HAVING COUNT(DISTINCT t.supplier_id) > 1
       )
     , cte_exploded_reference_id
    AS (
SELECT 'source_supplier_no'                                                        conflict_key_type
     , source_supplier_no                                                          conflict_key
     , extract_row_count
     , supplier_id_count
     , supplier_ids
     , supplier_names
     , reference_ids
     , canonical_supplier_id
  FROM cte_exploded_token
       )
SELECT conflict_key_type
     , conflict_key
     , extract_row_count
     , supplier_id_count
     , supplier_ids
     , supplier_names
     , reference_ids
     , canonical_supplier_id
  FROM cte_exact_reference_id
 UNION ALL
SELECT conflict_key_type
     , conflict_key
     , extract_row_count
     , supplier_id_count
     , supplier_ids
     , supplier_names
     , reference_ids
     , canonical_supplier_id
  FROM cte_exploded_reference_id
;

INSERT INTO validation_supplier_purchase_order_receipt_issues
SELECT 'fin_workday_extract_supplier'                                                      table_name
     , 'CONFLICT'                                                                          rule_type
     , conflict_key_type                                                                   col
     , conflict_key                                                                        po_id
     , 'supplier_id_count=' || COALESCE(CAST(supplier_id_count AS VARCHAR), '?')
       || '|ids=' || COALESCE(supplier_ids, '')
       || '||' || COALESCE(reference_ids, '')
       || CASE
              WHEN canonical_supplier_id IS NOT NULL
              THEN '|MIN_picks=' || canonical_supplier_id
              ELSE ''
          END                                                                              detail
     , 'Workday supplier extract reference_id mapping must not resolve to multiple Workday supplier_id values.'
       || CASE
              WHEN conflict_key_type = 'source_supplier_no'
              THEN '  (GP supplier-id token ' || COALESCE(conflict_key, '<NULL>') || ')'
              ELSE ''
          END                                                                              message
  FROM validate_workday_extract_supplier_reference_id_conflict
;
