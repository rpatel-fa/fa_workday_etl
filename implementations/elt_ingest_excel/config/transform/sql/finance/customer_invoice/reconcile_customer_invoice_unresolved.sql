DROP TABLE IF EXISTS reconcile_customer_invoice_unresolved
;
-- Triage for invoiced GP customers that do NOT resolve on reference_id token
-- (src_fin_customer_invoice_nrm.wd_customer_id IS NULL). Classifies each into:
--   MATCH_ADD_REFERENCE_ID = the customer already exists in Workday (matched by
--                            normalised name); fix = add the GP id to that Workday
--                            customer's reference_id, then it resolves on the next extract.
--   GENUINE_NEW            = no Workday name match; a real new-customer create.
--
-- This table is a raw one-row-per-(GP customer x Workday match) grid. It does NOT
-- consolidate multiple GP ids onto the same Workday customer and does NOT flag
-- ambiguous name collisions - those semantics are layered on in
-- reconcile_customer_invoice_unresolved_workday.sql so this table remains a clean
-- lineage input.
--
-- Name normalisation is the src_fin_customer_raw.sql chain (strip accents/dots,
-- WOMEN'S->WOMEN, FOOTBALL CLUB->FC, drop (...) qualifiers, UNITED->UTD, AND->&,
-- strip trailing LIMITED/LTD/COMPANY/PLC/LLP/INC) PLUS two steps that chain does NOT
-- apply: hyphen->space and whitespace-collapse. Those two are what let
-- "Coca -Cola Europacific Partners" (GP) match "COCA COLA EUROPACIFIC PARTNERS"
-- (C-000808) without colliding with "...GREAT BRITAIN" (C-000809). Applied to BOTH sides.
--
-- Review report only: the transform itself resolves on reference_id token, never on
-- name. A name join can mis-bind an invoice to the wrong legal entity, so it is
-- surfaced for human action (add the token in Workday), never applied automatically.
CREATE TABLE reconcile_customer_invoice_unresolved
    AS
  WITH cte_unresolved
    AS (
SELECT DISTINCT
       TRIM(n.business_unit)                        gp_business_unit
     , TRIM(n.customer_id)                          gp_customer_id
     , TRIM(n.customer_name)                        gp_customer_name
  FROM src_fin_customer_invoice_nrm                 n
 WHERE n.wd_customer_id                             IS NULL
       )
SELECT CASE
           WHEN w.customer_id IS NOT NULL
           THEN 'MATCH_ADD_REFERENCE_ID'
           ELSE 'GENUINE_NEW'
       END                                          match_status
     , u.gp_business_unit
     , u.gp_customer_id
     , u.gp_customer_name
     , w.customer_id                                wd_customer_id
     , w.customer                                   wd_customer_name
     , w.reference_id                               wd_reference_id_current
  FROM cte_unresolved                               u
       LEFT OUTER JOIN
       fin_workday_extract_customer                 w
          ON TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(
                REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(unaccent(w.customer)), '.', ''), '-', ' '), 'WOMEN''S', 'WOMEN'), 'WOMENS', 'WOMEN'), 'FOOTBALL CLUB', 'FC'),
                '\s*\([^)]*\)', '', 'g'), '\bUNITED\b', 'UTD', 'g'), '\bAND\b', '&', 'g'),
                '(\s+|^)((LIMITED|LTD|COMPANY|PLC|LLP|INC)\s*)+$', ''), '\s+', ' ', 'g'))
           =  TRIM(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(
                REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(unaccent(u.gp_customer_name)), '.', ''), '-', ' '), 'WOMEN''S', 'WOMEN'), 'WOMENS', 'WOMEN'), 'FOOTBALL CLUB', 'FC'),
                '\s*\([^)]*\)', '', 'g'), '\bUNITED\b', 'UTD', 'g'), '\bAND\b', '&', 'g'),
                '(\s+|^)((LIMITED|LTD|COMPANY|PLC|LLP|INC)\s*)+$', ''), '\s+', ' ', 'g'))
 ORDER BY match_status
        , u.gp_business_unit
        , u.gp_customer_name
;
