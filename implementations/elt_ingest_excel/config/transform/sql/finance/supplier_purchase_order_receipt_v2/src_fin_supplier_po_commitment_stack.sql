DROP TABLE IF EXISTS src_fin_supplier_po_commitment_stack;

CREATE TABLE src_fin_supplier_po_commitment_stack
    AS
  WITH cte_po_base_filtered
    AS (
-- Business pre-gates only. PO is a COMMITMENT (no P&L) and the requirement is
-- PO = hard + soft (FULL, un-netted): the already-invoiced overlap is handled on the
-- RECEIPT side only. No receipt data is referenced anywhere in this file.
SELECT t.*
  FROM src_fin_supplier_po_commitment_stack_nrm                                 t
 WHERE COALESCE(t.soft_commitment_value_gbp, 0) + COALESCE(t.hard_commitment_value_gbp, 0) > 0
   AND NOT
       (
       UPPER(TRIM(COALESCE(t.commitment_class, '')))  = 'SOFT'
   AND t.commitment_date < DATE '2025-08-01'
       )
   AND TRIM(UPPER(t.delivery_status))                 != UPPER('Fully Delivered (Closed)')
   AND t.order_number NOT IN (
         'FAPO200010562'
       , 'NFCPO300004541'
       , 'WNSLPO400011283'
       , 'FAPO200035755'
       , 'FAPO200039033'
       , 'WNSLPO400019136'
       , 'FAPO200049511'
       )
       )
     , cte_keyed
    AS (
SELECT t.*
     , ROW_NUMBER() OVER (
         ORDER BY
               t.order_number
             , TRY_CAST(NULLIF(TRIM(t.line_number), '') AS INTEGER)
             , TRY_CAST(NULLIF(TRIM(t.po_position), '') AS INTEGER)
       )                                                                        spreadsheet_key
  FROM cte_po_base_filtered                                                      t
       )
SELECT
       t.*
     -- Full gross commitment (no netting): remaining = the committed value.
     , ROUND(GREATEST(COALESCE(t.soft_commitment_value_gbp,   0), 0), 2)          nrm_soft_remaining_gbp
     , ROUND(GREATEST(COALESCE(t.hard_commitment_value_gbp,   0), 0), 2)          nrm_hard_remaining_gbp
     , ROUND(GREATEST(COALESCE(t.soft_commitment_value_local, 0), 0), 2)          nrm_soft_remaining_local
     , ROUND(GREATEST(COALESCE(t.hard_commitment_value_local, 0), 0), 2)          nrm_hard_remaining_local
     , CAST(1.0 AS DOUBLE)                                                        nrm_remaining_share            -- full gross (kept for downstream compatibility)
     , 'N'                                                                        nrm_po_receipt_rescued_ind     -- no netting/rescue in the full-gross model
     -- EIB extended_amount (transaction currency) = full soft + hard commitment.
     , ROUND(COALESCE(t.soft_commitment_value_local, 0)
           + COALESCE(t.hard_commitment_value_local, 0), 2)                       nrm_po_export_local
     , ROW_NUMBER() OVER (
         PARTITION BY
               t.order_number
         ORDER BY
               TRY_CAST(NULLIF(TRIM(t.line_number), '') AS INTEGER)
             , TRY_CAST(NULLIF(TRIM(t.po_position), '') AS INTEGER)
                         )                                        po_row_id
     , ROW_NUMBER() OVER (
         PARTITION BY
               t.order_number
             , t.line_number
         ORDER BY
               TRY_CAST(NULLIF(TRIM(t.po_position), '') AS INTEGER)
                         )                                        po_line_row_id
     , t.order_number
       || '_' || t.line_number
       || '_' || t.po_position                                    po_line_pos_id
  FROM cte_keyed                                                  t
 WHERE COALESCE(t.soft_commitment_value_gbp, 0)
     + COALESCE(t.hard_commitment_value_gbp, 0)                 > 0.005
;
