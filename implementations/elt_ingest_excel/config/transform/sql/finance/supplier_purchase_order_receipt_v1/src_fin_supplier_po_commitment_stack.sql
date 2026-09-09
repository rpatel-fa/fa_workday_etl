DROP TABLE IF EXISTS src_fin_supplier_po_commitment_stack;

CREATE TABLE src_fin_supplier_po_commitment_stack
    AS
  WITH cte_keyed
    AS (
SELECT t.*
     , ROW_NUMBER() OVER (
         ORDER BY
               t.order_number
             , TRY_CAST(NULLIF(TRIM(t.line_number), '') AS INTEGER)
             , TRY_CAST(NULLIF(TRIM(t.po_position), '') AS INTEGER)
       )                                                                        spreadsheet_key
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
SELECT
       t.*
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
;
