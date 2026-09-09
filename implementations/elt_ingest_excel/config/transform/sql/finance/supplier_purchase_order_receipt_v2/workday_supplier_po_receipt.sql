DROP TABLE IF EXISTS workday_supplier_po_receipt
;
CREATE TABLE workday_supplier_po_receipt
    AS
  WITH cte_po_split_eligible
    AS (
SELECT p.company_shortcode
     , p.order_number
     , p.line_number
     , p.po_position
  FROM src_fin_supplier_po_commitment_stack_nrm                         p
 WHERE COALESCE(p.soft_commitment_value_gbp, 0)
     + COALESCE(p.hard_commitment_value_gbp, 0)                        > 0
   AND NOT (
           UPPER(TRIM(COALESCE(p.commitment_class, '')))  = 'SOFT'
       AND p.commitment_date                                         < DATE '2025-08-01'
       )
   AND TRIM(UPPER(p.delivery_status))                                != UPPER('Fully Delivered (Closed)')
   AND p.order_number NOT IN (
         'FAPO200010562'
       , 'NFCPO300004541'
       , 'WNSLPO400011283'
       , 'FAPO200035755'
       , 'FAPO200039033'
       , 'WNSLPO400019136'
       , 'FAPO200049511'
       )
       )
     , cte_receipt
    AS (
SELECT t.*
     , ROW_NUMBER() OVER (
           PARTITION BY t.receipt_number
               ORDER BY TRY_CAST(t.line_number        AS INTEGER)
                      , t.line_number
                      , TRY_CAST(t.split_position     AS INTEGER)
                      , t.split_position
       )                                                              receipt_line_rnk
  FROM src_fin_supplier_po_receipt                                    t
 WHERE EXISTS (
           SELECT 1
             FROM cte_po_split_eligible                                e
            WHERE e.company_shortcode = UPPER(TRIM(SPLIT_PART(t.nrm_company_id, '-', -1)))
              AND e.order_number      = t.order_number
              AND e.line_number       = t.line_number
              AND e.po_position       = t.split_position
       )
       )
SELECT
       t.spreadsheet_key                                              column_b    -- spreadsheet_key
     , CASE t.receipt_line_rnk
         WHEN 1 THEN 'Y'
         ELSE 'N'
       END                                                            column_c    -- add_only
     , t.receipt_number                                               column_d    -- receipt
     , CASE t.receipt_line_rnk
         WHEN 1 THEN 'Y'
         ELSE 'N'
       END                                                            column_e    -- auto_complete
     , COALESCE(NULLIF(TRIM(t.invoice_status), ''), '')
       || CASE
            WHEN NULLIF(TRIM(t.invoice_status), '')  IS NOT NULL
             AND NULLIF(TRIM(t.delivery_status), '') IS NOT NULL
            THEN ' '
            ELSE ''
          END
       || COALESCE(NULLIF(TRIM(t.delivery_status), ''), '')
       || COALESCE(
            ' | Source: ' || NULLIF(TRIM(t.nrm_source_tab), '')
          , ''
          )
       || COALESCE(
            CASE
                WHEN UPPER(TRIM(COALESCE(t.commitment_class, ''))) = 'HARD'
                THEN ' | Commitment: HARD'
            END
          , ''
          )                                                           column_f    -- comment
     , CAST(NULL AS VARCHAR)                                          column_g    -- worker
     , t.receipt_number                                               column_h    -- receipt_number
     , CAST(NULL AS VARCHAR)                                          column_i    -- workday_locked
     , 'Y'                                                            column_j    -- submit
     , t.nrm_company_id                                               column_k    -- company
     , t.nrm_receipt_date                                             column_y    -- document_date
     , t.orderitemdescription                                         column_ab   -- memo
     , CAST(t.receipt_line_rnk AS VARCHAR)                            column_bt   -- receipt_service_line_row_id
     , t.order_number || '_' || t.line_number || '_' || t.split_position  column_bu  -- purchase_order_line (column_ez match key)
     , t.nrm_receipt_export_amount_text                               column_bz   -- extended_amount
  FROM cte_receipt                                                    t
 ORDER BY
       column_b
     , receipt_line_rnk
;
