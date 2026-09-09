DROP TABLE IF EXISTS workday_supplier_po_receipt
;
CREATE TABLE workday_supplier_po_receipt
    AS
  WITH cte_receipt
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
             FROM workday_supplier_po_commitment_stack                p
            WHERE p.column_ez = t.order_number
                               || '_' || t.line_number
                               || '_' || t.split_position
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
