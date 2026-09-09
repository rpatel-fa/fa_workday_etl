DROP TABLE IF EXISTS workday_supplier_po_commitment_stack_receipt
;
CREATE TABLE workday_supplier_po_commitment_stack_receipt
    AS
  WITH cte_grni
    AS (
-- Receipt EIB sourced ENTIRELY from the commitment stack's HARD commitment
-- (= received-not-invoiced GRNI). One line per PO+Line+Position with hard > 0.
-- The Proactis RECEIPT extract is NOT used: its receipted_quantity is a unit-count
-- (~56% of rows) and its invoice_value contradicts the commitment's fully_invoiced
-- (see FAPO200066017). Sourcing from the commitment makes the receipt EIB consistent
-- with the PO hard commitment by construction and removes the HARD/no-receipt anomalies.
SELECT t.company_shortcode
     , t.order_number
     , t.line_number
     , t.po_position
     , t.po_line_pos_id
     , t.nrm_company_id
     , t.wd_supplier_id
     , t.line_description
     , t.invoice_status
     , t.delivery_status
     , t.commitment_class
     , t.commitment_date
     , t.currency
     , COALESCE(t.hard_commitment_value_local, 0)                                hard_local
     , CASE
           WHEN UPPER(TRIM(COALESCE(t.currency, ''))) = 'JPY'
           THEN CAST(CAST(ROUND(COALESCE(t.hard_commitment_value_local, 0), 0) AS BIGINT)      AS VARCHAR)
           ELSE CAST(CAST(ROUND(COALESCE(t.hard_commitment_value_local, 0), 2) AS DECIMAL(18,2)) AS VARCHAR)
       END                                                                        grni_export_amount_text
  FROM src_fin_supplier_po_commitment_stack                                       t
 WHERE COALESCE(t.hard_commitment_value_gbp, 0)                                 > 0.005   -- has received-not-invoiced GRNI to accrue
       )
     , cte_ranked
    AS (
SELECT t.*
     , 'REC-' || t.order_number                                                  synthetic_receipt_number   -- one receipt document per PO
     , ROW_NUMBER() OVER (
           PARTITION BY t.order_number
               ORDER BY TRY_CAST(NULLIF(TRIM(t.line_number), '') AS INTEGER), t.line_number
                      , TRY_CAST(NULLIF(TRIM(t.po_position), '') AS INTEGER), t.po_position
       )                                                                          receipt_line_rnk
     , ROW_NUMBER() OVER (
           ORDER BY t.nrm_company_id, t.wd_supplier_id, t.order_number
                  , TRY_CAST(NULLIF(TRIM(t.line_number), '') AS INTEGER)
                  , TRY_CAST(NULLIF(TRIM(t.po_position), '') AS INTEGER)
       )                                                                          spreadsheet_key
  FROM cte_grni                                                                   t
       )
SELECT
       t.spreadsheet_key                                              column_b    -- spreadsheet_key
     , CASE t.receipt_line_rnk WHEN 1 THEN 'Y' ELSE 'N' END           column_c    -- add_only
     , t.synthetic_receipt_number                                     column_d    -- receipt (synthetic, one per PO)
     , CASE t.receipt_line_rnk WHEN 1 THEN 'Y' ELSE 'N' END           column_e    -- auto_complete
     , COALESCE(NULLIF(TRIM(t.invoice_status), ''), '')
       || CASE
            WHEN NULLIF(TRIM(t.invoice_status), '')  IS NOT NULL
             AND NULLIF(TRIM(t.delivery_status), '') IS NOT NULL
            THEN ' '
            ELSE ''
          END
       || COALESCE(NULLIF(TRIM(t.delivery_status), ''), '')
       || ' | Source: Commitment Stack (GRNI)'
       || CASE
            WHEN UPPER(TRIM(COALESCE(t.commitment_class, ''))) IN ('HARD', 'MIXED')
            THEN ' | Commitment: ' || TRIM(t.commitment_class)
            ELSE ''
          END                                                         column_f    -- comment
     , CAST(NULL AS VARCHAR)                                          column_g    -- worker
     , t.synthetic_receipt_number                                     column_h    -- receipt_number
     , CAST(NULL AS VARCHAR)                                          column_i    -- workday_locked
     , 'Y'                                                            column_j    -- submit
     , t.nrm_company_id                                               column_k    -- company
     , strftime(t.commitment_date, '%d/%m/%Y')                        column_y    -- document_date (commitment_date proxy — no receipt date in commitment source)
     , t.line_description                                             column_ab   -- memo
     , CAST(t.receipt_line_rnk AS VARCHAR)                            column_bt   -- receipt_service_line_row_id
     , t.po_line_pos_id                                               column_bu   -- purchase_order_line (matches PO column_ez)
     , t.grni_export_amount_text                                      column_bz   -- extended_amount (GRNI = hard commitment, transaction currency)
  FROM cte_ranked                                                     t
 ORDER BY column_b, receipt_line_rnk
;
