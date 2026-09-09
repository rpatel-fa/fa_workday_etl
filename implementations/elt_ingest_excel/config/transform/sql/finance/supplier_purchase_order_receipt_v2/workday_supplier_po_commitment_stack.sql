DROP TABLE IF EXISTS workday_supplier_po_commitment_stack
;
CREATE TABLE workday_supplier_po_commitment_stack
    AS
SELECT
       DENSE_RANK() OVER (
         ORDER BY
               t.nrm_company_id
             , t.wd_supplier_id
             , t.order_number
                         )                                                    column_b    -- spreadsheet_key
     , 'Y'                                                                    column_c    -- add_only
     , 'Y'                                                                    column_d    -- auto_complete
     , COALESCE(NULLIF(TRIM(t.invoice_status), ''), '')
       || CASE
            WHEN NULLIF(TRIM(t.invoice_status), '')  IS NOT NULL
             AND NULLIF(TRIM(t.delivery_status), '') IS NOT NULL
            THEN ' | '
            ELSE ''
          END
       || COALESCE(NULLIF(TRIM(t.delivery_status), ''), '')
       || ' | Source: Commitment Stack Report'
       || CASE
            WHEN UPPER(TRIM(COALESCE(t.commitment_class, ''))) IN ('SOFT', 'HARD', 'MIXED')
            THEN ' | Commitment: ' || TRIM(t.commitment_class)
            ELSE ''
          END
       || ' | Period: ' || COALESCE(NULLIF(TRIM(t.reporting_period), ''), '')
       || CASE
            WHEN t.nrm_po_receipt_rescued_ind = 'Y'
            THEN ' | RECEIPT-BACKED (fully invoiced; PO carried for receipt referential integrity)'
            ELSE ''
          END
                                                                              column_e    -- comment
     , 'Y'                                                                    column_h    -- submit
     , t.order_number                                                         column_j    -- document_number
     , t.nrm_company_id                                                       column_p    -- company
     , t.wd_supplier_id                                                       column_q    -- supplier
     , strftime(t.order_date, '%d/%m/%Y')                                     column_u    -- document_date
     , strftime(t.commitment_date, '%d/%m/%Y')                                column_ae   -- due_date
     , t.currency                                                             column_ah   -- currency
     , CASE t.wd_supplier_email_exists_ind
           WHEN 1 THEN 'EMAIL'
           ELSE 'PRINT'
       END                                                                    column_al   -- issue_option
     , CAST(NULL AS VARCHAR)                                                  column_az   -- memo
     , CAST(NULL AS VARCHAR)                                                  column_ba   -- internal memo
     , CAST(NULL AS VARCHAR)                                                  column_bq   -- Goods Line rowid
     , CAST(NULL AS VARCHAR)                                                  column_bu   -- line_company
     , CAST(NULL AS VARCHAR)                                                  column_bz   -- item_description
     , CAST(NULL AS VARCHAR)                                                  column_cj   -- spend category
     , CAST(NULL AS VARCHAR)                                                  column_dg   -- unit of measure
     , t.po_row_id                                                            column_ex   -- Service Line rowid
     , t.po_line_pos_id                                                       column_ez   -- service_order_line_id
     , t.po_row_id                                                            column_fa   -- line_number
     , t.nrm_company_id                                                       column_fb   -- line_company
     , t.line_description                                                     column_fc   -- item_description
     , COALESCE(
           r.spend_category_override_ref_id
         , r.spend_category_ref_id
         , 'Other_Opex'
       )                                                                      column_fk   -- resource_category
     , CASE
           WHEN UPPER(TRIM(COALESCE(t.currency, ''))) = 'JPY'
           THEN CAST(CAST(ROUND(t.nrm_po_export_local, 0) AS BIGINT)      AS VARCHAR)   -- JPY = whole yen, no decimals
           ELSE CAST(CAST(ROUND(t.nrm_po_export_local, 2) AS DECIMAL(18,2)) AS VARCHAR)
       END                                                                    column_gf   -- extended_amount (full gross soft_local+hard_local; JPY-integer to match receipt EIB)
     , r.cost_centre_ref_id                                                   column_gv   -- cost_centre
     , r.reporting_type                                                       column_gw   -- reporting_type
     , r.event_ref_id                                                         column_gx   -- events
     , r.fa_learning_course_ref_id                                            column_gy   -- fa_learning_courses
     , r.camp_ref_id                                                          column_gz   -- camps
     , r.program_ref_id                                                       column_ha   -- programmes
     , r.grant_ref_id                                                         column_hb   -- grant
     , r.commercial_partner_ref_id                                            column_hc   -- commercial_partner
     , COALESCE(r.project_capex_ref_id, r.project_opex_ref_id)                column_hd   -- project
  FROM src_fin_supplier_po_commitment_stack                            t
       LEFT OUTER JOIN
       src_fin_supplier_po_commitment_stack_ref                        r
          ON t.spreadsheet_key                                         = r.spreadsheet_key
-- WHERE t.wd_supplier_id IS NOT NULL
 ORDER BY
       column_b
     , column_ex
;
