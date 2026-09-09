DROP TABLE IF EXISTS src_fin_supplier_po_receipt_nrm;

CREATE TABLE src_fin_supplier_po_receipt_nrm
    AS
  WITH cte_src_receipt_base
    AS (
SELECT DISTINCT
       t.*
  FROM src_fin_supplier_po_receipt_base                                             t
 WHERE UPPER(TRIM(COALESCE(t.order_item_cancellation_status, 'False'))) <> 'TRUE'
   AND UPPER(TRIM(COALESCE(t.receipted_quantity_text, 'x'))) <> 'NULL'
QUALIFY ROW_NUMBER() OVER (
            PARTITION BY t.company_shortcode, t.order_number, t.line_number
                       , t.po_position, t.receipt_number, t.proactis_receipt_no
                       , t.receipted_quantity_text, t.currency
            ORDER BY COALESCE(TRY_CAST(NULLIF(t.receipt_created_date_excel, '') AS TIMESTAMP), DATE '1900-01-01') DESC
                   , COALESCE(TRY_CAST(NULLIF(t.order_closed_for_invoicing, '')    AS TIMESTAMP), DATE '1900-01-01') DESC
        ) = 1
       )
     , cte_po_context
    AS (
SELECT SPLIT_PART(c.supplier_code, '-', 1)                                     nrm_company_name
     , CASE
           WHEN UPPER(TRIM(c.company_shortcode)) = 'FA'
            AND TRIM(c.order_number)       = 'FAPO200025244'
            AND TRIM(c.supplier_code)     = 'FA-STE096'
           THEN 'STE090'
           ELSE SPLIT_PART(c.supplier_code, '-', 2)
       END                                                                        nrm_supplier_code
     , c.company_shortcode
     , c.order_number
     , c.line_number
     , c.po_position
     , c.supplier_name                                                           po_supplier_name
     , c.line_description                                                        orderitemdescription
     , c.accountingelement1
     , c.accountingelement2
     , c.accountingelement3
     , c.accountingelement4
     , c.net_value_gbp                                                           split_net_value_gbp
     , c.net_value_gbp_text                                                      split_net_value_gbp_text
     , c.unit_value                                                              po_unit_value
     , c.accepted_quantity_total                                                 po_accepted_quantity_total
     , c.fully_invoiced_quantity                                                 po_fully_invoiced_quantity
     , SUM(c.net_value_gbp) OVER (
           PARTITION BY c.company_shortcode, c.order_number, c.line_number
       )                                                                         line_total_net_value_gbp
  FROM src_fin_supplier_po_commitment_stack_base                                  c
       )
     , cte_joined
    AS (
SELECT r.*
     , c.nrm_company_name
     , c.nrm_supplier_code
     , c.po_supplier_name
     , c.orderitemdescription
     , r.po_position                                                             split_position
     , UPPER(TRIM(c.accountingelement1))                                         nrm_accountingelement1
     , UPPER(TRIM(c.accountingelement2))                                         nrm_accountingelement2
     , UPPER(TRIM(c.accountingelement3))                                         nrm_accountingelement3
     , UPPER(TRIM(c.accountingelement4))                                         nrm_accountingelement4
     , c.split_net_value_gbp
     , c.split_net_value_gbp_text
     , c.line_total_net_value_gbp
     , c.po_unit_value
     , c.po_accepted_quantity_total
     , c.po_fully_invoiced_quantity
     , CASE UPPER(TRIM(r.company_shortcode))
           WHEN 'FA'   THEN 'THE FA GROUP - FA'
           WHEN 'NFC'  THEN 'THE FA GROUP - NFC'
           WHEN 'WNSL' THEN 'THE FA GROUP - WNSL'
           ELSE 'UNKNOWN'
       END                                                                       nrm_company_id
     , CASE
          WHEN TRY_CAST(NULLIF(r.order_date_excel, '') AS TIMESTAMP) IS NOT NULL
          THEN CAST(TRY_CAST(NULLIF(r.order_date_excel, '') AS TIMESTAMP) AS DATE)
          WHEN TRY_CAST(NULLIF(r.order_date_excel, '') AS DOUBLE) IS NOT NULL
          THEN DATE '1899-12-30' + CAST(FLOOR(TRY_CAST(NULLIF(r.order_date_excel, '') AS DOUBLE)) AS INTEGER)
          ELSE NULL
       END                                                                       order_date
     , CASE
          WHEN TRY_CAST(NULLIF(r.receipt_date_excel, '') AS TIMESTAMP) IS NOT NULL
          THEN CAST(TRY_CAST(NULLIF(r.receipt_date_excel, '') AS TIMESTAMP) AS DATE)
          WHEN TRY_CAST(NULLIF(r.receipt_date_excel, '') AS DOUBLE) IS NOT NULL
          THEN DATE '1899-12-30' + CAST(FLOOR(TRY_CAST(NULLIF(r.receipt_date_excel, '') AS DOUBLE)) AS INTEGER)
          ELSE NULL
       END                                                                       receipt_date
     , CASE
          WHEN TRY_CAST(NULLIF(r.receipt_created_date_excel, '') AS TIMESTAMP) IS NOT NULL
          THEN CAST(TRY_CAST(NULLIF(r.receipt_created_date_excel, '') AS TIMESTAMP) AS DATE)
          WHEN TRY_CAST(NULLIF(r.receipt_created_date_excel, '') AS DOUBLE) IS NOT NULL
          THEN DATE '1899-12-30' + CAST(FLOOR(TRY_CAST(NULLIF(r.receipt_created_date_excel, '') AS DOUBLE)) AS INTEGER)
          ELSE NULL
       END                                                                       receipt_created_date
     -- Receipt value bases — both LINE-LEVEL (per order+line+pos), i.e. identical on
     -- every receipt event of a line (nominal/invoice/qty are line totals repeated).
     -- receipted_quantity is NOT used: it is a unit count on ~56% of rows (see BLOCKER-1).
     , GREATEST(COALESCE(r.net_value_original_curr, 0) - COALESCE(r.invoice_value, 0), 0)   nrm_receipt_basis_a_gbp   -- A: nominal - invoice, both in TRANSACTION currency (net_value_original_curr, invoice_value). No *_gbp fallback: falling back to the GBP-converted column would mix currencies. net_value_original_curr is non-null on all rows.
     , GREATEST((COALESCE(c.po_accepted_quantity_total, 0) - COALESCE(c.po_fully_invoiced_quantity, 0))
              * COALESCE(c.po_unit_value, 0), 0)                                                              nrm_receipt_basis_b_gbp   -- B: (accepted - fully_invoiced) x unit_value = received-not-invoiced GRNI (ACTIVE)
     , r.currency                                                                nrm_currency
  FROM cte_src_receipt_base                                                       r
 INNER JOIN
       cte_po_context                                                             c
    ON c.company_shortcode = r.company_shortcode
   AND c.order_number      = r.order_number
   AND c.line_number       = r.line_number
   AND c.po_position       = r.po_position
 WHERE UPPER(TRIM(c.nrm_company_name)) IN ('FA', 'NFC', 'WNSL')
       )
     , cte_valued
    AS (
SELECT t.*
     -- Cap = LEAST(A, B), both now in LOCAL/transaction currency, so currency-consistent
     -- for every currency (basis A uses net_value_original_curr; basis B uses unit_value):
     --   A : nominal − invoice          (line's remaining-to-invoice)
     --   B : (accepted − fully_invoiced) × unit  = received-not-invoiced GRNI (= PO hard, local)
     -- The old split_net cap is dropped: it never bound (0 rows) and was GBP-based, so it
     -- risked re-introducing a cross-currency artefact. (FAPO200067694 EUR now = 9,501, not 3,394.69.)
     , LEAST(
           COALESCE(t.nrm_receipt_basis_a_gbp, t.nrm_receipt_basis_b_gbp)
         , COALESCE(t.nrm_receipt_basis_b_gbp, t.nrm_receipt_basis_a_gbp)
       )                                                                             nrm_receipt_line_value_gbp
     -- The value is line-level; assign it to ONE representative receipt event (latest)
     -- so repeated receipts against the same line do not multiply the line value.
     , ROW_NUMBER() OVER (
           PARTITION BY t.company_shortcode, t.order_number, t.line_number, t.split_position
               ORDER BY t.receipt_created_date DESC NULLS LAST
                      , t.receipt_number       DESC
       )                                                                         nrm_line_pos_rnk
  FROM cte_joined                                                                 t
       )
     , cte_net
    AS (
SELECT t.*
     , CASE
           WHEN t.nrm_line_pos_rnk = 1 THEN t.nrm_receipt_line_value_gbp
           ELSE 0
       END                                                                        nrm_receipt_net_gbp
  FROM cte_valued                                                                 t
       )
     , cte_supplier_mapped
    AS (
SELECT t.*
     , COALESCE(s_exact.supplier_id, s_bare.supplier_id)                         wd_supplier_id
     , COALESCE(s_exact.nrm_supplier_name, s_bare.nrm_supplier_name)             nrm_supplier_name
     , COALESCE(s_exact.nrm_agg_supplier_number, s_bare.nrm_agg_supplier_number) nrm_agg_supplier_number
     , COALESCE(s_exact.nrm_email_exists_ind, s_bare.nrm_email_exists_ind, 0)    nrm_email_exists_ind
     , CASE
           WHEN UPPER(TRIM(COALESCE(t.nrm_currency, ''))) = 'JPY'
           THEN ROUND(t.nrm_receipt_net_gbp, 0)
           ELSE ROUND(t.nrm_receipt_net_gbp, 2)
       END                                                                        nrm_receipt_export_amount
     , CASE
           WHEN UPPER(TRIM(COALESCE(t.nrm_currency, ''))) = 'JPY'
           THEN CAST(CAST(ROUND(t.nrm_receipt_net_gbp, 0) AS BIGINT) AS VARCHAR)
           ELSE CAST(CAST(ROUND(t.nrm_receipt_net_gbp, 2) AS DECIMAL(18,2)) AS VARCHAR)
       END                                                                        nrm_receipt_export_amount_text
     , strftime(t.order_date,             '%d/%m/%Y')                            nrm_order_date
     , strftime(t.receipt_date,           '%d/%m/%Y')                            nrm_receipt_date
     , strftime(t.receipt_created_date,   '%d/%m/%Y')                            nrm_receipt_created_date
     , CASE
           WHEN COALESCE(s_exact.supplier_id, s_bare.supplier_id) IS NULL
           THEN 'DROP_NO_SUPPLIER'
           ELSE 'KEEP'
       END                                                                        nrm_row_status
  FROM cte_net                                                                    t
       LEFT OUTER JOIN
       src_fin_workday_supplier_map                                               s_exact
    ON s_exact.reference_id_token = TRIM(t.nrm_supplier_code) || '-' || UPPER(TRIM(t.nrm_company_name))
       LEFT OUTER JOIN
       src_fin_workday_supplier_map                                               s_bare
    ON s_bare.reference_id_token  = TRIM(t.nrm_supplier_code)
       )
SELECT t.*
     , t.receipt_line_key
    || '|SPLIT|' || t.split_position                                              nrm_receipt_split_line_key
     , 'HARD'                                                                     commitment_class
  FROM cte_supplier_mapped                                                        t
;
