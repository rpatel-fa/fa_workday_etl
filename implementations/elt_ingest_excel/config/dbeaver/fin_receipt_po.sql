-- DBeaver bookmarks: Supplier Purchase Order RECEIPTS-only (WORKDAY via Commitment Stack GRNI)
-- Receipt EIB = GRNI accrual (Hard commitment only) sourced from the staged PO subset
-- (raw hard_commitment_value_gbp > 0.005). The old Proactis receipt extract table
-- (fin_supplier_po_receipt) is NOT USED by any output SQL.

-- 1. GRNI totals (commitment stack source vs receipt EIB = 100% tie expected)
SELECT 'Staged PO Hard rows (raw hard_gbp > 0.005)'  bucket,
       count(*) n,
       round(sum(COALESCE(hard_commitment_value_gbp,0)),2) total_gbp
  FROM src_fin_supplier_po_commitment_stack
 WHERE COALESCE(hard_commitment_value_gbp,0) > 0.005;
SELECT 'Receipt EIB rows (actual)'                    bucket,
       count(*) n,
       0 total_gbp  -- GBP not in receipt EIB cols (local only); derive via PO join below
  FROM workday_supplier_po_commitment_stack_receipt;

-- 2. Receipt EIB line-by-line (shows synthetic REC-{order_number} doc, PO link, extended amount)
SELECT r.column_p   company
     , r.column_d   receipt_number        -- synthetic REC-{order_number}
     , r.column_y   document_date         -- proxy = commitment_date
     , r.column_bu  purchase_order_line   -- = PO EIB column_ez
     , r.column_ay  currency
     , r.column_bz  extended_amount_local -- = Hard local commitment value
     , r.column_j   submit_flag           -- always Y, all lines
     , r.column_c   add_only_flag         -- Y on header row (rnk=1 per doc)
     , r.column_e   auto_complete_flag    -- Y on header row (rnk=1 per doc)
     , r.column_f   comment_source        -- "Source: Commitment Stack (GRNI)"
  FROM workday_supplier_po_commitment_stack_receipt r
 ORDER BY r.column_p, r.column_d, r.column_bu;

-- 3. Receipt → PO linkage verification (0-row result = PASS)
WITH check_ref AS (
  SELECT r.column_bu dangling_ref, r.column_d receipt_num, r.column_bz amount
    FROM workday_supplier_po_commitment_stack_receipt r
   WHERE NOT EXISTS (
         SELECT 1 FROM workday_supplier_po_commitment_stack p
          WHERE p.column_ez = r.column_bu)
),
check_amt AS (
  SELECT p.column_ez po_line, p.column_gf po_gross_local, r.column_bz rcpt_hard_local
    FROM workday_supplier_po_commitment_stack p
    JOIN workday_supplier_po_commitment_stack_receipt r
      ON p.column_ez = r.column_bu
   WHERE TRY_CAST(p.column_gf AS DECIMAL(18,2))
       < TRY_CAST(r.column_bz AS DECIMAL(18,2))
)
SELECT 'Dangling receipt refs (expect 0)'  check_nm, count(*) n FROM check_ref
 UNION ALL
SELECT 'Over-receipt splits (expect 0)',     count(*)   FROM check_amt;

-- 4. Per-company receipt aggregates (transaction currency; GBP equivalence via staged PO)
SELECT COALESCE(p.nrm_company_id, r.column_p)            company
     , COALESCE(p.currency, r.column_ay)                 currency
     , count(*)                                          rcpt_lines
     , round(sum(TRY_CAST(r.column_bz AS DECIMAL(18,2))),2) rcpt_local_total
     , round(sum(COALESCE(p.nrm_hard_remaining_gbp, 0)),2) rcpt_gbp_total  -- tie target
  FROM workday_supplier_po_commitment_stack_receipt r
       LEFT JOIN
       src_fin_supplier_po_commitment_stack p
       ON p.po_line_pos_id = r.column_bu
 GROUP BY 1,2
 ORDER BY 1,2;

-- 5. Receipt documents with header flags (one header row per doc)
SELECT company, receipt_number, receipt_lines, header_row_id
     , document_date
     , submit_flag, add_only_flag, auto_complete_flag
     , round(sum_receipt_hard_local,2) sum_receipt_hard_local
  FROM (
       SELECT r.column_p                                company
            , r.column_d                                receipt_number
            , count(*)                                  receipt_lines
            , min(r.column_bu)                          header_row_id
            , min(r.column_y)                            document_date
            , min(CASE WHEN r.column_c = 'Y' THEN 'Y' ELSE 'N' END) add_only_flag
            , min(CASE WHEN r.column_e = 'Y' THEN 'Y' ELSE 'N' END) auto_complete_flag
            , min(r.column_j)                            submit_flag
            , sum(TRY_CAST(r.column_bz AS DECIMAL(18,2))) sum_receipt_hard_local
         FROM workday_supplier_po_commitment_stack_receipt r
        GROUP BY 1,2
       ) t
 ORDER BY company, receipt_number;

-- 6. Top 10 largest receipt documents (for eyeballing currency anomalies — JPY should be integer)
SELECT r.column_p company
     , r.column_d receipt_number
     , count(DISTINCT r.column_bu) n_po_lines
     , round(sum(TRY_CAST(r.column_bz AS DECIMAL(18,2))),2) receipt_total_local
     , min(r.column_ay) currency
  FROM workday_supplier_po_commitment_stack_receipt r
 GROUP BY 1,2
 ORDER BY 4 DESC
 LIMIT 10;

-- 7. Unresolved suppliers reconcile (PO + receipt share the unresolved set because
-- receipt is a subset of staged PO. If a PO has unmapped supplier, its receipt
-- rows also carry NULL supplier_ref into Workday — flagged in validation issues.)
SELECT u.*
  FROM reconcile_supplier_po_unresolved u
 ORDER BY u.total_commitment_gbp DESC
 LIMIT 50;
