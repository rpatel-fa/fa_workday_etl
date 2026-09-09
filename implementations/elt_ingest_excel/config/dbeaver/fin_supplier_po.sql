-- DBeaver bookmarks: Supplier Purchase Order + Receipt (WORKDAY via Commitment Stack)
-- Current architecture: BOTH EIBs are derived from the single Commitment Stack report.
-- The old Proactis PO Receipts Extract is ingested but NO LONGER USED in any output SQL.

-- 1. Issue/summary counts
SELECT count(*) FROM fin_ref_wd_fdm_mapping;
SELECT * FROM validation_supplier_purchase_order_receipt_issues ORDER BY rule_type, recon_key;
SELECT rule_type, count(*) AS n, sum(CASE WHEN detail ~ '^[0-9.-]+$' THEN CAST(REPLACE(detail,',','') AS DOUBLE) ELSE 0 END) AS approx_impact
  FROM validation_supplier_purchase_order_receipt_issues
 GROUP BY 1 ORDER BY 1;

-- 2. Commitment Stack PO source chain (base → nrm → staged → ref → EIB)
SELECT 'PO base_rows'     stage, count(*) n FROM src_fin_supplier_po_commitment_stack_base;
SELECT 'PO nrm_rows'      stage, count(*) n FROM src_fin_supplier_po_commitment_stack_nrm;
SELECT 'PO staged_rows'   stage, count(*) n FROM src_fin_supplier_po_commitment_stack;
SELECT 'PO ref_rows'      stage, count(*) n FROM src_fin_supplier_po_commitment_stack_ref;
SELECT 'PO eib_rows'      stage, count(*) n FROM workday_supplier_po_commitment_stack;

SELECT * FROM src_fin_supplier_po_commitment_stack_base ORDER BY order_number, line_number, po_position LIMIT 200;
SELECT * FROM src_fin_supplier_po_commitment_stack_nrm  ORDER BY order_number, line_number, po_position LIMIT 200;
SELECT * FROM src_fin_supplier_po_commitment_stack      ORDER BY order_number, line_number, po_position LIMIT 200;
SELECT * FROM src_fin_supplier_po_commitment_stack_ref  ORDER BY order_number, line_number, po_position LIMIT 200;
SELECT * FROM workday_supplier_po_commitment_stack      ORDER BY column_p, column_j, column_ez;

-- 3. Commitment Stack RECEIPT EIB (GRNI = Hard commitment, staged PO subset)
SELECT 'Receipt eib_rows' stage, count(*) n FROM workday_supplier_po_commitment_stack_receipt;
SELECT count(DISTINCT column_d) AS distinct_receipt_docs
  FROM workday_supplier_po_commitment_stack_receipt;
SELECT * FROM workday_supplier_po_commitment_stack_receipt ORDER BY column_d, column_bu;

-- 4. Cross-EIB reference check (must be 0 dangling, 0 over-receipt)
SELECT 'Dangling receipt refs (expect 0)' check_nm, count(*) n
  FROM workday_supplier_po_commitment_stack_receipt r
 WHERE NOT EXISTS (
       SELECT 1 FROM workday_supplier_po_commitment_stack p
        WHERE p.column_ez = r.column_bu);

SELECT 'Over-receipt splits (expect 0)' check_nm, count(*) n
  FROM workday_supplier_po_commitment_stack p
  JOIN workday_supplier_po_commitment_stack_receipt r
    ON p.column_ez = r.column_bu
 WHERE TRY_CAST(p.column_gf AS DECIMAL(18,2))
     < TRY_CAST(r.column_bz AS DECIMAL(18,2));

-- 5. Reconcile reports (row funnel, amount-by-currency, unmapped suppliers)
SELECT * FROM reconcile_cs_row_counts                   ORDER BY company;
SELECT * FROM reconcile_cs_amount_by_currency            ORDER BY company, currency;
SELECT * FROM reconcile_cs_currency_coverage             ORDER BY present_in DESC, currency;
SELECT * FROM reconcile_cs_unmapped_suppliers            ORDER BY total_commitment_gbp DESC;
SELECT * FROM reconcile_cs_report                        ORDER BY 1, 2;
SELECT * FROM reconcile_cs_board_summary                 ORDER BY sort_order;

-- 6. Receipt EIB header flags (add_only Y / auto_complete Y / distinct docs = same count)
SELECT column_flag, count(*) FROM (
  SELECT 'add_only_Y=header_rows'        column_flag FROM workday_supplier_po_commitment_stack_receipt WHERE column_c = 'Y'
  UNION ALL
  SELECT 'auto_complete_Y=header_rows'                FROM workday_supplier_po_commitment_stack_receipt WHERE column_e = 'Y'
  UNION ALL
  SELECT 'submit_Y=all_lines'                         FROM workday_supplier_po_commitment_stack_receipt WHERE column_j = 'Y'
  UNION ALL
  SELECT 'distinct_receipt_docs'                      FROM workday_supplier_po_commitment_stack_receipt GROUP BY column_d
) t GROUP BY 1 ORDER BY 1;

-- 7. Currency profile on EIBs (JPY → BIGINT, 2dp everywhere else)
SELECT column_ah currency, count(*) eib_lines,
       round(sum(TRY_CAST(column_gf AS DECIMAL(18,2))),2) eib_total_transaction_currency
  FROM workday_supplier_po_commitment_stack
 GROUP BY 1 ORDER BY 2 DESC;
SELECT column_ay currency, count(*) rcpt_lines,
       round(sum(TRY_CAST(column_bz AS DECIMAL(18,2))),2) rcpt_total_transaction_currency
  FROM workday_supplier_po_commitment_stack_receipt
 GROUP BY 1 ORDER BY 2 DESC;

-- 8. PO split distribution (how many positions per line, sanity check)
SELECT n_positions_per_line, count(*) AS po_lines
  FROM (
       SELECT order_number, line_number,
              count(DISTINCT po_position) n_positions_per_line
         FROM src_fin_supplier_po_commitment_stack
        GROUP BY 1,2
       ) t
 GROUP BY 1 ORDER BY 1;
