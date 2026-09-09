DROP TABLE IF EXISTS src_fin_customer_invoice_nrm;

CREATE TABLE src_fin_customer_invoice_nrm
    AS
  WITH cte_src_customer_invoice_base
    AS (
SELECT t.* 
  FROM src_fin_customer_invoice_base         t
       )
SELECT t.*
     , COALESCE(c_exact.customer_id, c_bare.customer_id)                         wd_customer_id
     , COALESCE(c_exact.nrm_customer_name, c_bare.nrm_customer_name)             wd_customer_name
     , COALESCE(c_exact.nrm_agg_customer_number, c_bare.nrm_agg_customer_number) wd_customer_reference_id
     , strftime(CAST(
       COALESCE(
          TRY_STRPTIME(NULLIF(TRIM(t.document_date), ''), '%Y-%m-%d %H:%M:%S')
        , TRY_STRPTIME(NULLIF(TRIM(t.document_date), ''), '%Y-%m-%d')
        , TRY_STRPTIME(NULLIF(TRIM(t.document_date), ''), '%d-%m-%Y')
       ) AS DATE), '%d/%m/%Y')               nrm_document_date
     , CASE (UPPER(TRIM(t.business_unit)))
           WHEN 'FA'   THEN 'THE FA GROUP - FA'
           WHEN 'NFC'  THEN 'THE FA GROUP - NFC'
           WHEN 'WNSL' THEN 'THE FA GROUP - WNSL'
           ELSE 'UNKNOWN'
       END                                   nrm_company_id
     /* ----------------------------------------------------------------------
        nrm_transaction_amt — VARCHAR pass-through pattern  (for the 1,439 GBP
        + 8 USD rows loaded in PRODUCTION on 10/08/26 this is correct and
        equivalent to the DOUBLE-cast pattern used in
        src_fin_supplier_invoice_nrm.sql; see comparison block below).

        BRANCHES TODAY:
          GBP  → COALESCE(current_transaction_amount_gbp, '0')  (VARCHAR)
          JPY  → CAST(ROUND(CAST(src AS DOUBLE), 0) AS VARCHAR)  (VARCHAR, int)
          ELSE → COALESCE(current_transaction_amount_in_originating_currncy, '0')
                                                                        (VARCHAR)

        This differs from supplier_invoice_nrm, which CASTs every branch
        AS DOUBLE up front (and TRIMs + NULLIF on the source first).  The
        downstream publish SQL (`workday_customer_invoice.sql`) writes
        this column into column_ha / column_mj AS-IS; DuckDB + Workday's
        EIB loader both treat numeric-looking VARCHARs in those cells as
        numbers, so there is no correctness difference on the current
        dataset.

        CONTEXT:  this is a PURE ONE-OFF OPENING-BALANCES MIGRATION.
                 Customer invoices are loaded into Workday ONCE at go-live
                 (AR / AP Conversion / Open AR Conversion).  The source
                 data is statically extracted on one date and is not
                 refreshed; there is no on-going BAU pipeline here.
                 Accordingly the SQL below was written as a one-time
                 transform, not as a defensive BAU/recurring transform.
                 There is therefore NO value today in rewriting the
                 VARCHAR pattern below to match supplier_invoice_nrm —
                 the code runs once, will not be re-run against new data,
                 and the output has already been signed off.

        WHY LEAVE AS VARCHAR TODAY (do not change without re-running the
        full customer pipeline + signing off):
          • 0 irregular strings in 1,447 rows (no leading/trailing ws, no
            embedded commas, no letters, no NULL-empty strings).  Verified
            by probe 10/08/26 — all customer amounts parse cleanly.
          • Customer currencies currently in play: GBP 1,439 / USD 8  →
            neither branch demands special rounding; JPY branch is dead
            code (0 rows) at the moment.
          • This is a ONE-OFF migration — not a recurring pipeline, so
            defensive casting/handling against *future* source rows is
            unnecessary; the 1,447 rows this SQL produces are the ONLY
            rows it will ever produce.
          • Gate at src_fin_customer_invoice.sql already re-casts via
            CAST(nrm_transaction_amt AS DOUBLE) and drops 9 rows with
            amounts that evaluate to 0; that cast-then-compare on 1,447
            rows × every validation run is the real cost of keeping this
            as VARCHAR — negligible in practice, and also a one-off.

        IF WE EVER WANTED TO ALIGN to the supplier-invoice pattern (for
        instance if the AR migration was ever re-extracted and re-run, or
        if a customer three-way recon was built that does DOUBLE compares
        with 0.005 tolerance), then replace THIS nrm_transaction_amt CASE
        expression verbatim with:

          , CASE
              WHEN UPPER(TRIM(t.currency_id)) = 'GBP'
              THEN CAST(COALESCE(NULLIF(TRIM(t.current_transaction_amount_gbp), ''), '0') AS DOUBLE)
              WHEN UPPER(TRIM(t.currency_id)) = 'JPY'
              THEN ROUND(CAST(COALESCE(NULLIF(TRIM(t.current_transaction_amount_in_originating_currncy), ''), '0') AS DOUBLE), 0)
              ELSE CAST(COALESCE(NULLIF(TRIM(t.current_transaction_amount_in_originating_currncy), ''), '0') AS DOUBLE)
            END                               nrm_transaction_amt

        AND in src_fin_customer_invoice.sql line 10 simplify the gate from:
              AND COALESCE(CAST(t.nrm_transaction_amt AS DOUBLE), 0) <> 0
        to:
              AND COALESCE(t.nrm_transaction_amt, 0) <> 0

        ROW-BY-ROW TEST RESULT (10/08/26, in-memory run against the 1,447
        PRODUCTION rows):  applying the supplier-style cast produces
        nrm_transaction_amt values that are pairwise DOUBLE-equal to
        CAST(current_VARCHAR AS DOUBLE) on EVERY row  (0 mismatches,
        0 rows dropped, identical grand total  £125,649,286.27).  The
        change is therefore a no-op on today's data — safe, but pointless
        given this is a one-time signed-off migration.
        ---------------------------------------------------------------------- */
     , CASE
         WHEN UPPER(TRIM(t.currency_id))     = 'GBP'
         THEN COALESCE(t.current_transaction_amount_gbp, '0')
         WHEN UPPER(TRIM(t.currency_id))     = 'JPY'
         THEN CAST(ROUND(CAST(COALESCE(t.current_transaction_amount_in_originating_currncy, '0') AS DOUBLE), 0) AS VARCHAR)
         ELSE COALESCE(t.current_transaction_amount_in_originating_currncy, '0')
       END                                   nrm_transaction_amt
  FROM cte_src_customer_invoice_base         t
       LEFT OUTER JOIN
       src_fin_workday_customer_map                                             c_exact
           ON c_exact.reference_id_token     = TRIM(t.customer_id) || '-' || TRIM(t.business_unit)
       LEFT OUTER JOIN
       src_fin_workday_customer_map                                             c_bare
           ON c_bare.reference_id_token      = TRIM(t.customer_id)
;
