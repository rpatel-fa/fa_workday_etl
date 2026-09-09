-- ============================================================================
-- RECONCILE: Supplier bank/settlement account proof (row-level)
-- Supplier-only - customer has no active bank aggregation (workday_customer_
-- settlement_account.sql is permanently disabled via a hardcoded `AND 1 = 2`).
--
-- One row per distinct identifiable bank account (COALESCE(bank_account_number,
-- bank_iban), with a bank_name) captured in the aggregated nrm_array_bank
-- (src_fin_supplier), showing whether it is present in
-- workday_supplier_settlement_account for the same supplier_id. The same
-- identifiability filter used by workday_supplier_settlement_account.sql
-- itself (bank_account_number or iban present AND bank_name present) is
-- applied to the source side, so bank stubs with no usable identifier are
-- not counted as false gaps.
-- ============================================================================

DROP TABLE IF EXISTS reconcile_supplier_bank_proof
;
CREATE TABLE reconcile_supplier_bank_proof
    AS
  WITH cte_source_bank
    AS (
SELECT DISTINCT
       s.supplier_id                                                                            supplier_id
     , u.bank.nrm_bank_name                                                                     bank_name
     , COALESCE(u.bank.nrm_bank_account_number, u.bank.nrm_bank_iban)                            bank_identifier
  FROM src_fin_supplier                                                                         s
     , UNNEST(s.nrm_array_bank)                                                                 u(bank)
 WHERE COALESCE(u.bank.nrm_bank_account_number, u.bank.nrm_bank_iban)                            IS NOT NULL
   AND NULLIF(TRIM(u.bank.nrm_bank_name), '')                                                    IS NOT NULL
       )
     , cte_target_bank
    AS (
SELECT DISTINCT
       supplier_id
     , COALESCE(bank_account_number, bank_iban)                                                 bank_identifier
  FROM workday_supplier_settlement_account
       )
SELECT sb.supplier_id                                                                           supplier_id
     , sb.bank_name                                                                              bank_name
     , sb.bank_identifier                                                                       bank_identifier
     , CASE WHEN tb.supplier_id IS NOT NULL THEN 'Y' ELSE 'N' END                                found_in_workday_supplier_settlement_account
  FROM cte_source_bank                                                                          sb
  LEFT JOIN cte_target_bank                                                                     tb
         ON tb.supplier_id                                                                      = sb.supplier_id
        AND tb.bank_identifier                                                                  = sb.bank_identifier
 ORDER BY found_in_workday_supplier_settlement_account
     , supplier_id
;
