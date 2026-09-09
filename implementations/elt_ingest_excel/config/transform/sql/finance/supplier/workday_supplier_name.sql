DROP TABLE IF EXISTS workday_supplier_name
;
CREATE TABLE workday_supplier_name 
    AS
SELECT
       t.supplier_id                   supplier_id
     , t.nrm_supplier_name             supplier_name
     , COALESCE(
           t.nrm_resolved_agg_supplier_number
         , t.nrm_agg_supplier_number
       )                               reference_id
     , t.nrm_supplier_category         supplier_category
     , CAST(NULL AS VARCHAR)           business_entity_name
     , CAST(NULL AS VARCHAR)           do_not_reimburse_contingent_worker_expense_reports
     , CAST(NULL AS VARCHAR)           worktag_only
     , CAST(NULL AS VARCHAR)           integration_system
     , CAST(NULL AS VARCHAR)           external_entity_id
     , CAST(NULL AS VARCHAR)           supplier_source
     , CAST(NULL AS VARCHAR)           supplier_change_source
     , CAST(NULL AS VARCHAR)           submit
     , CAST(NULL AS VARCHAR)           spend_category_or_hierarchy_plus
     , CAST(NULL AS VARCHAR)           supplier_security_segment
     , CAST(NULL AS VARCHAR)           create_supplier_from_customer
     , CAST(NULL AS VARCHAR)           create_supplier_from_financial_institution
     , CAST(NULL AS VARCHAR)           create_supplier_from_tax_authority
     , CAST(NULL AS VARCHAR)           create_supplier_from_investor
     , CAST(NULL AS VARCHAR)           customer_account_number
     , CAST(NULL AS VARCHAR)           certificate_of_insurance_date
     , CAST(NULL AS VARCHAR)           duns_number
     , CAST(NULL AS VARCHAR)           unique_entity_identifier
  FROM src_fin_supplier                t
 ORDER BY
       supplier_id
;
