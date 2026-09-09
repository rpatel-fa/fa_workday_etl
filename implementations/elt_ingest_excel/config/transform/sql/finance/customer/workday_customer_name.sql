DROP TABLE IF EXISTS workday_customer_name
;
CREATE TABLE workday_customer_name
    AS
SELECT
       t.customer_id                   customer_id
     , t.nrm_customer_name             customer_name
     , t.nrm_agg_customer_number       reference_id
     , t.nrm_customer_category         customer_category
     , CAST(NULL AS VARCHAR)           business_entity_name
     , CAST(NULL AS VARCHAR)           external_entity_id
     , CAST(NULL AS VARCHAR)           preferred_locale
     , CAST(NULL AS VARCHAR)           preferred_language
     , CAST(NULL AS VARCHAR)           lockbox
     , CAST(NULL AS VARCHAR)           customer_security_segment
     , CAST(NULL AS VARCHAR)           worktag_only
     , CAST(NULL AS VARCHAR)           submit
     , CAST(NULL AS VARCHAR)           create_customer_from_financial_institution
     , CAST(NULL AS VARCHAR)           create_customer_from_supplier
     , CAST(NULL AS VARCHAR)           create_customer_from_tax_authority
     , CAST(NULL AS VARCHAR)           create_customer_from_investor
  FROM src_fin_customer                t
 ORDER BY
       customer_id
;
