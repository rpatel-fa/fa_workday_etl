DROP TABLE IF EXISTS src_crm_customer_base
;
CREATE TABLE src_crm_customer_base
    AS
WITH cte_ca
  AS (
SELECT 'ACCOUNT'                       crm_customer_type
     , ca.account_number_id            crm_customer_id
     , ca.licencerefno                 crm_licenceref_no
     , ca.fan_primary_contact_contact  crm_fan
     , ca.created_on                   crm_created_on
     , ca.modified_on                  crm_modified_on
     , (UPPER(TRIM(ca.status)) = 'ACTIVE')  crm_is_active
  FROM crm_account                     ca 
 WHERE NULLIF(NULLIF(TRIM(ca.licencerefno), ''), 'NULL') IS NOT NULL
     )
     , cte_cc
  AS (
SELECT 'INDIVIDUAL'                    crm_customer_type
     , cc.fan                          crm_customer_id
     , cc.licencesref                  crm_licenceref_no
     , cc.fan                          crm_fan
     , cc.created_on                   crm_created_on
     , cc.modified_on                  crm_modified_on
     , (UPPER(TRIM(cc.system_status)) = 'ACTIVE')  crm_is_active
  FROM crm_contact                     cc 
 WHERE NULLIF(NULLIF(TRIM(cc.licencesref), ''), 'NULL') IS NOT NULL
     )
     , cte_crm
  AS (
SELECT * FROM cte_ca
UNION ALL
SELECT * FROM cte_cc
     )
     -- The extracts repeat some ids across rows; one value per CRM record keeps the downstream DISTINCTs one row per record.
     , cte_crm_dt
  AS (
SELECT * EXCLUDE (crm_created_on, crm_modified_on, crm_is_active)
     , MAX(crm_created_on)  OVER (PARTITION BY crm_customer_type, crm_customer_id)  crm_created_on
     , MAX(crm_modified_on) OVER (PARTITION BY crm_customer_type, crm_customer_id)  crm_modified_on
     , BOOL_OR(crm_is_active) OVER (PARTITION BY crm_customer_type, crm_customer_id)  crm_is_active
  FROM cte_crm
     )
SELECT *
     , TRIM(UNNEST(STRING_SPLIT(crm_licenceref_no, '|')))  crm_unnest_reference_id
  FROM cte_crm_dt
;
