DROP TABLE IF EXISTS src_crm_customer_wd_manual
;
CREATE TABLE src_crm_customer_manual
    AS
WITH cte_wd_mm 
  AS (
SELECT m.licence_ref_no               wd_licence_ref_no
     , m.fan                          wd_fan
     , m.contact                      wd_contact
     , m.account                      wd_account
     , m.workday_id                   wd_customer_id
     , c.*
  FROM 
       crm_workday_customer_manual    m
       LEFT OUTER JOIN
       src_crm_customer_base          c
         ON c.crm_licenceref_no       = m.licence_ref_no
     )
SELECT *
  FROM cte_wd_mm
;
