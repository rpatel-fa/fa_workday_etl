DROP TABLE IF EXISTS src_crm_customer_wd_manual_missing
;
CREATE TABLE src_crm_customer_wd_manual_missing
    AS
select wc.customer_id, wc.reference_id, wc.customer
     , workday_id, wmc.licence_ref_no, wmc.account
  from crm_workday_customer_manual  wmc
       left outer join 
       fin_workday_extract_customer wc
         on wc.customer_id = wmc.workday_id
;