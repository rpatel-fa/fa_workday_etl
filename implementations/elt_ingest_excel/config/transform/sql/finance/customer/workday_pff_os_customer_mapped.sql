DROP TABLE IF EXISTS workday_pff_os_customer_mapped
;
CREATE TABLE workday_pff_os_customer_mapped
    AS
  WITH cte_gp 
  AS (
SELECT distinct 
       customer_number
     , moasorg_id
     , workday_id 
  from pff_gp_discipline_customer
 WHERE NULLIF(TRIM(workday_id), '')    IS NOT NULL 
     )
   , cte_os
  AS (
SELECT c.club_id                       os_club_id
     , gp.workday_id                   workday_id
  FROM pff_pro_club                    c
       INNER JOIN 
       cte_gp                          gp
          ON gp.moasorg_id             = c.onsidecludorgid
     )
   , cte_mm
  AS (
select replace(new_workday_id, 'OS-', '') os_club_id
     , workday_id
  from pff_workday_manual_mapped
 where workday_id NOT IN (SELECT workday_id FROM cte_os)
     )
   , cte_wd_os
  AS (
SELECT * FROM cte_os
UNION
SELECT * FROM cte_mm
     )
   , cte_oswd
  as (
select 
       wo.workday_id                mapped_workday_id
     , 'OS-' || c.club_id           os_club_id
     , c.onsidecludorgid            os_onside_org_id
     , c.club_name                  os_club_name
     , c.club_type                  os_club_type
     , c.club_subtype               os_club_subtype
     , c.county                     os_county
     , c.contact_lname              os_contact_lname
     , c.addressline1               os_addressline1
     , c.addressline2               os_addressline2
     , c.addressline3               os_addressline3
     , c.town                       os_town
     , c.postcode                   os_postcode
     , c.country                    os_country
     , c.club_secretary_fan         os_club_secretary_fan
     , c.firstname                  os_firstname
     , c.lastname                   os_lastname
     , c.email                      os_email
     , w.customer                   wd_customer
     , w.reference_id 				      wd_reference_id
     , w.customer_id                wd_customer_id
     , w.customer_category          wd_customer_category
     , w.default_address            wd_default_address
     , w.default_phone_number       wd_default_phone_number
     , w.customer_contacts          wd_customer_contacts
     , w.customer_status            wd_customer_status
  from pff_pro_club                 c
       INNER JOIN
       cte_wd_os                    wo
         ON wo.os_club_id           = c.club_id
       INNER JOIN
       fin_workday_extract_customer w
         ON w.customer_id           = wo.workday_id
     )
select * 
  from cte_oswd     o 
 order by 1 
;

