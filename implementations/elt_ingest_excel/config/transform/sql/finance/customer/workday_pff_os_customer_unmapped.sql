DROP TABLE IF EXISTS workday_pff_os_customer_unmapped
;
CREATE TABLE workday_pff_os_customer_unmapped
    AS
WITH cte_gp 
    AS
     (
SELECT DISTINCT 
       customer_number
     , moasorg_id
     , workday_id
  FROM pff_gp_discipline_customer
 WHERE NULLIF(TRIM(workday_id), '') IS NOT NULL 
     )
     , cte_os_wd AS
     (
SELECT workday_id manually_mapped_workday_id
     , replace(new_workday_id, 'OS-', '') os_club_id
     , new_workday_id
  from pff_workday_manual_mapped
     )
     , cte_os 
    as
     (
select c.*
     , g.workday_id
     , ow.manually_mapped_workday_id
  from pff_pro_club                 c
       LEFT OUTER JOIN 
       cte_gp                       g
         ON g.moasorg_id            = c.onsidecludorgid
       LEFT OUTER JOIN
       fin_workday_extract_customer w
         ON w.customer_id           = g.workday_id
       LEFT OUTER JOIN
       cte_os_wd                    ow
         ON ow.os_club_id           = c.club_id
     )
select o.club_id                    new_workday_id
     , o.* 
  from cte_os     o 
 where o.workday_id                   IS NULL
   AND o.manually_mapped_workday_id   IS NULL
   and UPPER(TRIM(o.county))          NOT like '%TEST%'
   AND UPPER(TRIM(o.onsidecludorgid)) != 'NULL'
;
