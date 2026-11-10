DROP TABLE IF EXISTS workday_pff_os_customer_load
;
CREATE TABLE workday_pff_os_customer_load
    AS
-- Club load sheet for PFF business sign-off: mapped clubs (existing Workday customer id)
-- plus unmapped clubs (no Workday customer id yet, so wd_customer_id is NULL).
WITH cte_all
  AS (
SELECT m.wd_customer_id                wd_customer_id
     , m.os_onside_org_id              pff_id
     , m.os_club_name                  club_name
     , m.os_addressline1               addr_line1
     , m.os_addressline2               addr_line2
     , m.os_addressline3               addr_line3
     , m.os_town                       city
     , m.os_postcode                   postcode2
     , m.os_firstname                  sec_first
     , m.os_lastname                   sec_last
     , m.os_email                      sec_email
     , 'MAPPED'                        load_status
  FROM workday_pff_os_customer_mapped   m
 UNION ALL
SELECT CAST(NULL AS VARCHAR)           wd_customer_id
     , u.onsidecludorgid               pff_id
     , u.club_name                     club_name
     , u.addressline1                  addr_line1
     , u.addressline2                  addr_line2
     , u.addressline3                  addr_line3
     , u.town                          city
     , u.postcode                      postcode2
     , u.firstname                     sec_first
     , u.lastname                      sec_last
     , u.email                         sec_email
     , 'UNMAPPED'                      load_status
  FROM workday_pff_os_customer_unmapped u
     )
     -- PFF holds the literal word NULL (and blanks) for missing values: treat both as NULL.
   , cte_clean
  AS (
SELECT r.wd_customer_id
     , CASE WHEN UPPER(TRIM(r.pff_id)) IN ('', 'NULL') THEN NULL ELSE TRIM(r.pff_id) END pff_id
     , CASE WHEN UPPER(TRIM(r.club_name)) IN ('', 'NULL') THEN NULL ELSE TRIM(r.club_name) END club_name
     , CASE WHEN UPPER(TRIM(r.addr_line1)) IN ('', 'NULL') THEN NULL ELSE TRIM(r.addr_line1) END addr_line1
     , CASE WHEN UPPER(TRIM(r.addr_line2)) IN ('', 'NULL') THEN NULL ELSE TRIM(r.addr_line2) END addr_line2
     , CASE WHEN UPPER(TRIM(r.addr_line3)) IN ('', 'NULL') THEN NULL ELSE TRIM(r.addr_line3) END addr_line3
     , CASE WHEN UPPER(TRIM(r.city)) IN ('', 'NULL') THEN NULL ELSE TRIM(r.city) END city
     , CASE WHEN UPPER(TRIM(r.postcode2)) IN ('', 'NULL') THEN NULL ELSE TRIM(r.postcode2) END postcode2
     , CASE WHEN UPPER(TRIM(r.sec_first)) IN ('', 'NULL') THEN NULL ELSE TRIM(r.sec_first) END sec_first
     , CASE WHEN UPPER(TRIM(r.sec_last)) IN ('', 'NULL') THEN NULL ELSE TRIM(r.sec_last) END sec_last
     , CASE WHEN UPPER(TRIM(r.sec_email)) IN ('', 'NULL') THEN NULL ELSE TRIM(r.sec_email) END sec_email
     , r.load_status
  FROM cte_all r
     )
SELECT a.*
     , CASE WHEN NULLIF(TRIM(a.addr_line1), '') IS NULL
              OR NULLIF(TRIM(a.city), '')       IS NULL
              OR NULLIF(TRIM(a.postcode2), '')  IS NULL
            THEN 'MISSING ADDRESS' END                         flag_address
     , CASE WHEN NULLIF(TRIM(a.sec_email), '') IS NULL
            THEN 'MISSING EMAIL'
            WHEN NOT REGEXP_MATCHES(TRIM(a.sec_email), '^[^@\s]+@[^@\s]+\.[^@\s]+$')
            THEN 'INVALID EMAIL' END                           flag_email
     , CASE WHEN NULLIF(TRIM(a.sec_first), '') IS NULL
              OR NULLIF(TRIM(a.sec_last), '')  IS NULL
              OR NULLIF(TRIM(a.club_name), '') IS NULL
              OR NULLIF(TRIM(a.pff_id), '')    IS NULL
            THEN 'MISSING MANDATORY' END                       flag_mandatory
  FROM cte_clean a
 ORDER BY a.load_status, a.pff_id
;
