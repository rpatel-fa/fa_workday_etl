DROP TABLE IF EXISTS workday_pff_os_customer_load
;
CREATE TABLE workday_pff_os_customer_load
    AS
SELECT a.* EXCLUDE (wd_customer_status)
  FROM workday_pff_os_customer_load_base a
 ORDER BY a.load_status, a.pff_id
;
