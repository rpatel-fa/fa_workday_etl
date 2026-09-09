DROP TABLE IF EXISTS src_hcm_contingent_worker
;
CREATE TABLE src_hcm_contingent_worker
    AS
  WITH cte_contingent_worker_nrm
    AS (
SELECT t.*
     , ROW_NUMBER() OVER (ORDER BY t.username, t.user_id) + 0  row_id
  FROM src_hcm_contingent_worker_nrm   t
       LEFT OUTER JOIN
       src_hcm_contingent_worker_match m
          ON m.nrm_user_id             = t.nrm_user_id
 WHERE COALESCE(m.flg_wd_match_employee         , 0) = 0
       )
     , cte_rnk
    AS (
SELECT
       'CW-' || LPAD(row_id::VARCHAR, 6, '0')   rnk_id
     , t.*
  FROM 
       cte_contingent_worker_nrm       t
       )
SELECT 
       COALESCE(t.nrm_employee_number, rnk_id)  contingent_worker_id
     , t.*
  FROM cte_rnk                         t         
;
