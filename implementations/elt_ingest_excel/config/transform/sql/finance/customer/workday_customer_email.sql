DROP TABLE IF EXISTS workday_customer_email
;
CREATE TABLE workday_customer_email AS
  WITH cte_customer_email
    AS (
SELECT DISTINCT
       c.customer_id                                           customer_id
     , c.nrm_customer_name                                     customer_name
     , LOWER(c.nrm_agg_email_to_address)                       email_raw
     , 'to'                                                    email_type
     , '_EM1'                                                  suffix
  FROM src_fin_customer c
 WHERE NULLIF(UPPER(TRIM(c.nrm_agg_email_to_address)), '')     IS NOT NULL
UNION ALL
SELECT DISTINCT
       c.customer_id                                           customer_id
     , c.nrm_customer_name                                     customer_name
     , LOWER(c.nrm_agg_email_cc_address)                       email_raw
     , 'cc'                                                    email_type
     , '_EM2'                                                  suffix
  FROM src_fin_customer c
 WHERE NULLIF(UPPER(TRIM(c.nrm_agg_email_cc_address)), '')     IS NOT NULL
UNION ALL
SELECT DISTINCT
       c.customer_id                                           customer_id
     , c.nrm_customer_name                                     customer_name
     , LOWER(c.nrm_agg_email_bcc_address)                      email_raw
     , 'bcc'                                                   email_type
     , '_EM3'                                                  suffix
  FROM src_fin_customer c
 WHERE NULLIF(UPPER(TRIM(c.nrm_agg_email_bcc_address)), '')    IS NOT NULL
       )
     , cte_email_split
    AS (
SELECT DISTINCT
       e.customer_id                            customer_id
     , e.customer_name                          customer_name
     , TRIM(u.email)                            email_address
     , e.email_type                             email_type
     , e.suffix                                 suffix
  FROM cte_customer_email                       e
     , UNNEST(STRING_SPLIT(e.email_raw, ';')) u(email)
       )
     , cte_email_unique
    AS (
SELECT 
       t.customer_id                            customer_id
     , t.customer_name                          customer_name
     , t.email_address                          email_address
     , t.email_type                             email_type
     , t.suffix                                 suffix
     , RANK() OVER (PARTITION BY email_address
                        ORDER BY customer_id
                               , suffix                           
                   )                            unique_email_rnk
  FROM cte_email_split                          t
 WHERE NULLIF(TRIM(t.email_address), '')        IS NOT NULL
   AND t.email_address                          ~ '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'
       )  
     , cte_email_rnk
    AS (
SELECT s.customer_id                          customer_id
     , s.customer_name                        customer_name
     , s.email_address                        email_address
     , s.email_type                           email_type
     , s.suffix                               suffix
     , ROW_NUMBER() OVER (
           PARTITION BY s.customer_id
               ORDER BY CASE s.email_type
                           WHEN 'to'  THEN 1
                           WHEN 'cc'  THEN 2
                           WHEN 'bcc' THEN 3
                           ELSE 4
                        END 
                      , s.email_address
       )                                      email_rank
  FROM cte_email_unique s
 WHERE s.unique_email_rnk                     = 1
       )
SELECT s.customer_id                                                     customer_id
     , s.customer_name                                                   customer_name
     , TRIM(s.customer_id) || s.suffix || '_' || ROW_NUMBER() OVER (
           PARTITION BY s.customer_id, s.suffix ORDER BY s.email_address
       )                                                                 email_id
     , TRIM(LOWER(s.email_address))                                      email_address
     , CAST(NULL AS VARCHAR)                                             email_comment
     , 'Yes'                                                             is_public
     , CASE s.email_rank WHEN 1 THEN 'Yes' ELSE 'No' END                 primary_flag
     , s.email_type                                                      email_type
     , 'Business'                                                        use_for
     , CAST(NULL AS VARCHAR)                                             use_for_tenanted
     , CAST(NULL AS VARCHAR)                                             delete_flag
     , CAST(NULL AS VARCHAR)                                             do_not_replace_all
     , CAST(NULL AS VARCHAR)                                             additional_comment
     , CAST(NULL AS VARCHAR)                                             email
  FROM cte_email_rnk s
 ORDER BY 
       customer_id
     , primary_flag DESC
     , email_id
;