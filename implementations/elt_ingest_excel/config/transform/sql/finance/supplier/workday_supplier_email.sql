DROP TABLE IF EXISTS workday_supplier_email
;
CREATE TABLE workday_supplier_email 
    AS
  WITH cte_supplier_email
    AS (
SELECT DISTINCT
       s.supplier_id                                        supplier_id
     , s.nrm_supplier_name                                  supplier_name
     , LOWER(s.nrm_agg_email_to_address)                    email_raw
     , 'to'                                                 email_type
     , '_EM1'                                               suffix
  FROM src_fin_supplier s
 WHERE NULLIF(UPPER(TRIM(s.nrm_agg_email_to_address)), '')  IS NOT NULL
UNION ALL
SELECT DISTINCT
       s.supplier_id                                        supplier_id
     , s.nrm_supplier_name                                  supplier_name
     , LOWER(s.nrm_agg_email_cc_address)                    email_raw
     , 'cc'                                                 email_type
     , '_EM2'                                               suffix
  FROM src_fin_supplier s
 WHERE NULLIF(UPPER(TRIM(s.nrm_agg_email_cc_address)), '')  IS NOT NULL
UNION ALL
SELECT DISTINCT
       s.supplier_id                                        supplier_id
     , s.nrm_supplier_name                                  supplier_name
     , LOWER(s.nrm_agg_email_bcc_address)                   email_raw
     , 'bcc'                                                email_type
     , '_EM3'                                               suffix
  FROM src_fin_supplier s
 WHERE NULLIF(UPPER(TRIM(s.nrm_agg_email_bcc_address)), '') IS NOT NULL
       )
     , cte_email_split
    AS (
SELECT DISTINCT
       e.supplier_id                            supplier_id
     , e.supplier_name                          supplier_name
     , TRIM(u.email)                            email_address
     , e.email_type                             email_type
     , e.suffix                                 suffix
  FROM cte_supplier_email                       e
     , UNNEST(STRING_SPLIT(e.email_raw, ';')) u(email)
       )
     , cte_email_unique
    AS (
SELECT 
       t.supplier_id                            supplier_id
     , t.supplier_name                          supplier_name
     , t.email_address                          email_address
     , t.email_type                             email_type
     , t.suffix                                 suffix
     , RANK() OVER (PARTITION BY email_address
                        ORDER BY supplier_id
                               , suffix                           
                   )                            unique_email_rnk
  FROM cte_email_split                          t
 WHERE NULLIF(TRIM(t.email_address), '')        IS NOT NULL
   AND t.email_address                          ~ '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'
       )
     , cte_email_rnk
    AS (
SELECT s.supplier_id                          supplier_id
     , s.supplier_name                        supplier_name
     , s.email_address                        email_address
     , s.email_type                           email_type
     , s.suffix                               suffix
     , ROW_NUMBER() OVER (
           PARTITION BY s.supplier_id
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
SELECT s.supplier_id                                                       supplier_id
     , s.supplier_name                                                     supplier_name
     , TRIM(LOWER(s.email_address))                                        email_address
     , CAST(NULL AS VARCHAR)                                               email_comment
     , s.supplier_id || s.suffix || '_' || ROW_NUMBER() OVER (
           PARTITION BY s.supplier_id, s.suffix
               ORDER BY s.email_address
       )                                                                   email_id
     , 'Yes'                                                               public_flag
     , CASE s.email_rank WHEN 1 THEN 'Yes' ELSE 'No' END                   primary_flag
     , CASE s.email_rank WHEN 1 THEN 'Yes' ELSE 'No' END                   default_po
     , CAST(NULL AS VARCHAR)                                               email_type
     , 'SHIPPING|BILLING|REMIT'                                            use_for
     , CAST(NULL AS VARCHAR)                                               use_for_tenanted
     , CAST(NULL AS VARCHAR)                                               delete_flag
     , CAST(NULL AS VARCHAR)                                               do_not_replace_all
     , CAST(NULL AS VARCHAR)                                               additional_comments
     , CAST(NULL AS VARCHAR)                                               email
  FROM cte_email_rnk s
 ORDER BY 
       supplier_id
     , primary_flag DESC
     , email_id
;