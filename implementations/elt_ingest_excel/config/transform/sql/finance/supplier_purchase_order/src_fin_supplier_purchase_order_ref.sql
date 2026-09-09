DROP TABLE IF EXISTS src_fin_supplier_purchase_order_ref
;
CREATE TABLE src_fin_supplier_purchase_order_ref
    AS
  WITH cte_ref_mapping
    AS (
SELECT
       UPPER(TRIM(r.old_code))                           source_ref_id
     , TRIM(r.workday_ref_id)                            target_ref_id
     , TRIM(r.workday_name)                              target_name
     , r.*
  FROM fin_ref_wd_fdm_mapping                            r
       )
SELECT
       t.spreadsheet_key
     , t.wd_supplier_id
     , t.nrm_company_name
     , t.nrm_supplier_code
     , t.nrm_supplier_name
     , t.nrm_agg_supplier_number
     , t.accountingelement1
     , t.accountingelement2
     , t.accountingelement3
     , t.accountingelement4
     , cc.old_segment                                    hierarchy_cost_centre
     , cc.target_ref_id                                  cost_centre_ref_id
     , sc.old_segment                                    hierarchy_spend_category
     , sc.target_ref_id                                  spend_category_ref_id
     , ev.old_segment                                    hierarchy_event
     , ev.target_ref_id                                  event_ref_id
     , camp.old_segment                                  hierarchy_camp
     , camp.target_ref_id                                camp_ref_id
     , fae.old_segment                                   hierarchy_fa_learning
     , fae.target_ref_id                                 fa_learning_course_ref_id
     , prog.old_segment                                  hierarchy_program
     , prog.target_ref_id                                program_ref_id
     , gr.old_segment                                    hierarchy_grant
     , gr.target_ref_id                                  grant_ref_id
     , cp.old_segment                                    hierarchy_commercial_partner
     , cp.target_ref_id                                  commercial_partner_ref_id
     , CASE SUBSTRING(TRIM(t.nrm_accountingelement2), 1, 1)
         WHEN '1' THEN 'FA Operations'
         WHEN '2' THEN 'FA Operations'
         WHEN '3' THEN 'FA Operations'
         WHEN '4' THEN 'FA Investments'
         WHEN '5' THEN 'National Game Board (NGB)'
         WHEN '6' THEN 'Professional Game Board (PGB)'
         WHEN '7' THEN 'FA Operations'
         WHEN '8' THEN 'FA Operations'
         WHEN '9' THEN 'FA Operations'
         ELSE NULL
       END                                               reporting_type
  FROM src_fin_supplier_purchase_order                   t
       LEFT JOIN LATERAL 
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Cost Center'))
   AND m.source_ref_id                                   IN (t.nrm_accountingelement1, t.nrm_accountingelement2, t.nrm_accountingelement3, t.nrm_accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 cc
         ON TRUE
       LEFT JOIN LATERAL 
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Spend Category'))
   AND m.source_ref_id                                   IN (t.nrm_accountingelement1, t.nrm_accountingelement2, t.nrm_accountingelement3, t.nrm_accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 sc
         ON TRUE
       LEFT JOIN LATERAL 
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Events - Custom Org'))
   AND m.source_ref_id                                   IN (t.nrm_accountingelement1, t.nrm_accountingelement2, t.nrm_accountingelement3, t.nrm_accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 ev
         ON TRUE
       LEFT JOIN LATERAL 
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Camps - Custom Org'))
   AND m.source_ref_id                                   IN (t.nrm_accountingelement1, t.nrm_accountingelement2, t.nrm_accountingelement3, t.nrm_accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 camp
         ON TRUE
       LEFT JOIN LATERAL 
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('FA Learning - Custom Org'))
   AND m.source_ref_id                                   IN (t.nrm_accountingelement1, t.nrm_accountingelement2, t.nrm_accountingelement3, t.nrm_accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 fae
         ON TRUE
       LEFT JOIN LATERAL 
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Program - Delivered Org'))
   AND m.source_ref_id                                   IN (t.nrm_accountingelement1, t.nrm_accountingelement2, t.nrm_accountingelement3, t.nrm_accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 prog
         ON TRUE
       LEFT JOIN LATERAL 
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Commercial Partners - Custom Worktag'))
   AND m.source_ref_id                                   IN (t.nrm_accountingelement1, t.nrm_accountingelement2, t.nrm_accountingelement3, t.nrm_accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 cp
         ON TRUE
       LEFT JOIN LATERAL (
SELECT
       m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Grant'))
   AND m.source_ref_id                                   IN (t.nrm_accountingelement1, t.nrm_accountingelement2, t.nrm_accountingelement3, t.nrm_accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 gr
         ON TRUE
;
