DROP TABLE IF EXISTS src_fin_gl_ref
;
CREATE TABLE src_fin_gl_ref
    AS
  WITH cte_ref_mapping_src
    AS (
SELECT
       UPPER(TRIM(r.old_code))                           source_ref_id
     , TRIM(r.old_segment)                               old_segment
     , UPPER(TRIM(r.core_vs_prism))                      core_vs_prism
     , NULLIF(TRIM(r.workday_business_object), '')       workday_business_object
     , NULLIF(TRIM(r.workday_ref_id), '')                target_ref_id
     , NULLIF(TRIM(r.workday_name), '')                  target_name
     , CASE
         WHEN UPPER(TRIM(r.workday_ledger_code))         IN ('DO NOT MAP', 'NOT APPLICABLE', 'NOT SURE') THEN NULL
         ELSE NULLIF(TRIM(r.workday_ledger_code), '')
       END                                               workday_ledger_code
     , NULLIF(TRIM(r.interco_affiliate_flag), '')        interco_affiliate_flag
  FROM fin_ref_wd_fdm_gl_mapping                         r
       )
     , cte_ref_mapping_base
    AS (
SELECT *
  FROM cte_ref_mapping_src                         r
 WHERE 1 = 1 
   AND core_vs_prism                               NOT IN  ('PRISM')
   /* remove this exclusion - confirmed 20260907
   AND NOT (
           workday_business_object                 = 'Cost Center'
       AND target_ref_id                           IN 
           ( 'CC21001')
           )
   */
       )
     , cte_ref_mapping      
    AS (
SELECT
       source_ref_id
     , old_segment
     , target_ref_id
     , target_name
     , workday_ledger_code
     , workday_business_object
     , interco_affiliate_flag
  FROM cte_ref_mapping_base
       )
SELECT
       t.row_id
     , t.account
     , t.accountingelement1
     , t.accountingelement2
     , t.accountingelement3
     , t.accountingelement4
     , UPPER(TRIM(t.accountingelement1))
       || '-' || UPPER(TRIM(t.accountingelement2))
       || '-' || LPAD(UPPER(TRIM(t.accountingelement3)), 8, '0')
       || '-' || LPAD(UPPER(TRIM(t.accountingelement4)), 4, '0')
       || '-' || UPPER(TRIM(t.business_unit))              account_string_override_key
     , ba.old_segment                                    bank_account_hierarchy
     , ba.target_ref_id                                  bank_account_ref_id
     , ba.workday_ledger_code                            bank_account_workday_ledger_code
       -- Resolved bank account: raw mapping plus the NFC special case. Central
       -- source of truth so consumers (line_base, closing balance) share one
       -- derivation instead of hard-coding the override in each projection.
     , ba.target_ref_id                                  bank_account_ref_id_resolved
     , camp.old_segment                                  camp_hierarchy
     , camp.target_ref_id                                camp_ref_id
     , camp.workday_ledger_code                          camp_workday_ledger_code
     , cp.old_segment                                    commercial_partner_hierarchy
     , cp.target_ref_id                                  commercial_partner_ref_id
     , cp.workday_ledger_code                            commercial_partner_workday_ledger_code
     , cc.old_segment                                    cost_centre_hierarchy
     , cc.target_ref_id                                  cost_centre_ref_id
     , cc.workday_ledger_code                            cost_centre_workday_ledger_code
     , dnm.old_segment                                   do_not_map_hierarchy
     , dnm.target_ref_id                                 do_not_map_ref_id
     , dnm.workday_ledger_code                           do_not_map_workday_ledger_code
     , ev.old_segment                                    event_hierarchy
     , ev.target_ref_id                                  event_ref_id
     , ev.workday_ledger_code                            event_workday_ledger_code
     , fae.old_segment                                   fa_learning_hierarchy
     , fae.target_ref_id                                 fa_learning_course_ref_id
     , fae.workday_ledger_code                           fa_learning_course_workday_ledger_code
     , far.old_segment                                   far_hierarchy
     , far.target_ref_id                                 far_ref_id
     , far.workday_ledger_code                           far_workday_ledger_code
     , gr.old_segment                                    grant_hierarchy
     , gr.target_ref_id                                  grant_ref_id
     , gr.workday_ledger_code                            grant_workday_ledger_code
     , led.old_segment                                   ledger_hierarchy
     , led.target_ref_id                                 ledger_ref_id
     , led.workday_ledger_code                           ledger_workday_ledger_code
     , led.interco_affiliate_flag                        ledger_interco_affiliate_flag
     , lfar.old_segment                                  ledger_far_hierarchy
     , lfar.target_ref_id                                ledger_far_ref_id
     , lfar.workday_ledger_code                          ledger_far_workday_ledger_code
     , lvat.old_segment                                  ledger_vat_dimension_hierarchy
     , lvat.target_ref_id                                ledger_vat_dimension_ref_id
     , lvat.workday_ledger_code                          ledger_vat_dimension_workday_ledger_code
     , ns.old_segment                                    not_sure_hierarchy
     , ns.target_ref_id                                  not_sure_ref_id
     , ns.workday_ledger_code                            not_sure_workday_ledger_code
     , pc.old_segment                                    payroll_component_hierarchy
     , pc.target_ref_id                                  payroll_component_ref_id
     , pc.workday_ledger_code                            payroll_component_workday_ledger_code
     , proj.old_segment                                    project_opex_hierarchy
     , proj.target_ref_id                                  project_opex_ref_id
     , proj.workday_ledger_code                            project_opex_workday_ledger_code
     , projc.old_segment                                   project_capex_hierarchy
     , projc.target_ref_id                                 project_capex_ref_id
     , projc.workday_ledger_code                           project_capex_workday_ledger_code
     , COALESCE(projc.old_segment, proj.old_segment)        project_hierarchy
     , COALESCE(projc.target_ref_id, proj.target_ref_id)    project_ref_id
     , COALESCE(projc.workday_ledger_code, proj.workday_ledger_code) project_workday_ledger_code
     , prog.old_segment                                  program_hierarchy
     , prog.target_ref_id                                program_ref_id
     , prog.workday_ledger_code                          program_workday_ledger_code
     , rev.old_segment                                   revenue_category_hierarchy
     , rev.target_ref_id                                 revenue_category_ref_id
     , rev.workday_ledger_code                           revenue_category_workday_ledger_code
     , si.old_segment                                    sales_item_hierarchy
     , si.target_ref_id                                  sales_item_ref_id
     , si.workday_ledger_code                            sales_item_workday_ledger_code
     , sc.old_segment                                    spend_category_hierarchy
     , sc.target_ref_id                                  spend_category_ref_id
     , sc.workday_ledger_code                            spend_category_workday_ledger_code
     , sco.old_segment                                   spend_category_override_hierarchy
     , sco.target_ref_id                                 spend_category_override_ref_id
     , sco.workday_ledger_code                           spend_category_override_workday_ledger_code
     , scg.old_segment                                   spend_category_grant_hierarchy
     , scg.target_ref_id                                 spend_category_grant_ref_id
     , scg.workday_ledger_code                           spend_category_grant_workday_ledger_code
     , scgo.old_segment                                  spend_category_grant_override_hierarchy
     , scgo.target_ref_id                                spend_category_grant_override_ref_id
     , scgo.workday_ledger_code                          spend_category_grant_override_workday_ledger_code
     , sup.old_segment                                   supplier_hierarchy
     , sup.target_ref_id                                 supplier_ref_id
     , sup.workday_ledger_code                           supplier_workday_ledger_code
     , CASE SUBSTRING(TRIM(t.accountingelement2), 1, 1)
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
  FROM src_fin_gl                                        t
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Bank Account'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 ba
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Camps - Custom Org'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 camp
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Commercial Partners - Custom Worktag'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 cp
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Cost Center'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 cc
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('DO NOT MAP'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 dnm
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Events - Custom Org'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 ev
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('FA Learning - Custom Org'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 fae
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('FAR'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 far
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Grant'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 gr
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Ledger'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 led
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Ledger|FAR'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 lfar
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Ledger|VAT Dimesion'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 lvat
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Not Sure'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 ns
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Payroll Component'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 pc
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Projects (Opex)'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 proj
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Projects (Capex)'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 projc
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Program - Delivered Org'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 prog
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Revenue Category'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 rev
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Sales Item'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 si
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Spend Category'))
   AND UPPER(TRIM(m.old_segment))                       <> UPPER(TRIM('String Override'))
   AND COALESCE(UPPER(TRIM(m.target_ref_id)), '')       <> UPPER(TRIM('STRING OVERRIDE'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 sc
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Spend Category'))
   AND UPPER(TRIM(m.old_segment))                        = UPPER(TRIM('String Override'))
   AND m.source_ref_id                                   = (
             UPPER(TRIM(t.accountingelement1))
       || '-' || UPPER(TRIM(t.accountingelement2))
       || '-' || LPAD(UPPER(TRIM(t.accountingelement3)), 8, '0')
       || '-' || LPAD(UPPER(TRIM(t.accountingelement4)), 4, '0')
       || '-' || UPPER(TRIM(t.business_unit))
       )
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 sco
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Spend Category | Grant'))
   AND UPPER(TRIM(m.old_segment))                       <> UPPER(TRIM('String Override'))
   AND COALESCE(UPPER(TRIM(m.target_ref_id)), '')       <> UPPER(TRIM('STRING OVERRIDE'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 scg
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Spend Category | Grant'))
   AND UPPER(TRIM(m.old_segment))                        = UPPER(TRIM('String Override'))
   AND m.source_ref_id                                   = (
             UPPER(TRIM(t.accountingelement1))
       || '-' || UPPER(TRIM(t.accountingelement2))
       || '-' || LPAD(UPPER(TRIM(t.accountingelement3)), 8, '0')
       || '-' || LPAD(UPPER(TRIM(t.accountingelement4)), 4, '0')
       || '-' || UPPER(TRIM(t.business_unit))
       )
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 scgo
         ON TRUE
       LEFT JOIN LATERAL
       (
SELECT m.*
  FROM cte_ref_mapping                                   m
 WHERE UPPER(TRIM(m.workday_business_object))            = UPPER(TRIM('Supplier'))
   AND m.source_ref_id                                   IN (t.accountingelement1, t.accountingelement2, t.accountingelement3, t.accountingelement4)
 ORDER BY NULLIF(TRIM(m.old_segment), '') DESC NULLS LAST
        , m.old_segment                                  DESC
 LIMIT 1
       )                                                 sup
         ON TRUE
;
