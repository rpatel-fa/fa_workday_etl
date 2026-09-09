DROP TABLE IF EXISTS workday_supplier_purchase_order
;
CREATE TABLE workday_supplier_purchase_order
    AS
SELECT 
       t.spreadsheet_key                           column_b    -- spreadsheet_key
     , 'Y'                                         column_c    -- add_only
     , 'Y'                                         column_d    -- auto_complete
     , 'Y'                                         column_h    -- submit
     , t.displayno                                 column_j    -- document_number
     , t.nrm_company_id                            column_p    -- company
     , t.wd_supplier_id                            column_q    -- supplier
     , t.nrm_document_date                         column_u    -- document_date
     , t.nrm_document_date                         column_ae   -- due_date
     , t.currency                                  column_ah   -- currency
     , CASE t.nrm_email_exists_ind 
         WHEN 1 THEN 'EMAIL'
         ELSE 'PRINT'
       END                                         column_al   -- issue_option
     , CAST(NULL AS VARCHAR)                       column_az   -- memo
     , CAST(NULL AS VARCHAR)                       column_ba   -- internal memo
     , CAST(NULL AS VARCHAR)                       column_bq   -- Goods Line rowid
     , CAST(NULL AS VARCHAR)                       column_bu   -- line_company
     , CAST(NULL AS VARCHAR)                       column_bz   -- item_description
     , CAST(NULL AS VARCHAR)                       column_cj   -- spend category
     , CAST(NULL AS VARCHAR)                       column_dg   -- unit of measure
     , '1'                                         column_ex   -- Service Line rowid
     , t.nrm_company_id                            column_fb   -- line_company
     , t.description                               column_fc   -- item_description
     , r.spend_category_ref_id                     column_fk   -- resource_category
     , t.amount_excl_vat                           column_gf   -- extended_amount
     , r.cost_centre_ref_id                        column_gv   -- cost_centre
     , r.reporting_type                            column_gw   -- reporting_type
     , r.event_ref_id                              column_gx   -- events
     , r.fa_learning_course_ref_id                 column_gy   -- fa_learning_courses
     , r.camp_ref_id                               column_gz   -- camps
     , r.program_ref_id                            column_ha   -- programmes
     , r.grant_ref_id                              column_hb   -- grant
     , r.commercial_partner_ref_id                 column_hc   -- commercial_partner
  FROM src_fin_supplier_purchase_order             t
       LEFT OUTER JOIN
       src_fin_supplier_purchase_order_ref         r
          ON t.spreadsheet_key                     = r.spreadsheet_key
 ORDER BY 
       column_b
;
