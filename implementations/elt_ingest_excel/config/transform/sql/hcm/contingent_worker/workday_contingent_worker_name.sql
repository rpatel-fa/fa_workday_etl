DROP TABLE IF EXISTS workday_contingent_worker_name
;
CREATE TABLE workday_contingent_worker_name
    AS
SELECT
       t.contingent_worker_id                            worker_id
     , 'Contingent Worker'                               worker_type
     , t.nrm_first_name                                  first_name
     , NULL                                              middle_name
     , t.nrm_last_name                                   last_name
     , 'United Kingdom'                                  legal_name_country
     , 'United Kingdom'                                  payroll_country
     , NULL                                              gender
     , t.nrm_title                                       title
     , NULL                                              family_name_prefix
     , NULL                                              suffix
     , NULL                                              initials
     , NULL                                              maiden_name
     , NULL                                              preferred_first_name
     , NULL                                              preferred_middle_name
     , NULL                                              preferred_last_name
     , NULL                                              local_script_first_name
     , NULL                                              local_script_last_name
     , NULL                                              first_name_former_name
     , NULL                                              last_name_former_name
     , NULL                                              title_id
     , NULL                                              prefix_id
     , NULL                                              suffix_id
  FROM src_hcm_contingent_worker       t
 ORDER BY
       worker_id
;
