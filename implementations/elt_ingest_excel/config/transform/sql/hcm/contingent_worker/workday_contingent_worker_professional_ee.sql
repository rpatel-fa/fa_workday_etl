DROP TABLE IF EXISTS workday_contingent_worker_professional_ee
;
CREATE TABLE workday_contingent_worker_professional_ee
    AS
SELECT
       t.contingent_worker_id                                              worker_id
     , t.nrm_first_name                                                    first_name
     , t.nrm_last_name                                                     last_name
     , 'Contingent Worker'                                                 worker_type
     , t.nrm_department_1                                                  applicant_source_category
     , COALESCE(t.nrm_department, t.nrm_department_1)                      applicant_source_name
     , t.nrm_activated_date                                                hire_date_contract_start_date
     , t.nrm_created_date                                                  original_hire_date
     , NULL                                                                probation_end_date
     , NULL                                                                continuous_service_date
     , NULL                                                                seniority_date
     , CASE t.nrm_department_1
        WHEN UPPER('Consultant')
        THEN 'Sole Trader Staff'
        ELSE 'Supplier Staff'
       END                                                                 employee_contingent_worker_type
     , t.nrm_deactivation_date                                             employment_contract_end_date_fixed_term
     , t.nrm_time_type                                                     time_type
     , NULL                                                                pay_rate_type
     , NULL                                                                job_profile
     , NULL                                                                job_code
     , NULL                                                                position_id
     , t.nrm_location                                                      primary_position
     , COALESCE(nrm_title, 'Contingent Worker')                            position_name
     , t.nrm_activated_date                                                position_start_date
     , t.nrm_title                                                         business_title
     , t.nrm_location                                                      location_name
     , NULL                                                                work_space
     , t.nrm_hours_per_week                                                scheduled_weekly_hours
     , t.nrm_hours_per_week                                                default_weekly_hours
     , NULL                                                                fte
     , COALESCE(t.nrm_division, t.nrm_department_1, t.nrm_location)        company_org_name
     , NULL                                                                company_code
     , NULL                                                                establishment
     , NULL                                                                cost_center_code
     , NULL                                                                fund
     , NULL                                                                work_shift_country
     , NULL                                                                work_shift
     , NULL                                                                work_shift_id
     , NULL                                                                distance_home_work
     , NULL                                                                contingent_worker_agency
     , NULL                                                                remote_ee
     , NULL                                                                custom_org_1
     , NULL                                                                custom_org_2
     , NULL                                                                custom_org_3
     , NULL                                                                annual_work_period
     , NULL                                                                disbursement_plan_period
     , NULL                                                                contract_pay_rate
     , NULL                                                                contract_pay_currency
     , NULL                                                                contract_pay_frequency
     , NULL                                                                working_time_frequency
     , NULL                                                                working_time_unit
     , NULL                                                                working_time_value
     , NULL                                                                working_fte
     , NULL                                                                paid_fte
     , NULL                                                                region_assignment
     , NULL                                                                employment_system
     , NULL                                                                time_type_subtype
     , NULL                                                                appointment_type
     , NULL                                                                employee_tenure
     , NULL                                                                annuitant_indicator
     , NULL                                                                job_classification
     , NULL                                                                hire_employee_1
     , NULL                                                                hire_employee_2
     , NULL                                                                contract_contingent
     , NULL                                                                additional_job
  FROM src_hcm_contingent_worker       t
 ORDER BY
       worker_id
;
