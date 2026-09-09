DROP TABLE IF EXISTS src_hcm_contingent_worker_nrm
;
CREATE TABLE src_hcm_contingent_worker_nrm
    AS
  WITH cte_hcm_contingent_worker_nrm
    AS (
SELECT
       NULLIF(list_aggregate(list_transform(split(LOWER(TRIM(t.first_name            )), ' '), x -> upper(x[1]) || substr(x, 2)), 'string_agg', ' '), '') nrm_first_name
     , NULLIF(list_aggregate(list_transform(split(LOWER(TRIM(t.last_name             )), ' '), x -> upper(x[1]) || substr(x, 2)), 'string_agg', ' '), '') nrm_last_name
     , NULLIF(list_aggregate(list_transform(split(LOWER(TRIM(t.manager_name          )), ' '), x -> upper(x[1]) || substr(x, 2)), 'string_agg', ' '), '') nrm_manager_name
     , NULLIF(list_aggregate(list_transform(split(LOWER(TRIM(REGEXP_REPLACE(t.title, '[^a-zA-Z0-9 ]', '', 'g'))), ' '), x -> upper(x[1]) || substr(x, 2)), 'string_agg', ' '), '') nrm_title
     , NULLIF(LOWER(TRIM(t.user_id)), '')                                  nrm_user_id
     , NULLIF(LOWER(TRIM(t.username)), '')                                 nrm_username
     , NULLIF(TRIM(t.employee_number), '')                                 nrm_employee_number
     , 'Full Time'                                                         nrm_time_type
     , '40'                                                                nrm_hours_per_week
     , NULLIF(LOWER(TRIM(t.primary_email           )), '')                 nrm_primary_email
     , NULLIF(LOWER(TRIM(t.secondary_email         )), '')                 nrm_secondary_email
     , NULLIF(LOWER(TRIM(t.managerid               )), '')                 nrm_managerid
     , NULLIF(LOWER(TRIM(t.manager_email_address   )), '')                 nrm_manager_email_address
     , NULLIF(LOWER(TRIM(t.manager_email           )), '')                 nrm_manager_email
     , NULLIF(LOWER(TRIM(t.department              )), '')                 nrm_department
     , NULLIF(LOWER(TRIM(t.department_1            )), '')                 nrm_department_1
     , NULLIF(LOWER(TRIM(t.division                )), '')                 nrm_division
     , NULLIF(LOWER(TRIM(t.location                )), '')                 nrm_location
     , strftime(
          COALESCE(
          TRY_STRPTIME(NULLIF(TRIM(t.created_date ), ''), '%Y-%m-%dT%H:%M:%S.%fZ')
        , TRY_STRPTIME(NULLIF(TRIM(t.created_date ), ''), '%Y-%m-%d %H:%M:%S')
        , TRY_STRPTIME(NULLIF(TRIM(t.created_date ), ''), '%Y-%m-%d')
        , TRY_STRPTIME(NULLIF(TRIM(t.created_date ), ''), '%d-%m-%Y')
        ), '%Y-%m-%d 00:00:00'
       )                                                                   nrm_created_date
     , strftime(
          COALESCE(
          TRY_STRPTIME(NULLIF(TRIM(t.activated_date ), ''), '%Y-%m-%dT%H:%M:%S.%fZ')
        , TRY_STRPTIME(NULLIF(TRIM(t.activated_date ), ''), '%Y-%m-%d %H:%M:%S')
        , TRY_STRPTIME(NULLIF(TRIM(t.activated_date ), ''), '%Y-%m-%d')
        , TRY_STRPTIME(NULLIF(TRIM(t.activated_date ), ''), '%d-%m-%Y')
        ), '%Y-%m-%d 00:00:00'
       )                                                                   nrm_activated_date
     , strftime(
          COALESCE(
          TRY_STRPTIME(NULLIF(TRIM(t.last_login_date ), ''), '%Y-%m-%dT%H:%M:%S.%fZ')
        , TRY_STRPTIME(NULLIF(TRIM(t.last_login_date ), ''), '%Y-%m-%d %H:%M:%S')
        , TRY_STRPTIME(NULLIF(TRIM(t.last_login_date ), ''), '%Y-%m-%d')
        , TRY_STRPTIME(NULLIF(TRIM(t.last_login_date ), ''), '%d-%m-%Y')
        ), '%Y-%m-%d 00:00:00'
       )                                                                   nrm_last_login_date
     , strftime(
         COALESCE(
           TRY_STRPTIME(NULLIF(TRIM(t.activation_date), ''), '%Y-%m-%dT%H:%M:%S.%fZ')
         , TRY_STRPTIME(NULLIF(TRIM(t.activation_date), ''), '%Y-%m-%d %H:%M:%S')
         , TRY_STRPTIME(NULLIF(TRIM(t.activation_date), ''), '%Y-%m-%d')
         , TRY_STRPTIME(NULLIF(TRIM(t.activation_date), ''), '%d-%m-%Y')
         ), '%Y-%m-%d 00:00:00'
       )                                                                   nrm_activation_date
     , strftime(
         COALESCE(
           TRY_STRPTIME(NULLIF(TRIM(t.deactivation_date), ''), '%Y-%m-%dT%H:%M:%S.%fZ')
         , TRY_STRPTIME(NULLIF(TRIM(t.deactivation_date), ''), '%Y-%m-%d %H:%M:%S')
         , TRY_STRPTIME(NULLIF(TRIM(t.deactivation_date), ''), '%Y-%m-%d')
         , TRY_STRPTIME(NULLIF(TRIM(t.deactivation_date), ''), '%d-%m-%Y')
         , CURRENT_DATE + INTERVAL '6 MONTH'
         ), '%Y-%m-%d 00:00:00'
       )                                                                   nrm_deactivation_date
     , COALESCE(l.target_value, 'Wembley Stadium')                         nrm_location
     , COALESCE(rw.mapped_value, 'Sole Trader Staff')                      nrm_worker_type
     , t.*
  FROM 
       src_hcm_contingent_worker_base     t
       LEFT OUTER JOIN 
       ref_location_mapping               l
          ON l.source_column              = 'location'
         AND UPPER(l.source_value)        = UPPER(COALESCE(NULLIF(TRIM(t.location), ''), 'Wembley Stadium'))
       LEFT OUTER JOIN
       ref_worker_type_mapping            rw
          ON UPPER(rw.user_type)          = UPPER(COALESCE(NULLIF(TRIM(t.user_type)   , ''), 'NULL'))
         AND UPPER(rw.department_1)       = UPPER(COALESCE(NULLIF(TRIM(t.department_1), ''), 'NULL'))
       )
SELECT t.*
  FROM cte_hcm_contingent_worker_nrm t
 WHERE NOT
       (
       t.nrm_department                   IS NULL
   AND t.nrm_division                     IS NULL
   AND t.nrm_location                     IS NULL
   AND t.nrm_manager_name                 IS NULL
       )
;
