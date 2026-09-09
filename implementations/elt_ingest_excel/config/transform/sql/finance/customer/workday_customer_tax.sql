DROP TABLE IF EXISTS workday_customer_tax
;
CREATE TABLE workday_customer_tax
    AS
SELECT
       c.customer_id                                                             customer_id
     , c.nrm_customer_name                                                       customer_name
     , 'GBR 20% VAT (20%)'                                                       tax_code
     , fn_nvl2(c.nrm_tax_registration_number, NULL, c.nrm_country_name)          tax_id_country
     , c.nrm_tax_registration_number                                             tax_id
     , fn_nvl2(
         c.nrm_tax_registration_number,
         NULL,
         CASE
           WHEN c.nrm_country_code = 'GB'
           THEN CASE
                  WHEN NULLIF(REPLACE(REPLACE(c.nrm_tax_registration_number, ' ', ''), '-', ''), '') ~ '^[0-9]{9}$'
                  THEN 'VAT Reg No'
                  WHEN NULLIF(REPLACE(REPLACE(c.nrm_tax_registration_number, ' ', ''), '-', ''), '') ~ '^[A-Za-z0-9]{8}$'
                  THEN 'Company Number'
                  WHEN NULLIF(REPLACE(REPLACE(c.nrm_tax_registration_number, ' ', ''), '-', ''), '') ~ '^[0-9]{10}$'
                  THEN 'UTR - Unique Taxpayer Reference'
                  ELSE 'VAT Reg No'
                END
           ELSE COALESCE(dm.tax_id_type_label, 'TIN')
         END
       )                                                                         tax_id_type
     , fn_nvl2(c.nrm_tax_registration_number, NULL, 'Yes')                       primary_tax_id
     , fn_nvl2(c.nrm_tax_registration_number, NULL, 'Yes')                       transaction_tax_id
     , c.nrm_country_name                                                        tax_status_country
     , CASE
         WHEN UPPER(c.nrm_country_code) = 'GB'
         THEN CASE
                WHEN c.nrm_tax_schedule_id = 'SOS'
                THEN 'Customer of Services'
                ELSE 'Customer of Goods'
              END
         WHEN UPPER(c.nrm_country_code) IN (
                'AT','BE','BG','HR','CY','CZ','DE','DK','EE','ES','FI','FR'
              , 'GR','HU','IE','IT','LT','LU','LV','MT','NL','PL','PT','RO'
              , 'SE','SI','SK'
              )
         THEN CASE
                WHEN c.nrm_tax_registration_number IS NOT NULL
                THEN 'VAT Registered in THIS EU country- Domestic Scenario'
                WHEN c.nrm_tax_schedule_id = 'SZ'
                THEN 'EU Company - Exempt in THIS country'
                ELSE 'EU Company - Not Registered in this country'
              END
         ELSE CASE
                WHEN c.nrm_tax_registration_number IS NOT NULL
                THEN 'VAT Registered in THIS NON EU country- Domestic Scenario'
                ELSE 'Non EU Company Not VAT registered in this country'
              END
       END                                                                      transaction_tax_status
     , CAST(NULL AS VARCHAR)                                                    withholding_tax_status
  FROM src_fin_customer                                                         c
       LEFT OUTER JOIN
       ref_country_tax_id_type_mapping                                          dm
         ON  dm.country_code                                                    = c.nrm_country_code
         AND dm.is_default                                                      = TRUE
 ORDER BY
       customer_id
;
