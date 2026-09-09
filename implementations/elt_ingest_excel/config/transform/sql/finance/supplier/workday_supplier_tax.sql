DROP TABLE IF EXISTS workday_supplier_tax
;
CREATE TABLE workday_supplier_tax 
    AS
SELECT
       t.supplier_id                                                    supplier_id
     , t.nrm_supplier_name                                              supplier_name
     , fn_nvl2(t.nrm_tax_id_number, NULL, t.nrm_country_name)           tax_id_country
     , t.nrm_tax_id_number                                              tax_id
     , fn_nvl2(t.nrm_tax_id_number, NULL,
         CASE
           WHEN t.nrm_country_code = 'GB'
           THEN CASE
                  WHEN NULLIF(REPLACE(REPLACE(t.nrm_tax_id_number, ' ', ''), '-', ''), '') ~ '^[0-9]{9}$'
                  THEN 'VAT Reg No'
                  WHEN NULLIF(REPLACE(REPLACE(t.nrm_tax_id_number, ' ', ''), '-', ''), '') ~ '^[A-Za-z0-9]{8}$'
                  THEN 'Company Number'
                  WHEN NULLIF(REPLACE(REPLACE(t.nrm_tax_id_number, ' ', ''), '-', ''), '') ~ '^[0-9]{10}$'
                  THEN 'UTR - Unique Taxpayer Reference'
                  ELSE 'VAT Reg No'
                END
           ELSE COALESCE(dm.tax_id_type_label, 'TIN')
         END
       )                                                                tax_id_type
     , fn_nvl2(t.nrm_tax_id_number, NULL, 'Yes')                        primary_tax_id
     , fn_nvl2(t.nrm_tax_id_number, NULL, 'Yes')                        transaction_tax_id
     , t.nrm_country_name                                               tax_status_country
     , CAST(NULL AS VARCHAR)                                            tax_status
     , CASE
         WHEN UPPER(t.nrm_country_code) = 'GB'
         THEN CASE
                WHEN t.nrm_tax_schedule_id = 'POS'
                THEN 'Supplier of Services - VAT/GST Registered in this Country'
                WHEN t.nrm_tax_schedule_id IN ('PZ','PX')
                THEN 'Supplier of VAT/GST Exempt Goods'
                ELSE 'Supplier of Goods - VAT/GST Registered in this Country'
              END
         WHEN UPPER(t.nrm_country_code) IN (
                'AT','BE','BG','HR','CY','CZ','DE','DK','EE','ES','FI','FR'
              , 'GR','HU','IE','IT','LT','LU','LV','MT','NL','PL','PT','RO'
              , 'SE','SI','SK'
              )
         THEN CASE
                WHEN t.nrm_tax_schedule_id = 'POS'
                THEN 'Supplier of Services – Buyer to Reverse Charge VAT/GST'
                ELSE 'Supplier of Goods – Buyer to Reverse Charge VAT/GST'
              END
         ELSE CASE
                WHEN t.nrm_tax_schedule_id = 'POS'
                THEN 'Supplier of Services - VAT/GST Registered in this Country'
                ELSE 'Supplier of Goods - VAT/GST Registered in this Country'
              END
       END                                                              transaction_tax_status
     , CAST(NULL AS VARCHAR)                                            withholding_tax_status
     , CAST(NULL AS VARCHAR)                                            tax_authority_form_type
     , CAST(NULL AS VARCHAR)                                            irs_1099_supplier
     , CAST(NULL AS VARCHAR)                                            tax_document_date
     , CAST(NULL AS VARCHAR)                                            default_tax_code
     , CAST(NULL AS VARCHAR)                                            default_withholding_tax_code
     , CAST(NULL AS VARCHAR)                                            fatca
     , CAST(NULL AS VARCHAR)                                            business_entity_tax_id
  FROM src_fin_supplier                t
       LEFT OUTER JOIN
       ref_country_tax_id_type_mapping dm
         ON  dm.country_code           = t.nrm_country_code
         AND dm.is_default             = TRUE
 ORDER BY
       supplier_id
;
