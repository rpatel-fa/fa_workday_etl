DROP TABLE IF EXISTS src_fin_gl_base;

CREATE TABLE src_fin_gl_base
    AS
  WITH cte_base
    AS (
SELECT 'CONS' business_unit, t.* FROM fin_general_ledger_cons t
UNION ALL
SELECT 'FA'   business_unit, t.* FROM fin_general_ledger_fa   t
UNION ALL
SELECT 'NFC'  business_unit, t.* FROM fin_general_ledger_nfc  t
UNION ALL
SELECT 'WNSL' business_unit, t.* FROM fin_general_ledger_wnsl t
       )
SELECT string_split(t.account, '-')[1] accountingelement1
     , string_split(t.account, '-')[2] accountingelement2
     , string_split(t.account, '-')[3] accountingelement3
     , string_split(t.account, '-')[4] accountingelement4
     , CASE (UPPER(TRIM(t.business_unit)))
           WHEN 'FA'   THEN 'THE FA GROUP - FA'
           WHEN 'NFC'  THEN 'THE FA GROUP - NFC'
           WHEN 'WNSL' THEN 'THE FA GROUP - WNSL'
           ELSE 'UNKNOWN'
       END                                                    nrm_company_id
     , t.*
  FROM cte_base                                               t
 WHERE NULLIF(TRIM(t.account), '')                            IS NOT NULL
;  
