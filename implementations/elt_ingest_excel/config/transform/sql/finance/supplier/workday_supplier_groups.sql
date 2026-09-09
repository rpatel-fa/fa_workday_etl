DROP TABLE IF EXISTS workday_supplier_groups
;
CREATE TABLE workday_supplier_groups 
    AS
  WITH cte_group
    AS (
SELECT
       t.supplier_id                                  supplier_id
     , t.nrm_supplier_name                            supplier_name
     , CASE
           WHEN UPPER(SUBSTR(t.nrm_supplier_name, 1, 1)) BETWEEN 'A' AND 'L'
               THEN 'Suppliers A - L'
           WHEN UPPER(SUBSTR(t.nrm_supplier_name, 1, 1)) BETWEEN 'M' AND 'Z'
               THEN 'Suppliers M - Z'
           ELSE 'Suppliers 0 - 9'
       END                                            supplier_group
  FROM src_fin_supplier                t
UNION ALL
SELECT
       t.supplier_id                                  supplier_id
     , t.nrm_supplier_name                            supplier_name
     , r.target_value                                 supplier_group
  FROM src_fin_supplier                t
       INNER JOIN
       ref_workday_group_additional                   r
          ON r.source_type                            = 'supplier'
         AND UPPER(TRIM(r.source_value))              = UPPER(TRIM(t.vendor_class_id))
    )
SELECT *
  FROM cte_group
 ORDER BY
       supplier_id
;
