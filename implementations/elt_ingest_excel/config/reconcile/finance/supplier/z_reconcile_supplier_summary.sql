DROP TABLE IF EXISTS z_reconcile_supplier_summary
;
CREATE TABLE z_reconcile_supplier_summary
    AS
SELECT rule_type
     , table_name
     , COUNT(*)    issue_count
  FROM reconciliation_supplier_issues
 GROUP BY rule_type
     , table_name
 ORDER BY rule_type
     , table_name
;
