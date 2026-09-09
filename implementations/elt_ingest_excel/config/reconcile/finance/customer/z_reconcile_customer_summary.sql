DROP TABLE IF EXISTS z_reconcile_customer_summary
;
CREATE TABLE z_reconcile_customer_summary
    AS
SELECT rule_type
     , table_name
     , COUNT(*)    issue_count
  FROM reconciliation_customer_issues
 GROUP BY rule_type
     , table_name
 ORDER BY rule_type
     , table_name
;
