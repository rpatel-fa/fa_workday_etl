SELECT * FROM fin_workday_extract_customer;
SELECT * FROM fin_workday_extract_supplier;

SELECT count(*) FROM src_fin_supplier;
SELECT * FROM src_fin_supplier_purchase_order;
SELECT * FROM src_fin_supplier_invoice;

SELECT count(*) FROM src_fin_customer;
SELECT * FROM src_fin_customer_invoice_base;
SELECT * FROM src_fin_customer_invoice_nrm WHERE wd_customer_id IS null;
SELECT * FROM src_fin_customer_invoice;

SELECT * FROM fin_supplier_purchase_order;
SELECT * FROM fin_supplier_purchase_order_v1;
SELECT * FROM fin_ref_wd_fdm_mapping;

SELECT count(*) FROM src_fin_customer_base;
SELECT count(*) FROM src_fin_supplier_base;
SELECT count(*) FROM src_fin_customer;
SELECT count(*) FROM src_fin_supplier;
SELECT count(*) FROM fin_workday_extract_customer;
SELECT count(*) FROM fin_workday_extract_supplier;
SELECT count(*) FROM src_fin_customer_match;
SELECT count(*) FROM src_fin_supplier_match;
SELECT count(*) FROM src_fin_customer_missing;
SELECT count(*) FROM src_fin_supplier_missing;

SELECT nrm_country_code, nrm_country_name, count(*) FROM src_fin_customer_match   GROUP BY nrm_country_code, nrm_country_name ORDER BY 2;
SELECT nrm_country_code, nrm_country_name, count(*) FROM src_fin_customer_missing GROUP BY nrm_country_code, nrm_country_name ORDER BY 2;
SELECT nrm_country_code, nrm_country_name, count(*) FROM src_fin_supplier_match   GROUP BY nrm_country_code, nrm_country_name ORDER BY 2;
SELECT nrm_country_code, nrm_country_name, count(*) FROM src_fin_supplier_missing GROUP BY nrm_country_code, nrm_country_name ORDER BY 2;

SELECT * FROM validation_customer_issues WHERE detail != 'All Currencies';
SELECT * FROM validation_supplier_issues WHERE col NOT IN ('currency', 'accepted_currencies_plus');
