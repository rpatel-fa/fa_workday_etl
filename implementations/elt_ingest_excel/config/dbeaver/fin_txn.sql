SELECT count(*) FROM src_fin_customer_invoice_base;
SELECT count(*) FROM src_fin_customer_invoice_nrm;
SELECT * FROM src_fin_customer_invoice_nrm WHERE wd_customer_id IS NULL;
SELECT count(*) FROM src_fin_customer_invoice;
SELECT count(*) FROM src_fin_customer_invoice;
SELECT COUNT(*) FROM workday_customer_credit_note;
SELECT COUNT(*) FROM workday_customer_invoice;


SELECT count(*) FROM src_fin_supplier_invoice_base;
SELECT count(*) FROM src_fin_supplier_invoice_nrm;
SELECT * FROM src_fin_supplier_invoice_nrm WHERE wd_supplier_id IS NULL;
SELECT COUNT(*) FROM src_fin_supplier_invoice;
SELECT COUNT(*) FROM workday_supplier_credit_note;
SELECT COUNT(*) FROM workday_supplier_invoice;

SELECT COUNT(*) FROM src_fin_supplier_purchase_order_base;
SELECT COUNT(*) FROM src_fin_supplier_purchase_order_nrm;
SELECT * FROM src_fin_supplier_purchase_order_nrm WHERE wd_supplier_id IS NULL;
SELECT * FROM src_fin_supplier_purchase_order_nrm WHERE nrm_amount_excl_vat = 0;
SELECT COUNT(*) FROM src_fin_supplier_purchase_order;
SELECT COUNT(*) FROM workday_supplier_purchase_order;

SELECT * FROM src_fin_customer_invoice;
SELECT * FROM workday_customer_credit_note;
SELECT * FROM workday_customer_invoice;

SELECT DISTINCT document_type FROM src_fin_customer_invoice;

SELECT * FROM src_fin_customer_invoice WHERE upper(customer_name) LIKE '%FYLDE%'

SELECT * FROM workday_customer_credit_note WHERE column_aj = 'C00808'