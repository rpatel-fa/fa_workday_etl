-- Ensure all entries have a primry defined
SELECT * FROM workday_supplier_email WHERE supplier_id NOT IN (SELECT supplier_id FROM workday_supplier_email WHERE primary_flag = 'Yes');
SELECT * FROM workday_supplier_phone WHERE supplier_id NOT IN (SELECT supplier_id FROM workday_supplier_phone WHERE primary_flag = 'Yes');
SELECT * FROM workday_customer_email WHERE customer_id NOT IN (SELECT customer_id FROM workday_customer_email WHERE primary_flag = 'Yes');

-- Validate Region population
SELECT * FROM workday_supplier_address WHERE (country,region) NOT IN (SELECT country, instance FROM ref_workday_country_state_region);
SELECT * FROM workday_customer_address WHERE (country,region) NOT IN (SELECT country, instance FROM ref_workday_country_state_region);