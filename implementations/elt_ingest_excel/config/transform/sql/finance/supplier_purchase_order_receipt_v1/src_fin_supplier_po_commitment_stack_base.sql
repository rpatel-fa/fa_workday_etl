DROP TABLE IF EXISTS src_fin_supplier_po_commitment_stack_base;

CREATE TABLE src_fin_supplier_po_commitment_stack_base
    AS
SELECT TRIM(companyshortcode)                                                   company_shortcode
     , TRIM(order_number)                                                       order_number
     , TRIM(order_date)                                                         order_date_excel
     , TRIM(supplier)                                                           supplier_name
     , TRIM(supplier_code)                                                      supplier_code
     , TRIM(originator)                                                         originator_name
     , TRIM(description)                                                        line_description
     , TRIM(line_number)                                                        line_number
     , TRIM("position")                                                         po_position
     , TRIM(division)                                                           division_code
     , TRIM(divisionname)                                                       division_name
     , TRIM(divgrouping)                                                        division_grouping
     , TRIM(activity)                                                           activity_code
     , TRIM(activity)                                                           accountingelement1
     , TRIM(account)                                                            account_code
     , TRIM(account)                                                            accountingelement2
     , TRIM(event)                                                              event_code
     , TRIM(event)                                                              accountingelement3
     , TRIM(commercial)                                                         commercial_code
     , TRIM(commercial)                                                         accountingelement4
     , TRIM(net_value)                                                          net_value_gbp_text
     , TRY_CAST(NULLIF(TRIM(net_value), '') AS DOUBLE)                          net_value_gbp
     , TRIM(commitmentdate)                                                     commitment_date_excel
     , TRIM(orderedquantity)                                                    ordered_quantity_text
     , TRY_CAST(NULLIF(TRIM(orderedquantity), '') AS DOUBLE)                    ordered_quantity
     , TRIM(acceptedquantitytotal)                                              accepted_quantity_total_text
     , TRY_CAST(NULLIF(TRIM(acceptedquantitytotal), '') AS DOUBLE)             accepted_quantity_total
     , TRIM(closedforreceiptingquantity)                                        closed_for_receipting_quantity_text
     , TRY_CAST(NULLIF(TRIM(closedforreceiptingquantity), '') AS DOUBLE)      closed_for_receipting_quantity
     , TRIM(fullyinvoicedquantity)                                              fully_invoiced_quantity_text
     , TRY_CAST(NULLIF(TRIM(fullyinvoicedquantity), '') AS DOUBLE)             fully_invoiced_quantity
     , TRIM(closedinvoicequantity)                                              closed_invoice_quantity_text
     , TRY_CAST(NULLIF(TRIM(closedinvoicequantity), '') AS DOUBLE)             closed_invoice_quantity
     , TRIM(invoice_status)                                                     invoice_status
     , TRIM(delivery_status)                                                    delivery_status
     , TRIM(unit_value)                                                         unit_value_text
     , TRY_CAST(NULLIF(TRIM(unit_value), '') AS DOUBLE)                        unit_value
     , TRIM(soft_commitment_value)                                              soft_commitment_value_gbp_text
     , TRY_CAST(NULLIF(TRIM(soft_commitment_value), '') AS DOUBLE)           soft_commitment_value_gbp
     , TRIM(hard_commitment_value)                                              hard_commitment_value_gbp_text
     , TRY_CAST(NULLIF(TRIM(hard_commitment_value), '') AS DOUBLE)           hard_commitment_value_gbp
     , TRIM(currency)                                                           currency
     , TRIM(soft_commitment_value_1)                                            soft_commitment_value_local_text
     , TRY_CAST(NULLIF(TRIM(soft_commitment_value_1), '') AS DOUBLE)         soft_commitment_value_local
     , TRIM(hard_commitment_value_1)                                            hard_commitment_value_local_text
     , TRY_CAST(NULLIF(TRIM(hard_commitment_value_1), '') AS DOUBLE)         hard_commitment_value_local
     , TRIM(period)                                                             reporting_period
     , TRIM(activity) || '-' || TRIM(account) || '-' || TRIM(event) || '-' || TRIM(commercial)  account_string
     , TRIM(companyshortcode) || '|' || TRIM(order_number) || '|' || TRIM(line_number)          commitment_stack_line_key
  FROM fin_supplier_po_commitment_stack_report
 WHERE TRIM(COALESCE(order_number, '')) <> ''
;
