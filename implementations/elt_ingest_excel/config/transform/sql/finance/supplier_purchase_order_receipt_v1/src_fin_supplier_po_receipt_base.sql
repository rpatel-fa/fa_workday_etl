DROP TABLE IF EXISTS src_fin_supplier_po_receipt_base;

CREATE TABLE src_fin_supplier_po_receipt_base
    AS
  WITH cte_src_receipt_base
    AS (
SELECT TRIM(company)                                                          company_shortcode
     , TRIM(pono)                                                             order_number
     , TRIM(orderdate)                                                         order_date_excel
     , TRIM(orderlinenumber)                                                   line_number
     , TRIM("position")                                                        po_position
     , TRIM(receipt_number)                                                   receipt_number
     , TRIM(proactis_receipt_no)                                              proactis_receipt_no
     , TRIM(receipted_quantity)                                              receipted_quantity_text
     , TRY_CAST(NULLIF(TRIM(receipted_quantity), '') AS DOUBLE)                  receipted_quantity
     , TRIM(invoicevalue)                                                    invoice_value_text
     , TRY_CAST(NULLIF(TRIM(invoicevalue), '') AS DOUBLE)                    invoice_value
     , TRIM(nominal_net_value)                                              nominal_net_value_gbp_text
     , TRY_CAST(NULLIF(TRIM(nominal_net_value), '') AS DOUBLE)              nominal_net_value_gbp
     , TRIM(netvalue_original_curr)                                    net_value_original_curr_text
     , TRY_CAST(NULLIF(TRIM(netvalue_original_curr), '') AS DOUBLE)      net_value_original_curr
     , TRIM(currency)                                                         currency
     , TRIM(receipt_date)                                                  receipt_date_excel
     , TRIM(receipt_created_date)                                receipt_created_date_excel
     , TRIM(invoicestatus)                                                  invoice_status
     , TRIM(deliverystatus)                                                delivery_status
     , TRIM(delivery_reason)                                                delivery_reason
     , TRIM(order_closed_for_invoicing)                        order_closed_for_invoicing
     , TRIM(order_closed_for_invoicing_reason)      order_closed_for_invoicing_reason
     , TRIM(order_item_cancellation_status)              order_item_cancellation_status
     , 'PO Receipt Extract'                                                  nrm_source_tab
  FROM fin_supplier_po_receipt                                                 t
       )
     , cte_joined
    AS (
SELECT t.*
     ,  company_shortcode
     || '|' || order_number
     || '|' || line_number
     || '|' || po_position
     || '|' || receipt_number
     || '|' || proactis_receipt_no
     || '|Q|' || receipted_quantity_text                                           receipt_line_key
  FROM cte_src_receipt_base                                                    t
 WHERE TRIM(COALESCE(order_number, '')) <> ''
       )
SELECT t.*
  FROM cte_joined                                                               t
;
