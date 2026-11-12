# Overview
This describes the work done for mapping pff/crm customers to workday. covering the updating the workday customer id, and for pff the process that produces an output to then update/create the pff record in workday, and for crm the outputs produced to update the workday-customer-id to match their crm-customer-id (as how the interface built will expect it)

ANd how it supports the case when the interfacer goes live

and what the risks are.


note in terms of correctness of the solution, it can be argued that the solution is correct in thaat it is using the data ptrovided by the source systems as the source-of-truth - and hence an additional reconcile would be a mote -point (but word this properly)

It also establishes how the intefaces will behave when live and thus the fact that what has been done here in shaping the data is correct for what the interfaces expect

# Ask
I need this creating into a simple powerpoint deck showing what has been built and the guardrails in place to ensure the outputs from this process are correct. And the standing assumptions of how the interfaces when live are meant to act


# Implementation
under:
implementations/elt_ingest_excel/config/transform/sql/finance/customer

there are the transforms for handling the migration of data from PFF and CRM


implementations/elt_ingest_excel/config/ingest/finance/active_customer_debtor.json
describes the inputs that are initially created to create the source-input data


I need to document the approach for a plan of action.

fin_workday_extract_customer - the latest customer extract from workday

pff_gp_discipline_customer - the mapping of the GP customer to on-side-org-id and workday-customer-id

pff_pro_club - the mapping of on-side-org-id to club-id.  

pff_workday_manual_mapped - an additional list of pff-club-id to workday-customer-id as built/defined by finance.  This is ultimately used to define any additional mappings not covered by pff_gp_discipline_customer, and also to provide overrides against pff_gp_discipline_customer (altjhough this is not yet developed)

crm_account - crm customers that are organisations

crm_contact - crm customers that are individuals

crm_workday_customer_manual - manually created customers in workday

note in the case of crm - it handles duplicate mappings too.  see thew sql.

In terms of outputs:
workday_pff_os_customer_load - these are the customers that need to be updated against or created in workday based on their PFF data.

And then additionally there is this script:
implementations/elt_ingest_excel/scripts/workday_customer_crm_org_fan_update.py

That creates the update EIBs that update the workday customer id to be the crm customer id.  and then for crm customers thagt are orgs there is an additional update file for setting the FAN of the contact on thje crm-customer that is an organisation

so the above is about getting the data conformed in workday such that anhy missing customers are created (in the case of PFF), and that the workday-customer-id is updated to the PFF Club ID or the CRM customer ID (which in the case of a organisation is the account-id, and in the case of an individual is their FAN)


Now when the interfaces are live:
1. For PFF it guarantees that the customer exists in workday.  So that when invoicing needs to be done, the customer exists.  Note here the PFF customer created in workday has the pff club id prefixed with 'OS-'  and the inteface has ben designed the same to also add 'OS-' as a prefix when sending the workday-customer-id for the invoice payment.
2. For CRM, the interface will first check if the customer exists, on the get failing it will create the customer.  So in this instance we ensure the interface does not create duplicates, as we account for the customer in workday on its workday customer id bneing correctly set to the crm customer account-id or for the individual its FAN.  And the interface is meant to be designed to behave the same way

Now some gaps/risks to capture:
1. The GP migration into Workday resulted in duplicate customers.  
2. The mappings that have been created account for duplication as per the SQL that has been developed.
3. Finance AR team have marked PFF customers in workday as inactive, and their manual mapping is meant to show the correct workday customer to map to instead - this has to be developed
4. The list of manually created crm customers in workday wont get updated to set their workday-customer-id to the correct crm-id (see: crm_workday_customer_manual), and these need to be reviewed to see if they can be mapped to the crm extract on their crm-license number.  And if that does not create a mapping, then they need creating in  CRM, and then am additional process built to create the workday-customer-id to crm-customer-id (accountid or fan) mapping for 3+ to then update.





