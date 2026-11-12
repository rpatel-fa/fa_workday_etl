# Workday customer conforming — PFF & CRM

Stakeholder deck: https://claude.ai/artifact/MqfaxudcTdyaRGWiBTm2kn

## What this is

When the PFF and CRM interfaces are switched on, each one identifies Workday customers by the
customer ID held on its own side. Many of these customers are already in Workday, but under
different IDs. The two interfaces react differently when the ID doesn't match:

- **CRM** searches Workday for the customer and, if it isn't found, **creates a new one**, which
  is a duplicate of the customer already there under a different ID.
- **PFF** is an event bus. It sends customer creates, customer updates and invoices to Workday
  **as they occur** in PFF. It **can't replay past events**, so clubs that already exist must
  already be in Workday under their PFF ID at go-live. If one isn't, that club's invoices and
  updates can't be applied.

This work lines up the Workday customer ID with the PFF and CRM customer IDs before go-live:

- **PFF** — every PFF club exists in Workday with Customer ID **`OS-<PFF Club ID>`**. Clubs not yet
  in Workday are created; existing ones have their ID updated.
- **CRM** — existing Workday customers have their Customer ID updated to the CRM ID: the
  **Account ID** for organisations and the **FAN** for individuals. Organisations also get their
  primary contact's FAN recorded. **Scope is CRM customers already in Workday only.** CRM
  customers not yet in Workday are left for the CRM interface to create when it goes live.

After go-live, both interfaces take over: new customers and changes flow through them as normal.

## What's at stake

- **CRM: duplicate customers.** If an existing customer's Workday ID doesn't match its CRM ID, the
  interface won't find it and will create a second one. Invoices and payments then split across two
  records for the same person or organisation, and statements, credit control and reporting go
  wrong. Duplicates are much harder to unpick after go-live than to prevent now.
- **PFF: invoices and updates rejected.** If an existing club isn't in Workday as `OS-<Club ID>` at
  go-live, its invoices and customer updates have nowhere to go. The interface can't fix this
  itself, because it sends events as they occur and can't replay past events.

**This work is the fix for both risks.** It lines up every Workday customer ID with the ID held
by the system that owns the customer (PFF or CRM) before go-live. CRM then finds the existing
customer instead of creating a duplicate, and every PFF invoice and update has a customer to land
on.

If a problem still happens after go-live, it has one of three causes:

1. **The source data** was wrong or incomplete.
2. **This work** shaped the data wrongly.
3. **The interface** doesn't behave as its builders said it would.

This document is the agreement that separates those three. It says what this work delivers, what
it relies on, and what it expects of the interfaces. If something goes wrong after go-live, the
checks below show which of the three caused it, so the issue goes straight to the right owner
instead of being argued over.

## Source files

Everything is built from these files and nothing else. Each one comes from the system that owns
the data.

| From | File | What it provides |
|---|---|---|
| Workday | `Extract_Customers_20261027.xlsx` | Every customer in Workday as at 27 Oct 2026, with its ID and licence numbers |
| PFF (GP) | `Discipline Customers.xlsx` | Links each GP customer to its OnSide org ID and Workday customer ID |
| PFF | `Pro club details Extract[76].xlsx` | Every PFF club: Club ID, OnSide org ID, name, address, secretary |
| Finance | `pff_workday_manually_mapped_customer.xlsx` | Finance's own PFF Club → Workday customer links, where GP gives none |
| CRM | `Account Advanced Find View (with Licences).xlsx` | CRM organisations: Account ID, licence numbers, primary contact FAN |
| CRM | `Contact Advanced Find View (with Licences).xlsx` | CRM individuals: FAN, licence numbers |
| Finance | `Manual EBW Invoices.xlsx` | CRM customers that finance created by hand in Workday |

The CRM update also uses two Workday load templates: `Master Customer EBW CRM ID Update.xlsx` and
`FAN_Number CRM Update.xlsx`.

## What gets produced

| Output | Purpose |
|---|---|
| PFF club load sheet | One row per club: **MAPPED** (update the existing Workday customer) or **UNMAPPED** (create it). Flags missing address, email or mandatory fields. |
| `workday_crm_id_update.xlsx` | Workday load file that changes each matched customer's ID to its CRM ID. |
| `workday_customer_crm_org_fan_update.xlsx` | Workday load file that records the primary contact FAN on organisations whose ID was updated. Loaded **after** the ID update. |
| `workday_crm_id_update_stats.xlsx` | Counts and named lists of every customer *not* updated, and why. |
| Unmapped CRM lists | CRM organisations and individuals with a licence number but no matching Workday customer. |

## Why the results can be trusted

The results are built entirely on the sources of truth: the files from Workday, PFF, CRM and
Finance listed above. Nothing is invented or changed along the way, and every result can be traced
back to the source row it came from.

## The agreement

### What this work delivers

1. **PFF:** every in-scope club in the PFF club extract appears on the load sheet, either matched to
   an existing Workday customer or marked for creation, with ID `OS-<Club ID>`. Test clubs and
   clubs with no OnSide org ID are out of scope.
2. **CRM:** every CRM customer **already in Workday** (matched on licence number) gets its Workday
   ID set to its Account ID or FAN. CRM customers not in Workday are not in scope; the CRM
   interface creates them.
3. **No double-ups:** each Workday customer receives one CRM ID, and each CRM ID goes to one Workday
   customer. Where the GP migration left duplicates, a fixed rule picks the one to keep.
4. **Nothing silently dropped:** any customer that can't be matched or updated is named in an
   output, with the reason.
5. **Repeatable:** running it again on the same files gives the same result.

### What this work relies on

- The source files are accurate and complete as at their extract date. Mappings in them (GP
  discipline file, finance's manual mapping) are taken as correct.
- The licence number is the shared key between CRM and Workday.
- The load files are loaded into Workday as produced, in the order stated.
- Customers or clubs created **after** the extract dates are not covered. Refresh the extracts and
  re-run close to go-live.

### How the interfaces behave

This is how each interface works, as designed by its builders. The data has been shaped to match
it exactly. Behaviour that differs from this is an interface issue, not a data issue.

| Interface | Expected behaviour |
|---|---|
| **PFF** | Sends customer creates, customer updates and invoices from the PFF event bus to Workday as they occur in PFF, using Customer ID `OS-<PFF Club ID>` and adding the `OS-` prefix itself. **Can't replay past events**, so existing clubs must already be in Workday under that ID at go-live. |
| **CRM** | Searches Workday for the customer by Customer ID = Account ID (organisation) or FAN (individual). Creates a customer **only** if that search finds nothing. |

## If something goes wrong after go-live

| What you see | Check | Points to |
|---|---|---|
| PFF invoice or customer update rejected: customer not found | Is the club on the load sheet, and was it loaded? | Yes → the interface sent a different ID (check the `OS-` prefix). Out of scope or missing from the club extract → source data. |
| CRM creates a duplicate customer | Is the CRM ID in the ID-update file, and was it loaded? | Yes → the interface search is not working as agreed. On an unmapped list or the manual list → known gap below. Not in the CRM extract, or no licence number → source data. |
| Customer linked to the wrong record | Trace the row back to its source file | Wrong in the source → source data. Right in the source but shaped wrongly → this work. |

## Known gaps and risks

| # | Item | Status | Detail |
|---|---|---|---|
| 1 | Inactive PFF duplicates | To build | The way clubs were set up in GP meant the migration couldn't tell that some were the same club, so it created them in Workday more than once. Finance AR has marked the extras inactive and, in its manual mapping, named the active customer to use. The process doesn't apply that yet, so an affected club can still map to an inactive customer, and its invoices would land on a record Finance has retired. |
| 2 | CRM customers created by hand in Workday | To build | Customers Finance created by hand in Workday for CRM customers (`Manual EBW Invoices.xlsx`) are not covered, so the CRM interface won't find them and will create duplicates. Plan: match them to CRM on licence number; any not in CRM are created there first, then a Workday ID → Account ID / FAN update is built for 3+ to load. |
| 3 | GP migration duplicates (CRM) | Handled | The GP migration also created duplicate customers for CRM. The CRM mapping collapses them to one Workday customer per CRM ID. |
| 4 | Changes after the extract dates | Manage at cutover | Anything created or changed after the extracts isn't covered. Refresh the extracts and re-run before go-live. |
| 5 | CRM records without a licence number | Out of scope | They can't be matched to Workday. If one already exists in Workday, the interface will create a duplicate. |

---

## For maintainers

### Where the logic lives

| What | Path |
|---|---|
| Ingest of the source files | [config/ingest/finance/active_customer_debtor.json](../config/ingest/finance/active_customer_debtor.json) |
| PFF transforms | `config/transform/sql/finance/customer/workday_pff_os_customer_{mapped,unmapped,load}.sql` |
| CRM transforms | `config/transform/sql/finance/customer/src_crm_customer_base.sql`, `workday_crm_customer_{mapped,mapped_clean,mapped_wd,unmapped_account,unmapped_contact}.sql` |
| CRM load files + stats | [scripts/workday_customer_crm_org_fan_update.py](../scripts/workday_customer_crm_org_fan_update.py) |

The detailed rules (tie-breaks, exclusions, flags, stats balances) are in those files.

### Rebuilding the stakeholder deck

The audience is finance and project stakeholders, plus the interface owners. Nobody reviews the
code, so the deck must never show SQL, table names or repo paths, only source and output file
names. Update the **existing** deck (link at top) in place. All content comes from this doc, except
slides 6–9, which are written from the files under *Where the logic lives*.

Slide outline (keep the order; change only the slides whose source changed):

1. Cover
2. What this is: the three outcomes
3. What's at stake: CRM creates duplicates for customers it can't find; PFF can't replay past events, so existing clubs missing from Workday have their invoices/updates rejected; the three possible causes, and why an agreement
4. Source files: *Source files* table
5. Built · PFF: map via GP → finance links → confirm in Workday → load sheet
6. Built · CRM: licence-number match → one per CRM customer → one per Workday customer → unmapped lists
7. Built · CRM load files: ID update, organisation FAN, stats report
8. Guardrails: PFF / CRM / load files
9. Why the results can be trusted: built entirely on the sources of truth, traceable to source rows (make no claim about testing or reconciliation)
10. The agreement: delivers / relies on / interfaces do
11. Interface behaviour: the *How the interfaces behave* table
12. If something goes wrong: the trace table
13. Known gaps and risks
14. Next steps
