# Master Customer / Supplier setup (AR / AP create-helper)

The entry point for this capability. Start here, then follow the links for detail.

## What it is

A helper for AR and AP to **create** a customer or supplier in Workday:

1. The user fills a one-row-per-entity **input template**, an Excel workbook with drop-downs, from
   [`config/ingest_specs/ar_customer.yaml`](../config/ingest_specs/ar_customer.yaml) or
   [`ap_supplier.yaml`](../config/ingest_specs/ap_supplier.yaml).
2. In the Runner they pick **AR Customer Master** or **AP Supplier Master** and click Run Transform.
3. The transform clones the 3+ EIB `Master Customers.xlsx` / `Master Suppliers.xlsx` and fills it.
   AR or AP then upload that file into Workday.

Anything beyond the basic record is added in Workday **after** the customer or supplier exists.

## Rules (operator decisions, 2026-10-12)

The specs' `decisions:` blocks record these rules with dates, for example
`single_value_per_area_address_lines_spill` in `ar_customer.yaml` / `ap_supplier.yaml`. Quotes below are
the operator's own words from the 2026-10-12 build session. Do not widen these rules without the
operator.

- **1 address, 1 email, 1 phone number, 1 bank account** per customer or supplier. Customers have no
  bank-account area in v1; suppliers must provide one or they are rejected. Operator: "if they
  want to add additional addresses, phonenumbers, emails, etc, then they do it in workday - once the
  customer is created".
- **One EIB row per entity, except address lines.** Each non-blank Address Line 1/2/3 goes on its own
  EIB row, as in the 3+ sample: the address Row ID stays 1, the line Row ID is k, and the type is
  `ADDRESS_LINE_k`. Operator: "only address - ... for addresses - it means spilling to multiple lines
  since an address is multipe address lines".
- **Single values** everywhere else:
  - Use For is `BILLING`.
  - Type is `BUSINESS`.
  - Group is the automatic A-L / M-Z / 0-9 band, assigned from the name.
  - For suppliers, Payment Types Accepted is the Default Payment Type.
- **IDs are typed by users and never generated.** This differs from `elt_ingest_excel`. Operator:
  "we wont entertain auto generating the workday id for the custoer/suplier - they have to present
  them".
- **Reference ID is optional and passed through as typed.** Operator: "we simply passthrough to the
  EIBs".
- **Workday values.** Drop-down display values are translated to Workday reference IDs (for example
  `Net 30` becomes `NET_30`). A value with no reference ID rejects the **whole** entity to
  `<output>.rejects.csv`. There is no fallback or guessing.
- **The 3+ sample is the authority** for fixed values. Add Only and Submit are fixed in both
  transforms. For suppliers only, Approval Status is `APPROVED`, Address Format is `BASIC` and
  Effective Date is `1900-01-01`.

## Where things are

All paths are under `implementations/elt_transform_excel/`.

| Piece | AR Customer | AP Supplier |
|---|---|---|
| Input template spec (columns, drop-downs) | `config/ingest_specs/ar_customer.yaml` (36 columns) | `config/ingest_specs/ap_supplier.yaml` (46 columns) |
| Transform profile (Runner entry, paths, letter map) | `config/transforms/ar_customer_master.yaml` | `config/transforms/ap_supplier_master.yaml` |
| Mapping and decision record | `config/transform_specs/ar_customer_master.yaml` | `config/transform_specs/ap_supplier_master.yaml` |
| VBA module | `vba_src/transforms/modArCustomerMaster.bas` | `vba_src/transforms/modApSupplierMaster.bas` |
| Target EIB (3+) | `Master Customers.xlsx`, sheet `Submit Customer`, data row 6 | `Master Suppliers.xlsx`, sheet `Submit Supplier`, data row 6 |

The two transforms share these pieces:

- **Value lists:** `config/ingest_lookups/*.json`. They drive the template drop-downs and, at run
  time, the reference-ID translation.
- **Runtime lookups sheet:** `scripts/build_runner.py` writes the lists into a hidden `lookups` sheet
  in `Runner.xlsm`. VBA reads it through `modHelpers.LookupRuntimeValue`.
- **Address-line rows:** `vba_src/modEntityBlockWriter.bas`, registered as `entity_row_block` in
  `config/shared_helper_registry.yaml`.
- **EIB types:** `Customer` and `Supplier` in `config/eib_population_defaults.yaml` and in the
  build gates.

## How to build, run and change it

- **Render the templates and build the Runner:** [BUILD.md](BUILD.md), and
  [BUILD_WIN.md](BUILD_WIN.md) on Windows.
- **Run the transforms:** [RUNBOOK.md](RUNBOOK.md), sections "Running AR Customer Master" and
  "Running AP Supplier Master".
- **Spec and lookup format:** [CONFIGURATION.md](CONFIGURATION.md), section "Ingest Specs".
- **Changes:**
  - A drop-down value or a new reference ID is a lookup JSON edit.
  - A column change is a spec edit, plus the matching profile, spec and `.bas` edits for the
    transform. That is the same pattern as the invoice transforms; see
    [ADDING_A_TRANSFORM.md](ADDING_A_TRANSFORM.md).
  - Re-render the template after either kind of change.

## Status (2026-10-12)

Code-complete and merged (PRs #33 and #34). It passes the Python build gates, and Python reference
simulations of both transforms pass. **Neither transform has been run in Excel or uploaded to
Workday yet.**

## Next steps

1. Do the manual Excel import and `build_runner.py --verify` ([BUILD.md](BUILD.md)).
2. Run both Runner entries on a filled-in template. Compare the output with the 3+ sample rows:
   rows 6 and 9 for customers, rows 6 and 10 for suppliers.
3. Upload both outputs to the Workday test tenant.
4. Refine the templates and the mapping with AR and AP.

## Open items to settle on the first upload

- **Supplier Status:** the transform writes `ACTIVE`, as the 3+ sample does, but the tenant list's
  reference ID is `Active`.
- **Values with no Workday reference ID.** These reject the entity only if a user picks them:
  - 110 of 2,522 regions (for example `AE|Sharjah`, `GB|Manchester`)
  - the `Print` delivery method
  - ISO3 for Puerto Rico and Kosovo, which is needed for phones
  - 9 customer and 4 supplier tax-ID-type rows
- **Not loaded by design** (documented in the specs): Tax Code, Accepted Currencies, Email Use For and
  the Public flags.
- **Blank middle address line:** lines are renumbered contiguously. The 3+ sample gives no evidence
  either way.
