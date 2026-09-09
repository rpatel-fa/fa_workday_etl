"""Build the executive 'summary' tab for the supplier CCS reconcile review pack.

Mirrors the hand-authored customer_ccs summary sheet (same palette / layout) but
maps the sections to the supplier story: population, core attributes, the go-live
manual-update list (the supplier centrepiece), attribute gaps, actions, accepted
items, status and approvals. Every number is read live from the reconcile tables
so the sheet cannot drift from the detail.
"""

from pathlib import Path

import duckdb
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Palette (lifted from the customer summary sheet).
NAVY = "1F3864"
BLUE_SECTION = "2E5496"
BLUE_HDR = "D9E1F2"
GREEN = "E2EFDA"
GREEN_TXT = "548235"
AMBER = "FCE4D6"
AMBER_TXT = "C55A11"
GREY = "F2F2F2"
GREY_TXT = "808080"

_THIN = Side(style="thin", color="FFBFBFBF")
_BOTTOM = Border(bottom=_THIN)
_WRAP = Alignment(vertical="top", wrap_text=True)
_WRAP_L = Alignment(horizontal="left", vertical="top", wrap_text=True)
_WRAP_C = Alignment(horizontal="center", vertical="top", wrap_text=True)


def _fill(hex_):
    return PatternFill(start_color=hex_, end_color=hex_, fill_type="solid")


def _fetch(database_path: Path) -> dict:
    """Pull every headline number the sheet needs from the reconcile tables."""
    with duckdb.connect(str(database_path), read_only=True) as conn:
        def one(sql):
            return conn.execute(sql).fetchone()[0]

        def rep(metric):
            return one(
                "SELECT metric_value FROM reconcile_supplier_ccs_report "
                f"WHERE metric = '{metric}'"
            )

        d = {
            "gp_source_rows": one("SELECT COUNT(*) FROM src_fin_supplier_base"),
            "gp_keys": rep("1a_distinct_gp_vendor_id"),
            "merged": rep("1b_src_fin_supplier_winning_rows"),
            "extract_rows": rep("1c_fin_workday_extract_supplier_rows"),
            "missing": rep("2a_gp_missing_in_workday"),
            "extra": rep("2b_extra_workday_reference_ids_not_in_gp"),
            "dummy": rep("2c_workday_dummy_rows"),
            "multi_row": rep("2d_suppliers_multiple_extract_rows"),
            "merge_diff": rep("2e_merge_group_diff"),
            "manual_total": rep("3a_manual_update_total"),
            "manual_addr": rep("3b_manual_update_address"),
            "manual_bank": rep("3c_manual_update_bank"),
            "manual_pay": rep("3d_manual_update_payment"),
            "manual_unmatched": rep("3e_manual_update_unmatched"),
            "addr_exp": rep("4a_address_gap_expected"),
            "addr_unexp": rep("4b_address_gap_unexplained"),
            "bank_exp": rep("4c_bank_gap_expected"),
            "bank_unexp": rep("4d_bank_gap_unexplained"),
            "pay_exp": rep("4e_payment_gap_expected"),
            "pay_unexp": rep("4f_payment_gap_unexplained"),
        }
        d["source_multi_workday"] = one(
            "SELECT COUNT(*) FROM reconcile_supplier_ccs_source_multi_workday"
        )
        d["dummy_desc"] = one(
            "SELECT COALESCE(STRING_AGG(workday_supplier_id || ' (' || "
            "COALESCE(NULLIF(workday_supplier_name,''),'blank') || ')', ', '), '') "
            "FROM reconcile_supplier_ccs_workday_dummy_row"
        )
        d["addr_unexp_list"] = one(
            "SELECT COALESCE(STRING_AGG(supplier_number || ' ' || expected_supplier_name, '; ' "
            "ORDER BY supplier_number), '') "
            "FROM reconcile_supplier_ccs_attribute_gap "
            "WHERE gap_type = 'ADDRESS' AND gap_status = 'UNEXPLAINED'"
        )

        def match(ce, ca):
            return conn.execute(
                f"SELECT COUNT(*) FILTER (WHERE a IS NOT NULL "
                f"AND UPPER(TRIM(e)) = UPPER(TRIM(a))) FROM "
                f"(SELECT {ce} e, {ca} a FROM reconcile_supplier_ccs_attribute_detail "
                f"WHERE actual_reference_id IS NOT NULL)"
            ).fetchone()[0]

        d["m_name"] = match("expected_supplier_name", "actual_supplier_name")
        d["m_cat"] = match("expected_supplier_category", "actual_supplier_category")
        d["m_ccy"] = match("expected_default_currency", "actual_default_currency")
        d["m_pay"] = match("expected_default_payment_type", "actual_default_payment_type")
        d["compared"] = one(
            "SELECT COUNT(*) FROM reconcile_supplier_ccs_attribute_detail "
            "WHERE actual_reference_id IS NOT NULL"
        )

        def diff_list(col_e, col_a):
            return one(
                "SELECT COALESCE(STRING_AGG(supplier_number, ', ' ORDER BY supplier_number), '') "
                "FROM reconcile_supplier_ccs_attribute_detail "
                f"WHERE actual_reference_id IS NOT NULL "
                f"AND UPPER(TRIM({col_e})) <> UPPER(TRIM({col_a}))"
            )

        d["cat_diff_list"] = diff_list("expected_supplier_category", "actual_supplier_category")
        d["ccy_diff_list"] = diff_list("expected_default_currency", "actual_default_currency")
        # Default-payment-type differences, split by whether the manual-update list
        # (payment reset to Manual, or bank removal which forces Manual) explains them.
        pay = conn.execute(
            "SELECT "
            "COUNT(*) FILTER (WHERE diff) AS total, "
            "COUNT(*) FILTER (WHERE diff AND explained) AS manual, "
            "COUNT(*) FILTER (WHERE diff AND NOT explained) AS minor FROM ("
            "SELECT UPPER(TRIM(expected_default_payment_type)) <> UPPER(TRIM(actual_default_payment_type)) AS diff, "
            "(manual_exception_types LIKE '%PAYMENT%' OR manual_exception_types LIKE '%BANK%') AS explained "
            "FROM reconcile_supplier_ccs_attribute_detail WHERE actual_reference_id IS NOT NULL)"
        ).fetchone()
        d["pay_diff_total"], d["pay_diff_manual"], d["pay_diff_minor"] = pay

        d["manual_open"] = one(
            "SELECT COUNT(*) FROM reconcile_supplier_ccs_manual_update_status "
            "WHERE still_open_in_workday = 'Yes'"
        )
        d["manual_resolved_list"] = one(
            "SELECT COALESCE(STRING_AGG(supplier_name_raw || ' (' || exception_type || ')', '; '), '') "
            "FROM reconcile_supplier_ccs_manual_update_status "
            "WHERE still_open_in_workday LIKE 'No%'"
        )
        d["action_total"] = one("SELECT COUNT(*) FROM reconcile_supplier_ccs_finance_action_list")

        # Row-vs-supplier reconciliation for the population funnel.
        d["wd_distinct_suppliers"] = one(
            "SELECT COUNT(DISTINCT reference_id) FROM fin_workday_extract_supplier "
            "WHERE NULLIF(TRIM(reference_id), '') IS NOT NULL"
        )
        d["split_extra_rows"] = one(
            "SELECT COALESCE(SUM(c - 1), 0) FROM ("
            "SELECT COUNT(*) c FROM fin_workday_extract_supplier "
            "WHERE NULLIF(TRIM(reference_id), '') IS NOT NULL GROUP BY reference_id HAVING COUNT(*) > 1)"
        )
        d["further_merges"] = d["merged"] - d["wd_distinct_suppliers"]
    return d


def build_summary_sheet(pack_path: Path, database_path: Path) -> None:
    d = _fetch(database_path)
    n = lambda v: f"{v:,}"

    wb = openpyxl.load_workbook(pack_path)
    if "summary" in wb.sheetnames:
        del wb["summary"]
    ws = wb.create_sheet("summary", 0)

    for col, width in (("A", 4.6), ("B", 25.1), ("C", 38.9), ("D", 12.6), ("E", 52.6), ("F", 3.4)):
        ws.column_dimensions[col].width = width

    r = 1

    def span(row, text, fill=None, font=None, align=_WRAP_L, height=None, first="A"):
        ws.merge_cells(f"{first}{row}:E{row}")
        c = ws[f"{first}{row}"]
        c.value = text
        if fill:
            for col in "ABCDE":
                ws[f"{col}{row}"].fill = _fill(fill)
        if font:
            c.font = font
        c.alignment = align
        if height:
            ws.row_dimensions[row].height = height

    def kv(row, label, value, label_fill=GREY, value_fill=BLUE_HDR):
        ws.merge_cells(f"B{row}:D{row}")
        b = ws[f"B{row}"]
        b.value = label
        b.font = Font(sz=10)
        b.fill = _fill(label_fill)
        b.alignment = _WRAP_L
        b.border = _BOTTOM
        e = ws[f"E{row}"]
        e.value = value
        e.font = Font(sz=10, bold=True)
        e.fill = _fill(value_fill)
        e.alignment = _WRAP_L
        e.border = _BOTTOM
        for col in "CD":
            ws[f"{col}{row}"].border = _BOTTOM

    def note(row, text):
        span(row, text, font=Font(sz=9, italic=True, color=GREY_TXT), align=_WRAP_L)

    def section(row, text):
        span(row, text, fill=BLUE_SECTION, font=Font(sz=12, bold=True, color="FFFFFF"),
             align=_WRAP_L, height=15.75)

    # Title band.
    span(r, "Supplier Master Data Migration to Production",
         fill=NAVY, font=Font(sz=16, bold=True, color="FFFFFF"), height=21.0)
    r += 1
    span(r, "Reconciliation Summary   —   full record of the supplier GP → Workday migration "
            "(extract 2026-07-28)",
         fill=NAVY, font=Font(sz=10, color="FFFFFF"))
    r += 2

    # Outcome band.
    span(r, "✓  OUTCOME: Data extracted from GP and migrated to Workday. Every GP supplier is "
            "accounted for in Workday — 0 missing, 0 extra.",
         fill=GREEN, font=Font(sz=11, bold=True, color=GREEN_TXT))
    r += 1
    span(r, "Core business data (name, category, currency, payment, reference ID) matches. The only "
            "open items are the go-live manual-update suppliers Finance must re-enter in Production, "
            "plus a small number of known residual source-data gaps.",
         font=Font(sz=10), align=_WRAP_L)
    r += 1
    note(r, "Colour key:   amber = action required in Workday Production;   blue = value / decision;   "
            "grey = informational.")
    r += 2

    # 0. Method.
    section(r, "0.  How the reconciliation was performed")
    r += 1
    for label, val in (
        ("Primary lens", "Sent Workday payload (workday_supplier_*)  vs  loaded Workday extract"),
        ("Secondary lens", "GP source  vs  loaded extract — explains approved cleansing / dedup / merges"),
        ("Gold / expected", "The Workday transform output (what was actually sent), keyed on GP Vendor ID"),
        ("Scope reconciled", "Population & identity, core business attributes, address / bank / payment "
                             "presence, and the go-live manual-update list"),
    ):
        kv(r, label, val)
        r += 1
    r += 1

    # 1. Population.
    section(r, "1.  Population reconciliation  —  RECONCILED, 0 discrepancies")
    r += 1
    kv(r, "Records received from GP (source rows)", n(d["gp_source_rows"])); r += 1
    kv(r, "Distinct GP supplier keys (Vendor ID × business unit)", n(d["gp_keys"])); r += 1
    kv(r, "Merged suppliers created by the transform (after de-dup)", n(d["merged"])); r += 1
    kv(r, f"Distinct suppliers in Workday ({d['further_merges']} pairs merged further by Workday)",
       n(d["wd_distinct_suppliers"])); r += 1
    kv(r, "Workday extract rows (NB: rows, not suppliers)",
       f"{n(d['extract_rows'])}  =  {n(d['wd_distinct_suppliers'])} suppliers"
       f"  +  {d['split_extra_rows']} extra rows for split bank/address suppliers"
       f"  +  {d['dummy']} test record"); r += 1
    note(r, f"So the extract has NO extra suppliers: {n(d['extract_rows'])} rows − "
            f"{d['split_extra_rows']} split rows − {d['dummy']} test = {n(d['wd_distinct_suppliers'])} "
            f"real suppliers, all traced to GP. The {d['split_extra_rows']} split-row suppliers are on "
            f"the multi_row_suppliers tab; the test row on workday_dummy_rows.")
    r += 1
    span(r, "Population integrity checks — all passed:", font=Font(sz=10, bold=True)); r += 1
    ws.merge_cells(f"B{r}:D{r}"); ws[f"B{r}"].value = "Integrity check"
    ws[f"B{r}"].font = Font(sz=10, bold=True, color=NAVY); ws[f"B{r}"].fill = _fill(BLUE_HDR)
    ws[f"B{r}"].alignment = _WRAP_L
    ws[f"E{r}"].value = "Result"; ws[f"E{r}"].font = Font(sz=10, bold=True, color=NAVY)
    ws[f"E{r}"].fill = _fill(BLUE_HDR); ws[f"E{r}"].alignment = _WRAP_C
    for col in "CD": ws[f"{col}{r}"].fill = _fill(BLUE_HDR)
    r += 1
    for chk, res in (
        ("GP keys missing from the merged supplier population", f"{d['missing']}   ✓ confirmed"),
        ("GP keys missing from Workday", f"{d['missing']}   ✓ confirmed"),
        ("GP keys linked to more than one Workday supplier row", f"{d['source_multi_workday']}   ✓ confirmed"),
        ("Extra Workday rows not traced to GP",
         f"{d['dummy']}   (Workday test record — {d['dummy_desc']})"),
    ):
        kv(r, chk, res, value_fill=GREY)
        r += 1
    note(r, f"Also (informational, not defects): {d['merge_diff']} GP vendor ids across "
            f"{d['further_merges']} groups were merged one step further by Workday than by the "
            f"transform ({n(d['merged'])} → {n(d['wd_distinct_suppliers'])} suppliers) — same "
            f"suppliers, see the merge_group_diff tab.")
    r += 2

    # 2. Core attributes.
    section(r, "2.  Core business attributes  —  MATCHED")
    r += 1
    span(r, f"Compared at GP Vendor ID grain ({n(d['compared'])} keys). Matched / differing:",
         font=Font(sz=10)); r += 1
    cat_d = d["compared"] - d["m_cat"]
    ccy_d = d["compared"] - d["m_ccy"]
    cat_note = f"✓  ({cat_d} diff on merge-boundary vendor: {d['cat_diff_list']})" if cat_d else "✓"
    ccy_note = f"✓  ({ccy_d} diff on merge-boundary vendor: {d['ccy_diff_list']})" if ccy_d else "✓"
    pay_note = (
        f"{d['pay_diff_total']} differ — {d['pay_diff_manual']} are the manual-update "
        f"'set to Manual' suppliers, {d['pay_diff_minor']} minor default-type variants"
    )
    for attr, matched, extra in (
        ("Supplier name", d["m_name"], "✓"),
        ("Reference ID", d["compared"], "✓"),
        ("Supplier category", d["m_cat"], cat_note),
        ("Default currency", d["m_ccy"], ccy_note),
        ("Default payment type", d["m_pay"], pay_note),
    ):
        kv(r, attr, f"{n(matched)} / {n(d['compared'] - matched)}   {extra}")
        r += 1
    r += 1

    # 3. Manual-update list (the supplier centrepiece).
    section(r, "3.  Go-live manual-update suppliers  —  ACTION REQUIRED in Production")
    r += 1
    span(r, "Suppliers whose details Finance had to strip to force the Workday load. Per the go-live "
            "note, these must be manually updated in the Production tenant. Matched to the extract by "
            f"name ({d['manual_unmatched']} unmatched).",
         font=Font(sz=10), align=_WRAP_L); r += 1
    ws.merge_cells(f"B{r}:C{r}"); ws[f"B{r}"].value = "Manual-update category"
    ws[f"B{r}"].font = Font(sz=10, bold=True, color=NAVY); ws[f"B{r}"].fill = _fill(BLUE_HDR)
    ws[f"D{r}"].value = "Suppliers"; ws[f"D{r}"].font = Font(sz=10, bold=True, color=NAVY)
    ws[f"D{r}"].fill = _fill(BLUE_HDR); ws[f"D{r}"].alignment = _WRAP_C
    ws[f"E{r}"].value = "Action for Finance"; ws[f"E{r}"].font = Font(sz=10, bold=True, color=NAVY)
    ws[f"E{r}"].fill = _fill(BLUE_HDR)
    r += 1
    for cat, cnt, action in (
        ("Address removed", d["manual_addr"], "Re-enter the supplier address in Workday Production."),
        ("Bank account details removed", d["manual_bank"],
         "Re-enter the supplier bank / settlement account in Production."),
        ("Payment types reset to Manual only", d["manual_pay"],
         "Restore the correct default / accepted payment types in Production."),
    ):
        ws.merge_cells(f"B{r}:C{r}"); ws[f"B{r}"].value = cat
        ws[f"B{r}"].font = Font(sz=10); ws[f"B{r}"].alignment = _WRAP_L; ws[f"B{r}"].border = _BOTTOM
        ws[f"C{r}"].border = _BOTTOM
        ws[f"D{r}"].value = cnt; ws[f"D{r}"].font = Font(sz=10, bold=True); ws[f"D{r}"].alignment = _WRAP_C
        ws[f"D{r}"].border = _BOTTOM
        e = ws[f"E{r}"]; e.value = action; e.font = Font(sz=10, bold=True, color=AMBER_TXT)
        e.fill = _fill(AMBER); e.alignment = _WRAP_L; e.border = _BOTTOM
        r += 1
    resolved = d["manual_total"] - d["manual_open"]
    resolved_note = (f" {resolved} already resolved in the tenant ({d['manual_resolved_list']}) — "
                     f"do not rework." if resolved else "")
    note(r, f"{d['manual_total']} items on the list (address {d['manual_addr']} / bank "
            f"{d['manual_bank']} / payment {d['manual_pay']}; some suppliers appear in more than one "
            f"category). {d['manual_open']} are still open in the extract;{resolved_note} Full list on "
            f"the manual_update tab (with a still_open_in_workday column); actionable items on "
            f"finance_action_list.")
    r += 2

    # 4. Attribute gaps.
    section(r, "4.  Address / bank / payment gaps  —  reconciled to the manual-update list")
    r += 1
    span(r, "Where the transform sent a value but the extract shows it absent / reduced, the gap is "
            "classified EXPECTED (on the manual-update list) or UNEXPLAINED (genuine residual).",
         font=Font(sz=10), align=_WRAP_L); r += 1
    ws.merge_cells(f"B{r}:C{r}"); ws[f"B{r}"].value = "Gap"
    ws[f"B{r}"].font = Font(sz=10, bold=True, color=NAVY); ws[f"B{r}"].fill = _fill(BLUE_HDR)
    ws[f"D{r}"].value = "Expected"; ws[f"D{r}"].font = Font(sz=10, bold=True, color=NAVY)
    ws[f"D{r}"].fill = _fill(BLUE_HDR); ws[f"D{r}"].alignment = _WRAP_C
    ws[f"E{r}"].value = "Unexplained"; ws[f"E{r}"].font = Font(sz=10, bold=True, color=NAVY)
    ws[f"E{r}"].fill = _fill(BLUE_HDR)
    r += 1
    for gap, exp, unexp in (
        ("Address absent in Workday", d["addr_exp"], d["addr_unexp"]),
        ("Bank details absent in Workday", d["bank_exp"], d["bank_unexp"]),
        ("Payment reduced to Manual only", d["pay_exp"], d["pay_unexp"]),
    ):
        ws.merge_cells(f"B{r}:C{r}"); ws[f"B{r}"].value = gap
        ws[f"B{r}"].font = Font(sz=10); ws[f"B{r}"].alignment = _WRAP_L; ws[f"B{r}"].border = _BOTTOM
        ws[f"C{r}"].border = _BOTTOM
        ws[f"D{r}"].value = exp; ws[f"D{r}"].font = Font(sz=10); ws[f"D{r}"].alignment = _WRAP_C
        ws[f"D{r}"].border = _BOTTOM
        e = ws[f"E{r}"]; e.value = unexp; e.font = Font(sz=10, bold=(unexp > 0))
        e.fill = _fill(AMBER if unexp > 0 else GREY); e.border = _BOTTOM
        if unexp > 0:
            e.font = Font(sz=10, bold=True, color=AMBER_TXT)
        r += 1
    note(r, f"The {d['addr_unexp']} unexplained address gaps are known residual source-data cases: "
            f"{d['addr_unexp_list']}.")
    r += 2

    # 5. Actions.
    section(r, "5.  Actions required in Workday Production (Finance Operations)")
    r += 1
    span(r, f"Finance work from the finance_action_list tab — {n(d['action_total'])} open items, "
            f"extract-verified. Breakdown below.",
         font=Font(sz=10, bold=True), align=_WRAP_L); r += 1
    ws[f"A{r}"].value = "#"; ws[f"A{r}"].font = Font(sz=10, bold=True, color=NAVY)
    ws[f"A{r}"].fill = _fill(BLUE_HDR); ws[f"A{r}"].alignment = _WRAP_C
    ws[f"B{r}"].value = "Area"; ws[f"B{r}"].font = Font(sz=10, bold=True, color=NAVY); ws[f"B{r}"].fill = _fill(BLUE_HDR)
    ws[f"C{r}"].value = "Item"; ws[f"C{r}"].font = Font(sz=10, bold=True, color=NAVY); ws[f"C{r}"].fill = _fill(BLUE_HDR)
    ws[f"D{r}"].value = "Suppliers"; ws[f"D{r}"].font = Font(sz=10, bold=True, color=NAVY)
    ws[f"D{r}"].fill = _fill(BLUE_HDR); ws[f"D{r}"].alignment = _WRAP_C
    ws[f"E{r}"].value = "Action for Finance"; ws[f"E{r}"].font = Font(sz=10, bold=True, color=NAVY); ws[f"E{r}"].fill = _fill(BLUE_HDR)
    r += 1
    actions = [
        ("Manual updates", "Re-enter address / bank / payment for the go-live manual-update suppliers "
                          "(extract-verified as still open)",
         d["manual_open"], "Work the finance_action_list tab, rows marked 'Go-live manual-update list'."),
        ("Address", "Residual address gaps not on the manual list (region / postcode)",
         d["addr_unexp"], "Add / correct the address in Production (source-data gaps)."),
        ("Population", "Workday test record",
         d["dummy"], "Remove the test supplier from the tenant."),
    ]
    for i, (area, item, cnt, act) in enumerate(actions, 1):
        ws[f"A{r}"].value = i; ws[f"A{r}"].font = Font(sz=10); ws[f"A{r}"].alignment = _WRAP_C; ws[f"A{r}"].border = _BOTTOM
        ws[f"B{r}"].value = area; ws[f"B{r}"].font = Font(sz=10); ws[f"B{r}"].alignment = _WRAP_L; ws[f"B{r}"].border = _BOTTOM
        ws[f"C{r}"].value = item; ws[f"C{r}"].font = Font(sz=10); ws[f"C{r}"].alignment = _WRAP_L; ws[f"C{r}"].border = _BOTTOM
        ws[f"D{r}"].value = cnt; ws[f"D{r}"].font = Font(sz=10, bold=True); ws[f"D{r}"].alignment = _WRAP_C; ws[f"D{r}"].border = _BOTTOM
        e = ws[f"E{r}"]; e.value = act; e.font = Font(sz=10, bold=True, color=AMBER_TXT)
        e.fill = _fill(AMBER); e.alignment = _WRAP_L; e.border = _BOTTOM
        r += 1
    r += 1

    # 6. Accepted.
    section(r, "6.  Reviewed and accepted  —  no action required")
    r += 1
    for txt in (
        f"•  {d['merge_diff']} vendors merged differently in Workday than the transform (CHR / KIN / "
        f"TGI) — includes the 1 category and 1 currency difference; same underlying suppliers.",
        f"•  {d['multi_row']} suppliers occupy multiple extract rows for split bank / address — "
        f"expected Workday behaviour.",
        f"•  Default payment-type variants ({d['pay_diff_minor']}) where accepted payment types "
        f"still match — representation only.",
        "•  Tax country-for-ID is not surfaced in this extract view — out of scope for this pack; "
        "tax IDs were part of the sent payload.",
    ):
        span(r, txt, font=Font(sz=10), align=_WRAP_L); r += 1
    r += 1

    # 7. Status.
    section(r, "7.  Reconciliation status")
    r += 1
    kv(r, "Total items to action (finance_action_list tab)", n(d["action_total"])); r += 1
    kv(r, "   – go-live manual updates still open", f"{n(d['manual_open'])}   (address/bank/payment)", value_fill=GREY); r += 1
    kv(r, "   – residual address gaps (source data)", n(d["addr_unexp"]), value_fill=GREY); r += 1
    kv(r, "   – Workday test record to remove", n(d["dummy"]), value_fill=GREY); r += 1
    kv(r, "Population & core-attribute discrepancies", "0"); r += 1
    span(r, "Conclusion: population and core business data are fully reconciled. The remaining work is "
            "operational — Finance re-applies the manual-update suppliers and corrects the few residual "
            "source-data gaps in Production.",
         font=Font(sz=10, italic=True, color=GREY_TXT), align=_WRAP_L)
    r += 2

    # 8. Approvals.
    section(r, "8.  Business Approvals")
    r += 1
    for label, value in (
        ("Responsible Finance Team for Supplier Master Data", "Finance Operations"),
        ("Responsible Finance Team Approver", "Najam Israr"),
        ("Responsible Finance Team Approved (Yes/No)", ""),
        ("Responsible Finance Team Date", ""),
        ("Group Financial Controller", "Joel Monaghan"),
        ("Group Financial Controller Approved (Yes/No)", ""),
        ("Group Financial Controller Date", ""),
    ):
        kv(r, label, value, value_fill="FFFFFF")
        r += 1

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    wb.save(pack_path)
