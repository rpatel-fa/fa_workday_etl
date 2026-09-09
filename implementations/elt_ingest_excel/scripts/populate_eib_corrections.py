"""Populate the yellow CORRECTION column (column C) on the Workday EIB error tab.

Given a Workday "Submit Supplier" EIB workbook whose second sheet ("Sheet1") is the
EIB load *Error Summary*, this fills column C (the yellow CORRECTION column) for every
error line with the value the finance team should apply.

Two decoupled sources, joined on the stable business key only:
* The EIB's own bank block maps each error line to a specific supplier/account and is used
  to disambiguate which value an error refers to (its numeric cells are lossy - leading
  zeros and >15-digit accounts are already corrupted - so it is NOT used for the value).
* The authoritative correction VALUE is read from the transform's DuckDB output keyed on
  the stable reference_id (GP Vendor ID), never the run-scoped supplier_id, so it is safe
  across pipeline reruns and preserves leading zeros and the transform's field precedence.
  Pass --duckdb '' to fall back to EIB-only derivation (lossy).

Design notes
------------
* The workbook is NOT re-saved through openpyxl: EIB templates carry named ranges and
  data validations that a full openpyxl round-trip can silently drop. Instead the clone
  is a byte-for-byte copy of the original zip with ONLY the error-tab sheet XML rewritten,
  so the Submit Supplier sheet and every EIB feature are preserved exactly.
* Corrections are read-only derivations from the source cells; the original file and the
  Submit Supplier data are never mutated.

Correction logic (per error message)
-------------------------------------
* "Financial Account Number is required"  -> account recovered from the transaction's IBAN
  cell: GB IBAN -> chars 15-22; a value that is not a structural ISO-13616 IBAN (US/SG/...)
  -> used as the raw local account number. Also recovers country-specific account slices for
  CR / GE / UA / KZ IBANs where Workday still requires a separate domestic account field.
* "Routing Transit or Institution Number is required" -> GB sort code from IBAN chars 9-14;
  otherwise flagged MANUAL (no routing/institution in source).
* "... must be 9 numeric digits" / "check sum ... Routing" -> US routing zero-padded to 9.
* "Region Name must be valid" / "NOT ON NAMED RANGE TAB" -> correct country-scoped
  Country_Region_ID looked up from ref_workday_country_state_region, else MANUAL.
* BIC format errors -> recover from a valid sibling BIC on the same key when present, else
  suggest a unique published-output match for the same bank/country when available.
* City Subdivision / scrambled source fields -> MANUAL with a specific reason.

Usage
-----
    uv run python scripts/populate_eib_corrections.py "<input.xlsx>" [--out "<output.xlsx>"]

    # Strict business-safe mode: write only confirmed fixes to error-tab column C and leave
    # all non-confirmed rows blank. Do not modify Submit Supplier.
    uv run python scripts/populate_eib_corrections.py "<input.xlsx>" --strict-column-c
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import shutil
import sys
import zipfile
from collections import defaultdict
from xml.etree import ElementTree as ET

import openpyxl

# --- "Submit Supplier - v46.2" column map (1-based) -------------------------------------
KEY, REF, NAME = 2, 9, 11
ADDR_COUNTRY, MUNICIPALITY, REGION, REGION_DESC, SUBREGION = 115, 121, 128, 129, 133
SUBMUNI, SUBMUNI_TYPE = 127, 126  # City Subdivision 1 value (col127) + its Type tag (col126)
CITY_SUBDIVISION_TYPE = "CITY_SUBDIVISION_1"  # parallels ADDRESS_LINE_1 / REGION_SUBDIVISION_1
BANK_COUNTRY, ROUTING, ACCT, IBAN, BIC = 224, 229, 232, 236, 237

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = f"{{{MAIN_NS}}}"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
X14AC_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
XR_NS = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
XR2_NS = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
XR3_NS = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"

DEFAULT_INPUT = os.path.expanduser(
    "~/Downloads/Master Suppliers V1 24.07.2026 with errors.xlsx"
)
REF_REGION_CSV = os.path.join(
    os.path.dirname(__file__),
    "..",
    "config",
    "data",
    "ref_workday_country_state_region.csv",
)
REF_COUNTRY_ISO_JSON = os.path.join(
    os.path.dirname(__file__),
    "..",
    "config",
    "data",
    "ref_country_iso3_mapping.json",
)


# --- derivations ------------------------------------------------------------------------
FLOAT_EXACT_MAX = 2 ** 53  # doubles hold integers exactly only up to here (~15-16 digits)


def exact_text(v):
    """Exact digit/text of a cell value, plus a 'lossy' flag.

    Account/IBAN cells are stored as numbers in this workbook; a value beyond float
    precision (>15-16 digits) has ALREADY lost its low-order digits in the source file,
    so it must never be emitted as a correction. Returns (text, lossy).
    """
    if v in (None, ""):
        return "", False
    if isinstance(v, bool):
        return str(v), False
    if isinstance(v, int):
        return str(v), False
    if isinstance(v, float):
        if v.is_integer():
            return format(int(v), "d"), abs(v) >= FLOAT_EXACT_MAX
        return repr(v), True
    return str(v).strip(), False


def clean_iban(text: str) -> str:
    return re.sub(r"[\s-]", "", (text or "").upper())


def is_gb_iban(c: str) -> bool:
    return bool(re.match(r"^GB[0-9]{2}[A-Z]{4}[0-9]{14}$", c))


def is_valid_bic(text) -> bool:
    return bool(re.fullmatch(r"[A-Z0-9]{8}([A-Z0-9]{3})?", clean_iban(exact_text(text)[0])))


def canon_key(text):
    return re.sub(r"[^A-Z0-9]+", "", str(text or "").upper())


def recover_account(iban_raw):
    """Return (account, lossy). GB IBAN -> chars 15-22; a non-ISO-13616 value -> raw local
    account. lossy=True means the source numeric cell lost precision and the exact account
    cannot be recovered."""
    text, lossy = exact_text(iban_raw)
    c = clean_iban(text)
    if is_gb_iban(c):
        return c[14:22], False  # GB IBAN is text, never lossy
    # Country-specific domestic account slices for real IBAN countries still needing a
    # separate account number in Workday.
    iban_account_slices = (
        (re.compile(r"^CR[0-9]{2}[0-9]{4}[A-Z0-9]{14}$"), 8),   # CR + check + 4-digit bank
        (re.compile(r"^GE[0-9]{2}[A-Z]{2}[A-Z0-9]{16}$"), 6),   # GE + check + 2-char bank
        (re.compile(r"^UA[0-9]{2}[0-9]{6}[A-Z0-9]{19}$"), 10),  # UA + check + 6-digit bank
        (re.compile(r"^KZ[0-9]{2}[0-9]{3}[A-Z0-9]{13}$"), 7),   # KZ + check + 3-digit bank
    )
    for patt, start in iban_account_slices:
        if patt.match(c):
            return c[start:], False
    if c and not re.match(r"^[A-Z]{2}[0-9]{2}", c):
        return (None, True) if lossy else (text, False)
    return None, False


def recover_sort(iban_raw):
    text, _ = exact_text(iban_raw)
    c = clean_iban(text)
    return c[8:14] if is_gb_iban(c) else None  # chars 9-14


def pad_us_routing(routing):
    text, _ = exact_text(routing)
    digits = re.sub(r"\D", "", text)
    return "0" + digits if len(digits) == 8 else None


def is_structural_iban(iban_raw):
    """True if the value is a real ISO-13616 IBAN (country + check digits), vs a raw local
    account number misfiled in the IBAN column."""
    c = clean_iban(exact_text(iban_raw)[0])
    return bool(re.match(r"^[A-Z]{2}[0-9]{2}[A-Z0-9]+$", c))


def norm_text(text):
    return str(text or "").strip().casefold()


# --- workbook extraction ----------------------------------------------------------------
def cell(row, idx):
    v = row[idx - 1] if idx - 1 < len(row) else None
    if v in (None, ""):
        return None
    return str(v).strip() if isinstance(v, str) else v


def extract(input_path):
    """Return (errors, key_bank, key_addr).

    errors   : ordered list of (sheet_row, key, message) from the error tab
    key_bank : key -> list of dicts {acct, iban, routing, country, bic} (one per bank block)
    key_addr : key -> dict {ref, name, country, municipality, region, region_desc, subregion}
    """
    wb = openpyxl.load_workbook(input_path, read_only=True, data_only=True)

    errs = []
    ws_err = wb["Sheet1"]
    for i, row in enumerate(ws_err.iter_rows(values_only=True), start=1):
        if isinstance(row[0], int) and row[1]:
            errs.append((i, row[0], str(row[1])))

    key_bank = defaultdict(list)
    key_addr = {}
    submuni_rows = []  # (sheet_row, key, value) - City Subdivision 1 value in col127
    ws = wb["Submit Supplier"]
    cur = None
    rownum = 5
    for row in ws.iter_rows(min_row=6, values_only=True):
        rownum += 1
        k = cell(row, KEY)
        if isinstance(k, int):
            cur = k
        sub = cell(row, SUBMUNI)
        if sub is not None:
            submuni_rows.append((rownum, cur, str(sub)))
        if cur is None:
            continue
        # bank block on this sub-row?
        acct, iban = cell(row, ACCT), cell(row, IBAN)
        routing, bic, bctry = cell(row, ROUTING), cell(row, BIC), cell(row, BANK_COUNTRY)
        if any(v is not None for v in (acct, iban, routing, bic)):
            key_bank[cur].append(
                {"acct": acct, "iban": iban, "routing": routing, "country": bctry, "bic": bic}
            )
        # address block: capture first populated address for the key
        actry = cell(row, ADDR_COUNTRY)
        if actry is not None and cur not in key_addr:
            key_addr[cur] = {
                "ref": cell(row, REF),
                "name": cell(row, NAME),
                "country": actry,
                "municipality": cell(row, MUNICIPALITY),
                "region": cell(row, REGION),
                "region_desc": cell(row, REGION_DESC),
                "subregion": cell(row, SUBREGION),
            }
        if cur not in key_addr:
            key_addr[cur] = {"ref": cell(row, REF), "name": cell(row, NAME)}
    wb.close()
    return errs, key_bank, key_addr, submuni_rows


def load_duckdb_bank(db_path):
    """Reference-keyed authoritative bank outputs plus reusable valid-BIC mappings."""
    import duckdb

    con = duckdb.connect(os.path.expanduser(db_path), read_only=True)
    sid_acct = defaultdict(list)
    sid_sort = defaultdict(list)
    sid_bic = defaultdict(list)
    sid_bank_meta = defaultdict(list)
    bic_by_bank = defaultdict(set)
    for sid, acct, sort, bic, country, bank_name in con.execute(
        "SELECT supplier_id, bank_account_number, bank_code_routing_number, "
        "bank_swift_code, bank_country, bank_name "
        "FROM workday_supplier_settlement_account"
    ).fetchall():
        if acct not in (None, ""):
            sid_acct[sid].append(str(acct).strip())
        if sort not in (None, ""):
            sid_sort[sid].append(str(sort).strip())
        if country or bank_name:
            sid_bank_meta[sid].append((str(country or "").strip(), str(bank_name or "").strip()))
        if is_valid_bic(bic):
            bic = clean_iban(exact_text(bic)[0])
            sid_bic[sid].append(bic)
            bic_by_bank[(canon_key(country), canon_key(bank_name))].add(bic)
    ref_acct = defaultdict(list)
    ref_sort = defaultdict(list)
    ref_bic = defaultdict(list)
    ref_bank_meta = defaultdict(list)
    for sid, ref in con.execute(
        "SELECT supplier_id, reference_id FROM workday_supplier_name WHERE reference_id IS NOT NULL"
    ).fetchall():
        for tok in str(ref).split("|"):
            tok = tok.strip()
            ref_acct[tok].extend(sid_acct.get(sid, []))
            ref_sort[tok].extend(sid_sort.get(sid, []))
            ref_bic[tok].extend(sid_bic.get(sid, []))
            ref_bank_meta[tok].extend(sid_bank_meta.get(sid, []))
    con.close()
    return {
        "acct": ref_acct,
        "sort": ref_sort,
        "bic": ref_bic,
        "meta": ref_bank_meta,
        "bic_by_bank": bic_by_bank,
    }


def load_region_ref():
    """(country_lower, locality_lower) -> set(Country_Region_ID). Also ISO2 -> country name."""
    by_locality = defaultdict(set)
    iso_to_country = {}
    iso_path = os.path.normpath(REF_COUNTRY_ISO_JSON)
    if os.path.exists(iso_path):
        with open(iso_path, encoding="utf-8") as fh:
            for row in json.load(fh):
                iso2 = (row.get("country_code_iso2") or "").strip().upper()
                country = (row.get("country_name") or "").strip()
                if iso2 and country:
                    iso_to_country[iso2] = country
    path = os.path.normpath(REF_REGION_CSV)
    if not os.path.exists(path):
        return by_locality, iso_to_country
    with open(path, newline="", encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            country = (r.get("country") or "").strip()
            rid = (r.get("reference_ids") or r.get("reference_id") or "").strip()
            inst = (r.get("instance") or "").strip()
            desc = (r.get("workbook_value_descriptor") or "").strip()
            if not (country and rid):
                continue
            for loc in (inst, desc):
                if loc:
                    by_locality[(norm_text(country), norm_text(loc))].add(rid)
    return by_locality, iso_to_country


# --- correction resolution --------------------------------------------------------------
def _nz(s):
    """Strip leading zeros for leading-zero-insensitive matching."""
    return (str(s).lstrip("0") or "0")


def resolve_corrections(
    errs,
    key_bank,
    key_addr,
    region_ref,
    iso_to_country,
    db_bank=None,
    key_submuni=None,
):
    """Return list of (sheet_row, key, ref, message, correction, status).

    status in {AUTO, SUGGEST, MANUAL}. When db_bank is provided, correction VALUES are taken
    from the authoritative transform output (DuckDB) keyed on the stable reference_id; the
    EIB's own IBAN derivation is used only to map which bank block an error line refers to.
    This preserves leading zeros and the transform's field-precedence (e.g. GP sort code over
    the IBAN-embedded one), which the EIB's lossy numeric cells cannot.
    """
    # pre-compute ordered derivations per key (only for blocks missing the local field).
    # acct entries are (value, lossy): value None + lossy True => source precision lost.
    acct_by_key = {}
    sort_by_key = {}
    for k, blocks in key_bank.items():
        accts = []
        for b in blocks:
            if b["acct"]:
                continue
            val, lossy = recover_account(b["iban"])
            if val is not None or lossy:
                accts.append((val, lossy))
        acct_by_key[k] = accts
        sort_by_key[k] = [s for b in blocks if not b["routing"]
                          and (s := recover_sort(b["iban"])) is not None]

    def uniq(values):
        seen = set()
        out = []
        for v in values:
            if v in seen or v in (None, ""):
                continue
            seen.add(v)
            out.append(v)
        return out

    # per-key authoritative pools from DuckDB (mutable - consumed as error lines resolve)
    db_acct_pool = {}
    db_sort_pool = {}
    db_bic_pool = {}
    db_bank_meta = {}
    if db_bank is not None:
        for k in key_bank:
            toks = [t.strip() for t in str(key_addr.get(k, {}).get("ref") or "").split("|")]
            db_acct_pool[k] = [a for t in toks for a in db_bank["acct"].get(t, [])]
            db_sort_pool[k] = [s for t in toks for s in db_bank["sort"].get(t, [])]
            db_bic_pool[k] = uniq([b for t in toks for b in db_bank["bic"].get(t, [])])
            db_bank_meta[k] = uniq([m for t in toks for m in db_bank["meta"].get(t, [])])

    def take_authoritative(pool, hint):
        """Pop the authoritative value from pool matching hint (leading-zero-insensitive);
        else the first remaining (best-effort for lossy/None hints). Returns value or None."""
        if not pool:
            return None
        if hint:
            for i, v in enumerate(pool):
                if _nz(v) == _nz(hint):
                    return pool.pop(i)
        return pool.pop(0)

    counters = defaultdict(int)

    def next_val(k, kind, values):
        i = counters[(k, kind)]
        counters[(k, kind)] += 1
        return values[i] if i < len(values) else None

    def region_lookup(k):
        a = key_addr.get(k, {})
        country_name = iso_to_country.get(str(a.get("country") or "").strip().upper())
        # Prefer the workbook's human region descriptor first: it is the exact value the EIB
        # resolver tried to map. Municipality/subregion remain useful fallbacks for places where
        # the workbook carries locality text instead of the higher-level district/province.
        localities = [a.get("region_desc"), a.get("municipality"), a.get("subregion")]
        for loc in localities:
            if not loc:
                continue
            if country_name:
                hits = region_ref.get((norm_text(country_name), norm_text(loc)), set())
                if hits:
                    return hits, True
        hits = set()
        for loc in localities:
            if not loc:
                continue
            for (_c_lower, l_lower), ids in region_ref.items():
                if l_lower == norm_text(loc):
                    hits |= ids
        return hits, False

    out = []
    for sheet_row, k, msg in errs:
        ref = key_addr.get(k, {}).get("ref")
        m = msg.replace("Validation error occurred. ", "")
        correction, status = None, "MANUAL"

        if m.startswith("Financial Account Number is required"):
            entry = next_val(k, "acct", acct_by_key.get(k, []))
            eib_hint = entry[0] if entry else None
            auth = take_authoritative(db_acct_pool.get(k, []), eib_hint) if db_bank else None
            if auth:
                correction, status = auth, "AUTO"
            elif eib_hint:
                correction, status = eib_hint, "AUTO"  # no DuckDB match; EIB-derived
            elif entry and entry[1]:
                correction = ("MANUAL: account stored as a number >15 digits - precision "
                              "lost in source; re-key exact account")
            else:
                correction = "MANUAL: no account in IBAN column - source value required"

        elif m.startswith("Routing Transit or Institution Number is required"):
            eib_hint = next_val(k, "sort", sort_by_key.get(k, []))
            auth = take_authoritative(db_sort_pool.get(k, []), eib_hint) if db_bank else None
            if auth:
                correction, status = auth, "AUTO"
            elif eib_hint:
                correction, status = eib_hint, "AUTO"
            else:
                correction = "MANUAL: no routing/institution in source for this country"

        elif "Routing Transit Number must be 9 numeric digits" in m or "check sum validation failed" in m:
            blocks = key_bank.get(k, [])
            padded = next((p for b in blocks if (p := pad_us_routing(b["routing"]))), None)
            if padded:
                correction, status = padded, "AUTO"
            else:
                correction = "MANUAL: routing number invalid - confirm with bank"

        elif m.startswith("Region Name must be valid") or "NOT ON NAMED RANGE TAB" in m:
            hits, scoped = region_lookup(k)
            if len(hits) == 1:
                correction = next(iter(hits))
                status = "AUTO" if scoped else "SUGGEST"
            else:
                a = key_addr.get(k, {})
                correction = (
                    f"MANUAL: set correct Country_Region_ID for {a.get('country')}"
                    f" / {a.get('municipality') or a.get('region_desc')}"
                )

        elif "City Subdivision 1" in m:
            val = (key_submuni or {}).get(k)
            if val:
                # Keep the correction cell business-friendly: just the value to apply.
                correction, status = val, "AUTO"
            else:
                country = re.search(r"components for ([^:]+):", m)
                correction = (
                    f"MANUAL: City Subdivision 1 value missing for "
                    f"{country.group(1).strip() if country else '?'}"
                )

        elif "missing address components" in m:
            comp = re.search(r"components for ([^:]+):(.*)", m)
            correction = (
                f"MANUAL: address component(s) missing"
                + (f" ({comp.group(2).strip()})" if comp else "")
            )

        elif "Bank Identification Code" in m or "BIC" in m:
            sibling_bics = uniq(
                [
                    clean_iban(str((b.get("bic") or "")).strip())
                    for b in key_bank.get(k, [])
                    if is_valid_bic(b.get("bic"))
                ]
            )
            if len(sibling_bics) == 1:
                correction, status = sibling_bics[0], "AUTO"
            elif len(db_bic_pool.get(k, [])) == 1:
                correction, status = db_bic_pool[k][0], "AUTO"
            else:
                bank_bics = []
                for country, bank_name in db_bank_meta.get(k, []):
                    hits = db_bank["bic_by_bank"].get((canon_key(country), canon_key(bank_name)), set())
                    if len(hits) == 1:
                        bank_bics.extend(sorted(hits))
                bank_bics = uniq(bank_bics)
                if len(bank_bics) == 1:
                    correction, status = bank_bics[0], "SUGGEST"
                else:
                    correction = "MANUAL: BIC/SWIFT invalid in source - confirm valid 8/11-char BIC"

        elif "Bank Code must be 3 numeric digits" in m:
            correction = "MANUAL: bank code invalid in source (junk value)"

        elif "Bank Sort Code must be 6 numeric digits" in m:
            correction = "MANUAL: sort-code field holds a non-sort-code value in source"

        elif m.startswith("Branch ID is required"):
            correction = "MANUAL: branch/transit ID required for this country - not in source"

        else:
            correction = "MANUAL: review - unclassified error"

        out.append((sheet_row, k, ref, msg, correction, status))
    return out


# --- surgical zip / sheet-XML rewrite ---------------------------------------------------
def error_tab_xml_name(zf):
    """Return the archive path of the 'Sheet1' worksheet XML."""
    wbxml = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_by_name = {}
    for s in wbxml.iter(f"{NS}sheet"):
        rid = s.get(f"{{http://schemas.openxmlformats.org/officeDocument/2006/relationships}}id")
        rid_by_name[s.get("name")] = rid
    rid = rid_by_name.get("Sheet1")
    target = None
    for rel in rels:
        if rel.get("Id") == rid:
            target = rel.get("Target")
    if target is None:
        raise SystemExit("Could not locate 'Sheet1' worksheet in workbook rels")
    return "xl/" + target.lstrip("/")


def rewrite_error_tab(xml_bytes, corrections_by_row):
    ET.register_namespace("", MAIN_NS)
    ET.register_namespace("r", REL_NS)
    ET.register_namespace("mc", MC_NS)
    ET.register_namespace("x14ac", X14AC_NS)
    ET.register_namespace("xr", XR_NS)
    ET.register_namespace("xr2", XR2_NS)
    ET.register_namespace("xr3", XR3_NS)
    root = ET.fromstring(xml_bytes)
    sheet_data = root.find(f"{NS}sheetData")
    rows = {int(r.get("r")): r for r in sheet_data.findall(f"{NS}row")}
    for sheet_row, value in corrections_by_row.items():
        row_el = rows.get(sheet_row)
        if row_el is None:
            continue
        ref = f"C{sheet_row}"
        c = next((x for x in row_el.findall(f"{NS}c") if x.get("r") == ref), None)
        if c is None:
            c = ET.SubElement(row_el, f"{NS}c")
            c.set("r", ref)
        c.set("t", "inlineStr")
        for child in list(c):
            c.remove(child)
        is_el = ET.SubElement(c, f"{NS}is")
        t_el = ET.SubElement(is_el, f"{NS}t")
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t_el.text = value
    # ElementTree drops unused namespace declarations, but Excel relies on the worksheet root
    # retaining the original mc/xr/x14ac prefix declarations referenced by mc:Ignorable.
    root_open = (
        f'<worksheet xmlns="{MAIN_NS}"'
        f' xmlns:r="{REL_NS}"'
        f' xmlns:mc="{MC_NS}"'
        f' mc:Ignorable="{root.get(f"{{{MC_NS}}}Ignorable", "")}"'
        f' xmlns:x14ac="{X14AC_NS}"'
        f' xmlns:xr="{XR_NS}"'
        f' xmlns:xr2="{XR2_NS}"'
        f' xmlns:xr3="{XR3_NS}"'
    )
    xr_uid = root.get(f"{{{XR_NS}}}uid")
    if xr_uid:
        root_open += f' xr:uid="{xr_uid}"'
    root_open += ">"
    decl = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    body = ET.tostring(root, encoding="unicode")
    body = body.replace(body.split(">", 1)[0] + ">", root_open, 1)
    return decl + body.encode("utf-8")


def submit_supplier_xml_name(zf):
    """Archive path of the 'Submit Supplier' worksheet XML."""
    wbxml = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid = {s.get("name"): s.get(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        for s in wbxml.iter(f"{NS}sheet")}.get("Submit Supplier")
    tgt = next((r.get("Target") for r in rels if r.get("Id") == rid), None)
    return "xl/" + tgt.lstrip("/") if tgt else None


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def patch_city_subdivision(xml_bytes, submuni_rows):
    """Byte-level insert of the City Subdivision Type cell (col126) before the existing
    Submunicipality value cell (col127) for each given row. The 132MB sheet is edited by
    targeted insertion (no full XML parse) to keep every other cell byte-identical.
    Returns (patched_bytes, count)."""
    dv, dw = col_letter(SUBMUNI_TYPE), col_letter(SUBMUNI)  # DV, DW
    count = 0
    for sheet_row, _key, _val in submuni_rows:
        anchor = f'<c r="{dw}{sheet_row}"'.encode()
        i = xml_bytes.find(anchor)
        if i == -1:
            continue
        if f'<c r="{dv}{sheet_row}"'.encode() in xml_bytes:  # already present - skip
            continue
        cell_xml = (f'<c r="{dv}{sheet_row}" t="inlineStr"><is><t>'
                    f'{CITY_SUBDIVISION_TYPE}</t></is></c>').encode()
        xml_bytes = xml_bytes[:i] + cell_xml + xml_bytes[i:]
        count += 1
    return xml_bytes, count


def clone_and_patch(input_path, output_path, corrections_by_row, submuni_rows=None):
    shutil.copyfile(input_path, output_path)
    with zipfile.ZipFile(input_path, "r") as zin:
        err_name = error_tab_xml_name(zin)
        ss_name = submit_supplier_xml_name(zin) if submuni_rows else None
        names = zin.namelist()
        patched_err = rewrite_error_tab(zin.read(err_name), corrections_by_row)
        infos = {i.filename: i for i in zin.infolist()}
        subdiv_count = 0
        patched_ss = None
        if ss_name and submuni_rows:
            patched_ss, subdiv_count = patch_city_subdivision(zin.read(ss_name), submuni_rows)
        with zipfile.ZipFile(input_path, "r") as zsrc, \
             zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                if name == err_name:
                    data = patched_err
                elif name == ss_name and patched_ss is not None:
                    data = patched_ss
                else:
                    data = zsrc.read(name)
                zout.writestr(infos[name], data)
    return subdiv_count


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("input", nargs="?", default=DEFAULT_INPUT, help="EIB workbook with error tab")
    ap.add_argument("--out", default=None, help="output path (default: '<input> CORRECTED.xlsx')")
    ap.add_argument("--audit", default=None, help="optional CSV audit of every correction")
    ap.add_argument(
        "--strict-column-c",
        action="store_true",
        help="Write only AUTO corrections to error-tab column C, leave SUGGEST/MANUAL rows blank, "
        "and do not modify Submit Supplier.",
    )
    ap.add_argument(
        "--duckdb",
        default="~/Documents/__data/duckdb/rpatel.duckdb",
        help="authoritative transform DB for correction values (reference_id keyed). "
        "Pass '' to derive from the EIB alone (lossy for numeric account cells).",
    )
    args = ap.parse_args(argv)

    input_path = os.path.expanduser(args.input)
    if not os.path.exists(input_path):
        raise SystemExit(f"Input not found: {input_path}")
    stem, ext = os.path.splitext(input_path)
    output_path = os.path.expanduser(args.out) if args.out else f"{stem} CORRECTED{ext}"

    errs, key_bank, key_addr, submuni_rows = extract(input_path)
    key_submuni = {k: v for _, k, v in submuni_rows if k is not None}
    region_ref, iso_to_country = load_region_ref()
    db_bank = None
    db_path = os.path.expanduser(args.duckdb) if args.duckdb else None
    if db_path and os.path.exists(db_path):
        db_bank = load_duckdb_bank(db_path)
        print(f"Authoritative values from: {db_path}")
    elif args.duckdb:
        print(f"WARNING: DuckDB not found ({db_path}); deriving from EIB alone (lossy).")
    resolved = resolve_corrections(errs, key_bank, key_addr, region_ref, iso_to_country,
                                   db_bank=db_bank, key_submuni=key_submuni)

    if args.strict_column_c:
        corrections_by_row = {row: corr for row, _, _, _, corr, status in resolved if status == "AUTO"}
        patch_rows = None
    else:
        corrections_by_row = {row: corr for row, _, _, _, corr, _ in resolved}
        patch_rows = submuni_rows
    subdiv_count = clone_and_patch(input_path, output_path, corrections_by_row, patch_rows)
    print(f"City Subdivision Type tags written to Submit Supplier: {subdiv_count}")

    counts = defaultdict(int)
    for *_, status in [(r[5],) for r in resolved]:
        counts[status] += 1
    print(f"Input : {input_path}")
    print(f"Output: {output_path}")
    print(f"Error lines: {len(resolved)}  |  "
          f"AUTO={counts['AUTO']}  SUGGEST={counts['SUGGEST']}  MANUAL={counts['MANUAL']}")
    print("-" * 100)
    for sheet_row, k, ref, msg, corr, status in resolved:
        tag = {"AUTO": "OK ", "SUGGEST": "SUG", "MANUAL": "MAN"}[status]
        print(f"[{tag}] key={k:<5} {str(ref):<10} C{sheet_row:<4} {msg[:52]:<52} -> {corr}")

    if args.audit:
        with open(os.path.expanduser(args.audit), "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["error_tab_row", "spreadsheet_key", "reference_id", "error", "correction", "status"])
            w.writerows(resolved)
        print(f"\nAudit CSV: {args.audit}")


if __name__ == "__main__":
    main()
