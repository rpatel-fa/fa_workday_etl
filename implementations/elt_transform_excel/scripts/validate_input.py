#!/usr/bin/env python3
"""Validate input workbook schemas for a registered elt_transform_excel transform.

Uses the transform's YAML profile (source_layout) as the single source of truth
for the required schema.  Produces a structured human-readable report of 12
checks across the source workbook, mapping workbook, supplier workbook, and
customer workbook, with EXIT=0 on all-passed or EXIT=1 on any failure.

Usage:
  uv run --with openpyxl --with pyyaml --no-project python \
    elt_transform_excel/scripts/validate_input.py --transform ap_amex_hotel

  # Override shell-defaults paths (useful for a new file before reconfiguring YAML):
  uv run --with openpyxl --with pyyaml --no-project python \
    elt_transform_excel/scripts/validate_input.py -t ap_amex_hotel \
    --source-data ~/Downloads/Hotels_July_26.xlsx

Checks performed (all are printed individually with symbols):
  FILE SOURCE DATA      exists on disk, opens as xlsx, not corrupt
  SHEET SOURCE LAYOUT   worksheet name exists (or first_sheet / literal_row layout)
  DIMENSIONS SOURCE     non-empty rows / cols reported by openpyxl
  HEADER ROW DETECT     required-headers all match on EXACT source_layout.header_row
                        (literal_row mode) or on first row with 100% match (scan mode)
  REQUIRED HEADERS      each required header listed with column position
                        ❌ = any missing header is enumerated with its expected name
  OPTIONAL HEADERS      ℹ️  each optional header is reported as present/absent
                        (absent is NOT a failure; these are gracefully skipped)
  REQUIRED COLS NON-ZERO sample of data rows: counts blanks in each required column
                        (rows that are entirely blank are filtered out as trailers)
  FILE TEMPLATE EIB     target EIB workbook exists / opens
  FILE MAPPING          mapping spreadsheet exists / opens + named sheet present
  SHEET MAPPING COLUMNS MAPPING sheet has the 5 canonical lookup headers
  FILE SUPPLIER         supplier extract exists / opens (AP-only)
  SHEET SUPPLIER COLS   supplier sheet has Reference ID + Supplier ID
  FILE CUSTOMER         customer extract exists / opens (AR-only)
  SHEET CUSTOMER COLS   customer sheet has Reference ID + Customer ID
"""
from __future__ import annotations

import argparse
import os
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

try:
    import openpyxl
    import yaml
except ImportError as e:  # pragma: no cover - surfaced clearly to CLI users
    sys.stderr.write(
        "ERROR: missing dependencies.  Re-run with:\n"
        "  uv run --with openpyxl --with pyyaml --no-project python "
        "elt_transform_excel/scripts/validate_input.py ...\n"
        f"  (missing: {e.name})\n"
    )
    sys.exit(2)


ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = ROOT / "config"
PROFILE_DIR = CONFIG_DIR / "transforms"
REF_PATHS_PATH = CONFIG_DIR / "reference_paths.yaml"
DEFAULTS_PATH = CONFIG_DIR / "eib_population_defaults.yaml"
MAX_HEADER_SCAN_ROWS = 40
MAX_DATA_SAMPLE_ROWS = 200
VALID_SPREADSHEET_EXTS = {".xlsx", ".xlsm", ".xls", ".csv"}


def _column_letter_to_index(letters: str) -> int:
    letters = (letters or "").strip().upper()
    result = 0
    for ch in letters:
        if "A" <= ch <= "Z":
            result = (result * 26) + (ord(ch) - ord("A") + 1)
    return result


@dataclass
class CheckResult:
    name: str
    ok: bool
    detail: str


def _load_yaml(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def _norm_token(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value != value:  # NaN guard
        return ""
    return str(value).strip().lower()


def _split_header_aliases(header_spec: object) -> list[str]:
    """Split a required-header entry into its alias list.

    Mirrors VBA SplitHeaderAliases:
      - Plain text "Description" -> one-element ["Description"]
      - Pipe-delimited "Decscription|Description" -> two elements, each stripped
      - Empty or whitespace-only alternatives are dropped
    """
    spec_str = str(header_spec or "")
    if "|" not in spec_str:
        stripped = spec_str.strip()
        return [stripped] if stripped else []
    parts: list[str] = []
    for alt in spec_str.split("|"):
        alt_clean = alt.strip()
        if alt_clean:
            parts.append(alt_clean)
    if not parts:
        return [spec_str.strip()]
    return parts


def _shell_defaults_variants(profile_data: dict) -> list[dict]:
    """Return shell_defaults with any per-variant override applied.

    If shell_defaults.variants is a non-empty list, return one dict per variant
    (base shell_defaults merged with the per-variant override map, plus an
    extra _binding_key / _binding_label for display). Otherwise return a
    single-element list containing the plain base shell_defaults (with a
    synthetic binding key sourced from transform.key / transform.label).
    """
    base = profile_data.get("shell_defaults") or {}
    base_clean = {k: v for k, v in base.items() if k != "variants"}
    variants_raw = base.get("variants")
    if not isinstance(variants_raw, list) or not variants_raw:
        transform = profile_data.get("transform") or {}
        variant = dict(base_clean)
        variant.setdefault(
            "_binding_key",
            str(transform.get("key") or ""),
        )
        variant.setdefault(
            "_binding_label",
            str(transform.get("label") or ""),
        )
        return [variant]
    variant_field_map = {
        "profile_key": "_binding_key",
        "profile_label": "_binding_label",
        "source_data": "source_data",
        "target_eib": "target_eib",
        "output_file_path": "output_file_path",
        "mapping_spreadsheet": "mapping_spreadsheet",
        "mapping_sheet": "mapping_sheet",
        "customer_spreadsheet": "customer_spreadsheet",
        "supplier_spreadsheet": "supplier_spreadsheet",
        "workday_reference_ids": "workday_reference_ids",
        "source_worksheet_name": "source_worksheet_name",
        "source_worksheet_resolution": "source_worksheet_resolution",
        "source_header_row": "source_header_row",
        "source_data_start_row": "source_data_start_row",
    }
    bindings: list[dict] = []
    for raw_variant in variants_raw:
        if not isinstance(raw_variant, dict):
            continue
        merged = dict(base_clean)
        for yaml_key, merged_key in variant_field_map.items():
            if yaml_key in raw_variant:
                merged[merged_key] = raw_variant[yaml_key]
        merged.setdefault("_binding_key", "")
        merged.setdefault("_binding_label", "")
        bindings.append(merged)
    return bindings


def _parse_all_profiles() -> list[dict]:
    profiles: list[dict] = []
    for yp in sorted(PROFILE_DIR.glob("*.yaml")):
        data = _load_yaml(yp)
        for variant in _shell_defaults_variants(data):
            binding_key = variant.get("_binding_key") or (data.get("transform") or {}).get("key")
            if not binding_key:
                continue
            profiles.append(
                {
                    "path": yp,
                    "data": data,
                    "binding_key": binding_key,
                    "binding_label": variant.get("_binding_label")
                    or (data.get("transform") or {}).get("label", ""),
                    "shell_defaults_override": variant,
                }
            )
    return profiles


def _apply_variant_to_profile(profile_data: dict, binding: dict | None) -> dict:
    """Return a shallow copy of profile_data with a variant's shell_defaults + source_layout overlay applied.

    validate_input takes a flat (transform.key, shell_defaults) view. For
    variant-backed bindings we:
      (1) swap the base shell_defaults for the per-variant merged dict
          (with source_* / path overrides resolved),
      (2) patch source_layout with any shell_defaults.source_* overrides
          (same precedence as build_runner / create_runner_shell: shell wins
          over declared source_layout).
    """
    if not binding:
        return profile_data
    variant_sd = binding.get("shell_defaults_override") or {}
    merged = dict(profile_data)
    new_shell = dict(profile_data.get("shell_defaults") or {})
    for k, v in variant_sd.items():
        if k.startswith("_binding"):
            continue
        new_shell[k] = v
    merged["shell_defaults"] = new_shell

    base_sl = dict(profile_data.get("source_layout") or {})
    sl_patch_map = [
        ("source_worksheet_name", "worksheet_name"),
        ("source_worksheet_resolution", "worksheet_resolution"),
        ("source_header_row", "header_row"),
        ("source_data_start_row", "data_start_row"),
    ]
    for shell_key, sl_key in sl_patch_map:
        val = variant_sd.get(shell_key)
        if val in (None, ""):
            continue
        try:
            if sl_key in ("header_row", "data_start_row"):
                base_sl[sl_key] = int(val)
            else:
                base_sl[sl_key] = str(val)
        except (TypeError, ValueError):
            continue
    merged["source_layout"] = base_sl
    return merged


def _resolve_profile(transform_key: str) -> tuple[Path, dict]:
    for entry in _parse_all_profiles():
        if entry.get("binding_key") == transform_key:
            merged = _apply_variant_to_profile(entry["data"], entry)
            return entry["path"], merged
    base_keys = sorted(
        {
            ((e.get("data") or {}).get("transform") or {}).get("key", "")
            for e in _parse_all_profiles()
            if ((e.get("data") or {}).get("transform") or {}).get("key")
        }
    )
    all_keys = sorted({e.get("binding_key", "") for e in _parse_all_profiles() if e.get("binding_key")})
    raise SystemExit(
        f"ERROR: transform key {transform_key!r} not found under {PROFILE_DIR}.  "
        f"Known profiles: " + ", ".join(base_keys) + f".  Known bindings (including variants): " + ", ".join(all_keys)
    )


def _read_reference_paths() -> dict:
    return (_load_yaml(REF_PATHS_PATH).get("reference_paths")) or {}


def _resolve_workbook_paths(profile: dict) -> dict[str, str]:
    """Resolve shell_defaults paths with fallback to reference_paths shared."""
    refs = _read_reference_paths()
    shell_defaults = profile.get("shell_defaults") or {}
    resolved: dict[str, str] = {}
    for shell_key, ref_key in [
        ("source_data", None),
        ("target_eib", None),
        ("output_file_path", None),
        ("mapping_spreadsheet", "mapping_spreadsheet"),
        ("mapping_sheet", "mapping_sheet"),
        ("supplier_spreadsheet", "supplier_spreadsheet"),
        ("customer_spreadsheet", "customer_spreadsheet"),
    ]:
        val = shell_defaults.get(shell_key)
        if not val and ref_key:
            val = refs.get(ref_key)
        resolved[shell_key] = str(val or "")
    return resolved


def _open_workbook(path_str: str, read_only: bool = True):
    """Try to open a workbook; return (wb_or_None, failure_msg).

    Very large source workbooks (AMEX etc. with 1M+ rows) should use read_only=True
    for memory efficiency.  Small reference workbooks (Workday supplier/customer
    extracts, mapping masters) must use read_only=False, because their pivots/
    charts/cache cause openpyxl in read_only mode to report rows=1 cols=1 even
    when the file is perfectly healthy.
    """
    if not path_str:
        return None, "path is blank (no shell_defaults configured)"
    p = Path(os.path.expanduser(path_str))
    if not p.exists():
        return None, f"file not found on disk: {p}"
    ext_ok = p.suffix.lower() in VALID_SPREADSHEET_EXTS
    if not ext_ok:
        return None, (
            f"file is not a supported spreadsheet type: extension {p.suffix.lower()!r} "
            f"not in sorted list: " + ", ".join(sorted(VALID_SPREADSHEET_EXTS))
        )
    try:
        if p.suffix.lower() == ".csv":
            return None, "CSV format not supported from this validator (please convert to .xlsx first, or Runner VBA will open it via OpenCsvAsText but this CLI does not parse CSVs)"
        wb = openpyxl.load_workbook(p, read_only=read_only, data_only=True)
    except Exception as e:  # pragma: no cover - openpyxl varies
        return None, f"failed to open as xlsx: {type(e).__name__}: {e}"
    return wb, ""


def _find_source_worksheet(wb, source_layout: dict):
    reso = source_layout.get("worksheet_resolution") or "named"
    explicit = source_layout.get("worksheet_name")
    if reso == "first_sheet" or (not explicit and reso != "named"):
        if not wb.sheetnames:
            return None, "workbook has no sheets"
        return wb[wb.sheetnames[0]], f"first_sheet -> {wb.sheetnames[0]!r}"
    # named lookup
    if explicit not in wb.sheetnames:
        return None, (
            f"sheet {explicit!r} not found; sheets present: "
            + ", ".join(repr(s) for s in wb.sheetnames[:20])
            + (" …" if len(wb.sheetnames) > 20 else "")
        )
    return wb[explicit], f"named lookup -> {explicit!r}"


def _scan_row_for_headers(ws, row_idx: int, headers: list[str]) -> dict[str, int]:
    """Return {header_spec_from_input_list: col_index_1based} for headers on the row.

    Each header entry may be a pipe-delimited alias list (e.g. "SL Date|SL Dates");
    ANY alias match counts as present for that header entry (mirrors VBA
    RowContainsHeader + SplitHeaderAliases).
    """
    want_tokens: dict[str, str] = {}  # normalized_alias -> original_header_spec_in_list
    for h in headers:
        for alias in _split_header_aliases(h):
            tok = _norm_token(alias)
            if tok:
                want_tokens.setdefault(tok, h)
    found: dict[str, int] = {}
    max_col = ws.max_column or 0
    for c in range(1, min(max_col + 1, 1025)):
        tok = _norm_token(ws.cell(row_idx, c).value)
        if tok in want_tokens:
            spec = want_tokens[tok]
            found.setdefault(spec, c)
            if len(found) == len(headers):
                break
    return found


def _first_matching_row(ws, headers: list[str], max_rows: int):
    """Find first row in 1..max_rows where EVERY header spec has >= 1 alias matched.

    Mirrors VBA FindHeaderRow scan_required_headers behaviour: each header entry
    may be pipe-delimited; any one alternative counts as present for that entry.
    """
    want_by_header: list[set[str]] = [
        {_norm_token(a) for a in _split_header_aliases(h) if _norm_token(a)}
        for h in headers
    ]
    total_want = len(headers)
    rows = ws.max_row or 0
    cols = ws.max_column or 0
    for r in range(1, min(rows, max_rows) + 1):
        row_tokens: set[str] = set()
        for c in range(1, min(cols, 1025) + 1):
            t = _norm_token(ws.cell(r, c).value)
            if t:
                row_tokens.add(t)
        matched = 0
        for alias_set in want_by_header:
            if row_tokens & alias_set:
                matched += 1
                if matched == total_want:
                    return r
    return None


def _header_row_candidates(ws, source_layout: dict, required_headers: list[str]):
    detection = source_layout.get("header_detection") or "scan_required_headers"
    literal = source_layout.get("header_row")
    if detection == "literal_row" and literal:
        return "literal_row", int(literal)
    if detection == "scan_required_headers":
        r = _first_matching_row(ws, required_headers, MAX_HEADER_SCAN_ROWS)
        if r is None:
            return "scan_required_headers", literal if literal else None
        return "scan_required_headers", r
    # fallback: use literal if declared, else fail
    return detection, literal if literal else None


def _column_nonblank_counts(ws, header_cols: dict[str, int], data_start_row: int):
    """Count non-blanks per required header across a real-data sample.

    Excel workbooks often have their `used range` extend to row 1_048_576 even
    though the real payload ends around row 800.  Rather than sampling from
    the declared data_start blindly (which can land in the trailing empty
    millions), we:
      1. Find first real-data row at or after data_start_row within a
         SEARCH_WINDOW of 2000 rows.
      2. From there, sample up to MAX_DATA_SAMPLE_ROWS contiguous rows.
    """
    SEARCH_WINDOW = 5000
    rows_total = ws.max_row or 0
    # Step 1: find first non-empty row within search window
    first_nonempty = None
    search_end = min(rows_total, data_start_row + SEARCH_WINDOW - 1)
    for r in range(data_start_row, search_end + 1):
        for hc in header_cols.values():
            if _norm_token(ws.cell(r, hc).value) != "":
                first_nonempty = r
                break
        if first_nonempty:
            break
    if first_nonempty is None:
        # No data anywhere in first 5k rows -> column counts are all 0
        return {h: [0, 0] for h in header_cols}, data_start_row, data_start_row
    sample_end = min(rows_total, first_nonempty + MAX_DATA_SAMPLE_ROWS - 1)
    counts: dict[str, list[int]] = {h: [0, 0] for h in header_cols}
    cols_sorted = sorted(header_cols.values())
    for r in range(first_nonempty, sample_end + 1):
        entirely_blank_row = True
        row_result_per_col: list[tuple[str, bool]] = []
        for h, c in header_cols.items():
            v = ws.cell(r, c).value
            blank = _norm_token(v) == ""
            if not blank:
                entirely_blank_row = False
            row_result_per_col.append((h, blank))
        if entirely_blank_row:
            continue  # trailer
        for h, blank in row_result_per_col:
            counts[h][0 if blank else 1] += 1
    return counts, first_nonempty, sample_end


def _mapping_sheet_columns_ok(mapping_ws) -> tuple[bool, str]:
    canonical = [
        "OLD CODE",
        "OLD SEGMENT",
        "Core vs Prism",
        "WORKDAY BUSINESS OBJECT",
        "WORKDAY REF ID",
    ]
    found = _scan_row_for_headers(mapping_ws, 1, canonical)
    if len(found) == len(canonical):
        return True, f"row 1 contains all 5 canonical mapping columns"
    # try rows 1-8
    for r in range(1, 9):
        f = _scan_row_for_headers(mapping_ws, r, canonical)
        if len(f) == len(canonical):
            return True, f"row {r} contains all 5 canonical mapping columns"
    return False, (
        "mapping sheet missing canonical columns: missing "
        + ", ".join(sorted(set(canonical) - set(found.keys())))
    )


def _ref_sheet_columns_ok(ref_wb, entity_label: str, id_col: str) -> tuple[bool, str]:
    # Pick first sheet; look for exact Reference ID + entityLabel+ID col pair.
    # Workday extracts typically use exact headers:
    #   Supplier extract -> "Reference ID" + "Supplier ID"
    #   Customer extract -> "Reference ID" + "Customer ID"
    # The alternate permutation ("Supplier" bare token + "Reference ID") is also
    # accepted for compatibility with Finance-staged extracts.
    # Note: Workday extracts typically have 3-6 filter rows before the real header
    # row (e.g. row 1 = title, rows 2-3 = Created Date From/To, row 4 = Supplier Name,
    # row 5 = Supplier ID, row 6+ = Reference ID + other columns).  Search rows 1-25.
    SEARCH_ROWS = 25
    if not ref_wb.sheetnames:
        return False, "ref workbook has no sheets"
    ws = ref_wb[ref_wb.sheetnames[0]]
    exact_pairs = [
        ("Reference ID", f"{entity_label} {id_col}"),
    ]
    for r in range(1, SEARCH_ROWS + 1):
        # Grab all tokens on this row
        row_tokens: dict[str, int] = {}
        for c in range(1, min(ws.max_column or 0, 251)):
            tok = _norm_token(ws.cell(r, c).value)
            if not tok:
                continue
            row_tokens.setdefault(tok, c)
        for left, right in exact_pairs:
            if _norm_token(left) in row_tokens and _norm_token(right) in row_tokens:
                return True, (
                    f"sheet {ws.title!r} row {r} has headers: "
                    f"{left} (col {row_tokens[_norm_token(left)]}), "
                    f"{right} (col {row_tokens[_norm_token(right)]})"
                )
        # Accept a looser match (Reference ID + any header that ends with " ID")
        if _norm_token("Reference ID") in row_tokens:
            any_id = [k for k in row_tokens if k.endswith(" id")]
            if any_id:
                return True, (
                    f"sheet {ws.title!r} row {r} has Reference ID (col "
                    f"{row_tokens[_norm_token('Reference ID')]}) + ID-like headers: "
                    + ", ".join(sorted(any_id)[:6])
                )
    return False, (
        f"sheet {ws.title!r} rows 1-{SEARCH_ROWS} missing both 'Reference ID' + "
        f"'{entity_label} {id_col}' expected headers"
    )


def _count_nonempty_rows(ws, start_row: int, max_row: int) -> int:
    """Count rows from start_row..max_row that have at least one non-empty cell in col 1..max_col."""
    total = 0
    cols_n = min(ws.max_column or 0, 100)
    for r in range(start_row, min(max_row, ws.max_row or 0) + 1):
        for c in range(1, cols_n + 1):
            if _norm_token(ws.cell(r, c).value) != "":
                total += 1
                break
    return total


def _matches_allowed_all_blank(header_spec: str, allowed_list: list[str]) -> bool:
    """Return True if any alias of header_spec matches any alias in allowed_list.

    Header specs on both sides may be pipe-delimited alias lists; we expand
    both sides and compare normalized tokens (case-insensitive, trimmed).
    """
    if not allowed_list:
        return False
    spec_aliases = {_norm_token(a) for a in _split_header_aliases(header_spec)}
    for allowed in allowed_list:
        for alias in _split_header_aliases(allowed):
            if _norm_token(alias) in spec_aliases:
                return True
    return False


def run_checks(args: argparse.Namespace) -> tuple[list[CheckResult], list[str]]:
    results: list[CheckResult] = []
    infos: list[str] = []

    profile_path, profile_data = _resolve_profile(args.transform)
    source_layout = profile_data.get("source_layout") or {}
    required_headers_raw = source_layout.get("required_headers") or []
    required_headers: list[str] = [str(h) for h in required_headers_raw]
    optional_headers_raw = source_layout.get("optional_headers") or []
    optional_headers: list[str] = [str(h) for h in optional_headers_raw]
    allowed_all_blank_raw = source_layout.get("allowed_all_blank_required_headers") or []
    allowed_all_blank: list[str] = [str(h) for h in allowed_all_blank_raw]
    blank_header_cols_raw = source_layout.get("blank_header_columns") or []
    blank_header_cols: list[tuple[str, int]] = []
    if isinstance(blank_header_cols_raw, list):
        for entry in blank_header_cols_raw:
            if not isinstance(entry, dict): continue
            name = str(entry.get("display_name") or entry.get("name") or "(unnamed blank-header column)")
            letter = str(entry.get("column_letter") or "")
            col_idx = _column_letter_to_index(letter) if letter else 0
            if col_idx > 0:
                blank_header_cols.append((name, col_idx))
    paths = _resolve_workbook_paths(profile_data)

    # CLI overrides
    for attr, key in [
        ("source_data", "source_data"),
        ("target_eib", "target_eib"),
        ("mapping_spreadsheet", "mapping_spreadsheet"),
        ("mapping_sheet", "mapping_sheet"),
        ("supplier_spreadsheet", "supplier_spreadsheet"),
        ("customer_spreadsheet", "customer_spreadsheet"),
    ]:
        override = getattr(args, attr, None)
        if override:
            paths[key] = override

    eib_type = ((profile_data.get("transform") or {}).get("eib_population_type") or "").lower()
    is_ap = "supplier" in eib_type
    is_ar = "customer" in eib_type

    infos.append(f"Profile          : {profile_path}")
    infos.append(f"Label            : {(profile_data.get('transform') or {}).get('label', '')}")
    infos.append(f"EIB type         : {(profile_data.get('transform') or {}).get('eib_population_type', '')}")
    infos.append(f"Source data      : {paths['source_data']}")
    infos.append(f"Target EIB tpl   : {paths['target_eib']}")
    infos.append(f"Mapping wb/sheet : {paths['mapping_spreadsheet']} / {paths['mapping_sheet'] or '<default>'}")
    if is_ap:
        infos.append(f"Supplier ref wb  : {paths['supplier_spreadsheet']}")
    if is_ar:
        infos.append(f"Customer ref wb  : {paths['customer_spreadsheet']}")

    # --- Source workbook
    src_wb, src_msg = _open_workbook(paths["source_data"], read_only=True)
    if src_wb is None:
        results.append(CheckResult("FILE SOURCE DATA", False, src_msg))
        return results, infos
    results.append(
        CheckResult(
            "FILE SOURCE DATA",
            True,
            f"opened OK, sheets={len(src_wb.sheetnames)}",
        )
    )

    src_ws, ws_msg = _find_source_worksheet(src_wb, source_layout)
    if src_ws is None:
        results.append(CheckResult("SHEET SOURCE LAYOUT", False, ws_msg))
        return results, infos
    results.append(CheckResult("SHEET SOURCE LAYOUT", True, ws_msg))

    dims_ok = (src_ws.max_row or 0) > 0 and (src_ws.max_column or 0) > 0
    results.append(
        CheckResult(
            "DIMENSIONS SOURCE",
            dims_ok,
            f"rows={src_ws.max_row}, cols={src_ws.max_column}",
        )
    )
    if not dims_ok:
        return results, infos

    # --- Header row detect
    detection, header_row = _header_row_candidates(src_ws, source_layout, required_headers)
    declared = source_layout.get("header_row")
    if declared and header_row is not None:
        header_row_matches_declared = int(declared) == header_row
    else:
        header_row_matches_declared = header_row is not None
    results.append(
        CheckResult(
            "HEADER ROW DETECT",
            header_row_matches_declared and header_row is not None,
            f"mode={detection}; found row={header_row}"
            + (
                f" (declared literal={declared})"
                if declared
                else ""
            ),
        )
    )
    if header_row is None:
        # Still continue — try declared or row 1 for reporting
        header_row = int(declared) if declared else 1

    # --- Required headers per-column
    req_cols = _scan_row_for_headers(src_ws, header_row, required_headers)
    missing = [h for h in required_headers if h not in req_cols]
    rh_ok = len(missing) == 0
    rh_detail_parts = []
    for h in required_headers:
        if h in req_cols:
            rh_detail_parts.append(f"✅ {h} @ col {req_cols[h]}")
        else:
            rh_detail_parts.append(f"❌ {h}  <MISSING>")
    # --- Blank-header mandatory columns (e.g., col AH Hotel Name = blank header in raw AMEX export)
    bad_blank = []
    for bh_name, bh_col in blank_header_cols:
        bh_maxcol = src_ws.max_column or 0
        if bh_col > bh_maxcol:
            bad_blank.append(f"❌ {bh_name} (col letter => {bh_col}) exceeds worksheet cols={bh_maxcol}")
            rh_ok = False
        else:
            req_cols[f"[blank-header] {bh_name}"] = bh_col
            rh_detail_parts.append(f"✅ {bh_name} (blank header row {header_row}, resolved column letter) @ col {bh_col}")
    for b in bad_blank:
        rh_detail_parts.append(b)
    if missing or bad_blank:
        all_missing = missing + [b.split("<")[1].split(">")[0] if "<MISSING>" in b else b for b in bad_blank]
        rh_detail_parts.append(f"Summary: missing {len(missing) + len(bad_blank)}/{len(required_headers) + len(blank_header_cols)} binding(s): " + ", ".join(missing + [x.split(" (")[0].replace("❌ ","") for x in bad_blank]))
    else:
        rh_detail_parts.append(f"Summary: all {len(required_headers) + len(blank_header_cols)} required bindings ({len(required_headers)} named headers + {len(blank_header_cols)} blank-header letter refs) resolved on row {header_row}")
    results.append(
        CheckResult(
            "REQUIRED HEADERS",
            rh_ok,
            "\n    ".join(rh_detail_parts),
        )
    )

    # --- Optional headers (informational, not pass/fail)
    if optional_headers:
        opt_cols = _scan_row_for_headers(src_ws, header_row, optional_headers)
        parts = []
        for h in optional_headers:
            if h in opt_cols:
                parts.append(f"ℹ️  {h} @ col {opt_cols[h]}")
            else:
                parts.append(f"ℹ️  {h}  <absent (optional, ok)>")
        infos.append("Optional headers :\n    " + "\n    ".join(parts))

    # --- Required cols nonblank-count sanity
    data_start = source_layout.get("data_start_row")
    if isinstance(data_start, int) and req_cols:
        counts, ds, se = _column_nonblank_counts(src_ws, req_cols, data_start)
        warn_rows = []
        tolerate_bits = []
        for h, (bl, nonbl) in sorted(counts.items()):
            total = bl + nonbl
            tolerated = _matches_allowed_all_blank(h, allowed_all_blank)
            if total == 0:
                if tolerated:
                    tolerate_bits.append(f"ℹ️  {h}: zero rows (tolerated: in allowed_all_blank_required_headers)")
                else:
                    warn_rows.append(f"⚠️  {h}: zero rows of data in sample (empty column?)")
            elif nonbl == 0:
                if tolerated:
                    tolerate_bits.append(f"ℹ️  {h}: 100% blank in first sample ({se - ds + 1} rows from row {ds} (tolerated: in allowed_all_blank_required_headers)")
                else:
                    warn_rows.append(
                        f"⚠️  {h}: {bl}/{total} blank rows in first sample "
                        f"({se - ds + 1} rows from row {ds})"
                    )
            elif bl / total > 0.25:
                # 25..99% blank rows: warn regardless of tolerate list — partial blanks are
                # always worth flagging (they indicate either trailers or partial data-entry corruption
                warn_rows.append(
                    f"⚠️  {h}: {bl}/{total} blank rows in first sample "
                    f"({se - ds + 1} rows from row {ds})"
                )
        if warn_rows:
            detail_lines = warn_rows + tolerate_bits
            results.append(CheckResult("REQUIRED COLS NON-ZERO", False, "\n    ".join(detail_lines)))
        else:
            summary_bits = [f"{h}={counts[h][1]}/{sum(counts[h])} nonblank" for h in sorted(counts)]
            summary_line = "; ".join(summary_bits)
            if tolerate_bits:
                summary_line += "\n    " + "\n    ".join(tolerate_bits)
            results.append(
                CheckResult(
                    "REQUIRED COLS NON-ZERO",
                    True,
                    summary_line,
                )
            )

    src_wb.close()

    # --- Target EIB
    tpl_wb, tpl_msg = _open_workbook(paths["target_eib"], read_only=False)
    results.append(
        CheckResult(
            "FILE TEMPLATE EIB",
            tpl_wb is not None,
            tpl_msg if tpl_wb is None else f"opened OK, sheets={len(tpl_wb.sheetnames)}",
        )
    )
    if tpl_wb is not None:
        tpl_wb.close()

    # --- Mapping workbook + sheet
    map_wb, map_msg = _open_workbook(paths["mapping_spreadsheet"], read_only=False)
    msheet_name = paths["mapping_sheet"] or "MAPPING MASTER"
    if map_wb is None:
        results.append(CheckResult("FILE MAPPING", False, map_msg))
    else:
        results.append(
            CheckResult(
                "FILE MAPPING",
                True,
                f"opened OK, sheets={len(map_wb.sheetnames)}",
            )
        )
        if msheet_name not in map_wb.sheetnames:
            results.append(
                CheckResult(
                    "SHEET MAPPING",
                    False,
                    f"sheet {msheet_name!r} missing; have: "
                    + ", ".join(map_wb.sheetnames[:15]),
                )
            )
        else:
            mws = map_wb[msheet_name]
            mcols_ok, mcols_msg = _mapping_sheet_columns_ok(mws)
            results.append(
                CheckResult(
                    "SHEET MAPPING COLUMNS",
                    mcols_ok,
                    mcols_msg + (
                        f" (rows={mws.max_row}, cols={mws.max_column})"
                        if mcols_ok
                        else ""
                    ),
                )
            )
        map_wb.close()

    # --- Supplier extract (AP only)
    if is_ap:
        sup_wb, sup_msg = _open_workbook(paths["supplier_spreadsheet"], read_only=False)
        if sup_wb is None:
            results.append(CheckResult("FILE SUPPLIER", False, sup_msg))
        else:
            results.append(
                CheckResult(
                    "FILE SUPPLIER",
                    True,
                    f"opened OK, sheets={len(sup_wb.sheetnames)}",
                )
            )
            sok, smsg = _ref_sheet_columns_ok(sup_wb, "Supplier", "ID")
            results.append(CheckResult("SHEET SUPPLIER COLS", sok, smsg))
            sup_wb.close()

    # --- Customer extract (AR only)
    if is_ar:
        cust_wb, cust_msg = _open_workbook(paths["customer_spreadsheet"], read_only=False)
        if cust_wb is None:
            results.append(CheckResult("FILE CUSTOMER", False, cust_msg))
        else:
            results.append(
                CheckResult(
                    "FILE CUSTOMER",
                    True,
                    f"opened OK, sheets={len(cust_wb.sheetnames)}",
                )
            )
            cok, cmsg = _ref_sheet_columns_ok(cust_wb, "Customer", "ID")
            results.append(CheckResult("SHEET CUSTOMER COLS", cok, cmsg))
            cust_wb.close()

    return results, infos


def create_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=(
            "Validate a transform's input workbooks (source, mapping, supplier/customer) "
            "against the schema declared in its YAML profile.  Exit 0 when all checks "
            "pass, 1 otherwise."
        )
    )
    p.add_argument("-t", "--transform", required=True, help="Profile transform key, e.g. ap_amex_hotel")
    p.add_argument("--source-data", help="Override shell_defaults source_data path")
    p.add_argument("--target-eib", help="Override shell_defaults target_eib path")
    p.add_argument("--mapping-spreadsheet", help="Override shell_defaults mapping_spreadsheet path")
    p.add_argument("--mapping-sheet", help="Override mapping_sheet name inside mapping workbook")
    p.add_argument("--supplier-spreadsheet", help="Override shell_defaults supplier_spreadsheet path")
    p.add_argument("--customer-spreadsheet", help="Override shell_defaults customer_spreadsheet path")
    return p


def main(argv: list[str] | None = None) -> int:
    parser = create_parser()
    args = parser.parse_args(argv)
    results, infos = run_checks(args)
    if infos:
        print("== Profile & Paths ==")
        for line in infos:
            print(f"  {line}")
        print()
    print("== Validation Checks ==")
    failures = 0
    for r in results:
        symbol = "✅" if r.ok else "❌"
        if not r.ok:
            failures += 1
        print(f"  {symbol} {r.name:24s} -> {r.detail}")
    print()
    if failures:
        print(f"VERDICT: ❌ {failures} check(s) failed.")
        return 1
    print("VERDICT: ✅ All checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
