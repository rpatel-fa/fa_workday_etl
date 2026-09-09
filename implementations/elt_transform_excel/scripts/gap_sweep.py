#!/usr/bin/env python3
"""Gap sweep: cross-reference every inferred constant in 9 transform .bas
modules against the profile + spec YAMLs + eib_population_defaults.

Reports:
  - FILLED: constant exists in some YAML
  - GAP:    hardcoded constant ONLY exists in .bas (config bug per user rule)
  - JUDGEMENT: tenant-level Ref-ID / company-entity maps (Prompt 3 territory)
"""
import sys, pathlib, re, json, importlib.util

ROOT = pathlib.Path(__file__).resolve().parent.parent

spec = importlib.util.spec_from_file_location("build_runner", ROOT / "scripts" / "build_runner.py")
br = importlib.util.module_from_spec(spec)
sys.modules["build_runner"] = br
spec.loader.exec_module(br)
yaml_mod = br._import_yaml()

TRIO_STEMS = [
    "ap_amex_hotel",
    "ap_amex_trains",
    "ap_amex_trains_v2",
    "ap_cfa_distributions",
    "ap_prize_funds",
    "ap_referees",
    "ar_discipline_invoice",
    "ar_invoice",
    "ar_invoice_run",
]

BAS_MAP = {
    "ap_amex_hotel":         "modApAmexHotelSupplierInvoice.bas",
    "ap_amex_trains":        "modApAmexTrainsSupplierInvoice.bas",
    "ap_amex_trains_v2":     "modApAmexTrainV2SupplierInvoice.bas",
    "ap_cfa_distributions":  "modApCfaDistribSupplierInvoice.bas",
    "ap_prize_funds":        "modApPrizeFundSupplierInvoice.bas",
    "ap_referees":  "modApRefereeSupplierInvoice.bas",
    "ar_discipline_invoice": "modArDisciplineCustInvoice.bas",
    "ar_invoice":            "modArInvoiceCustomerInvoice.bas",
    "ar_invoice_run":        "modArInvoiceRunCustInvoice.bas",
}

EIB_KNOWN = {
    "Submit Supplier Invoice",
    "Submit Customer Invoice",
}

# Ref-ID lookups, tenant-wide shared engine files, helper paths:
# These are *intentionally* not per-profile config but shared engine items.
SHARED_ENGINE_CONSTS = {
    "RUNNER_SHEET_NAME",        # modHelpers / dispatcher constant (gate-enforced)
    "MAPPING_SHEET_NAME",       # shared
    "CONFIG_BANNER_SHEET_NAME", # shared
    "REGISTRY_SHEET_NAME",      # shared
    "SUPPLIER_REF_SHEET",       # from reference_paths
    "CUSTOMER_REF_SHEET",       # from reference_paths
    "WORKDAY_REF_ID_SHEET",     # from reference_paths
    "REGISTRY_START_COLUMN",    # shared dispatcher constant
    "REGISTRY_COL_PROFILE_KEY", # shared dispatcher constant
    "REGISTRY_COL_STEP_LIST",   # shared dispatcher constant
}

# --- Helper: flatten YAML dict into "key=value" set for quick membership check ---
def flatten_yaml(obj, prefix=""):
    out = set()
    out_strs = set()
    if isinstance(obj, dict):
        for k, v in obj.items():
            newp = f"{prefix}.{k}" if prefix else k
            out.add((newp, str(v) if not isinstance(v, (dict, list)) else type(v).__name__))
            if isinstance(v, str):
                out_strs.add(v)
            a, b = flatten_yaml(v, newp)
            out |= a; out_strs |= b
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            newp = f"{prefix}[{i}]"
            if isinstance(v, str):
                out_strs.add(v)
            a, b = flatten_yaml(v, newp)
            out |= a; out_strs |= b
    else:
        out.add((prefix, str(obj)))
        if isinstance(obj, str):
            out_strs.add(obj)
    return out, out_strs

# --- Load all config ---
eib_def_path = ROOT / "config" / "eib_population_defaults.yaml"
eib_defs = yaml_mod.safe_load(eib_def_path.read_text())
eib_flat_keys, eib_flat_strs = flatten_yaml(eib_defs)

ref_path = ROOT / "config" / "reference_paths.yaml"
ref_paths = yaml_mod.safe_load(ref_path.read_text()) if ref_path.exists() else {}
ref_flat_keys, ref_flat_strs = flatten_yaml(ref_paths)

# ingest lookups (ref_company.json etc) - strings in these are tenant maps, not config gaps
INGEST_LOOKUP_STRS = set()
for jf in (ROOT / "config" / "ingest_lookups").glob("*.json"):
    try:
        data = json.loads(jf.read_text())
        _, s = flatten_yaml(data)
        INGEST_LOOKUP_STRS |= s
    except Exception:
        pass

# Regex: Private Const NAME = "value" or Private Const NAME = 123
RE_PRIVATE_CONST = re.compile(r'^\s*Private\s+Const\s+(\w+)\s+(?:As\s+\w+\s+)?=\s*(.+?)\s*$', re.MULTILINE)
# Hardcoded sheet names inside Worksheets("...") or Sheets("...")
RE_WS_LITERAL = re.compile(r'(?:Worksheets|Sheets)\s*\(\s*"([^"]+)"\s*\)')
# Hardcoded string literals assigned to target column vars: targetXxxSheet = "..."
RE_TARGET_SHEET = re.compile(r'(?:target\w*Sheet|eib\w*Sheet|submit\w*Sheet|submitSheet)\s*=\s*"([^"]+)"', re.IGNORECASE)
# Hardcoded "GBP", "FA", "WNSL" etc assignments
RE_COMPANY_CONST = re.compile(r'(?:company(?:Code)?|entity(?:Code)?)\s*=\s*"([A-Z0-9_]+)"', re.IGNORECASE)
RE_CURRENCY_CONST = re.compile(r'(?:currency(?:Code)?)\s*=\s*"([A-Z]{3})"', re.IGNORECASE)
# Hardcoded .NumberFormat = "@" (just detect the pattern, not a gap actually per C2)
RE_DATE_FMT = re.compile(r'\.NumberFormat\s*=\s*"([^"]+)"')
# Magic row/column numbers (startRow = 5 etc)
RE_ROW_ASSIGN = re.compile(r'(?:startRow|dataStartRow|headerRow|startCol|firstDataRow)\s*(?:=\s*|\()(\d+)', re.IGNORECASE)

def check_in_config(value_str, profile_flat_strs, spec_flat_strs):
    """Check membership in all 4 config layers + ingest lookups."""
    vs = value_str.strip()
    if not vs:
        return "EMPTY"
    if vs in eib_flat_strs:
        return "FILLED[eib_pop_defaults]"
    if vs in ref_flat_strs:
        return "FILLED[reference_paths]"
    if vs in profile_flat_strs:
        return "FILLED[profile]"
    if vs in spec_flat_strs:
        return "FILLED[transform_spec]"
    if vs in INGEST_LOOKUP_STRS:
        return "TENANT_MAP[ingest_lookup]"
    if vs in EIB_KNOWN:
        return "GAP! eib target sheet literal in .bas NOT propagated to shared config reference (use eib_defaults via build)"
    return "GAP! not found in any config layer"

report_lines = []
report_lines.append("=" * 100)
report_lines.append("TRANSFORM CONFIG GAP SWEEP — 9 trios (.bas + profile.yaml + spec.yaml)")
report_lines.append("Rule: ANY non-shared-engine hardcoded constant in .bas MUST exist in some YAML config layer")
report_lines.append("=" * 100)

totals = {"filled": 0, "gap": 0, "tenant": 0, "shared_engine": 0}

for stem in TRIO_STEMS:
    report_lines.append("")
    report_lines.append("-" * 100)
    report_lines.append(f"TRIO: {stem}")
    report_lines.append("-" * 100)

    profile_path = ROOT / "config" / "transforms" / f"{stem}.yaml"
    spec_path = ROOT / "config" / "transform_specs" / f"{stem}.yaml"
    bas_path = ROOT / "vba_src" / "transforms" / BAS_MAP[stem]

    profile = yaml_mod.safe_load(profile_path.read_text())
    try:
        spec = yaml_mod.safe_load(spec_path.read_text()) if spec_path.exists() else {}
    except yaml_mod.YAMLError as e:
        report_lines.append(f"  ⚠ transform_spec parse error (docs-only, not a build gap): {e}")
        spec = {}
    bas_text = bas_path.read_text(encoding="latin-1", errors="replace")

    _, profile_strs = flatten_yaml(profile)
    _, spec_strs = flatten_yaml(spec)

    # --- Private Consts ---
    pcs = list(RE_PRIVATE_CONST.finditer(bas_text))
    if pcs:
        report_lines.append(f"  Private Consts ({len(pcs)}):")
        for m in pcs:
            name, raw_val = m.group(1), m.group(2).strip()
            # Strip VBA line comments — everything after first ' preceded by whitespace is comment
            # This prevents "' Rule 1 -- hardcoded; CFA is UK-only FA" suffix being included in value
            if "'" in raw_val:
                m2 = re.search(r"\s+'\s", raw_val)
                if m2:
                    raw_val = raw_val[:m2.start()].strip()
            # Strip string quotes
            if raw_val.startswith('"') and raw_val.endswith('"'):
                val = raw_val[1:-1]
                kind = "str"
            else:
                val = raw_val
                kind = "num"
            if name in SHARED_ENGINE_CONSTS:
                status = "SHARED_ENGINE"
                totals["shared_engine"] += 1
            else:
                status = check_in_config(val, profile_strs, spec_strs) if kind == "str" else f"NUMERIC({val}) — check if row/col enum in source_layout"
                if status.startswith("FILLED"):
                    totals["filled"] += 1
                elif status.startswith("TENANT"):
                    totals["tenant"] += 1
                elif status.startswith("GAP"):
                    totals["gap"] += 1
            report_lines.append(f"    {'✓' if 'FILLED' in status or 'SHARED' in status or 'TENANT' in status else '✗'}  {name:32s} = {raw_val:40s}  -> {status}")

    # --- Worksheet literal access ---
    ws_lits = RE_WS_LITERAL.findall(bas_text)
    if ws_lits:
        report_lines.append(f"  Worksheet(\"...\") literals:")
        seen = set()
        for w in ws_lits:
            if w in seen: continue
            seen.add(w)
            status = check_in_config(w, profile_strs, spec_strs)
            if status.startswith("FILLED"): totals["filled"] += 1
            elif status.startswith("TENANT"): totals["tenant"] += 1
            elif status.startswith("GAP"): totals["gap"] += 1
            mark = "✓" if ("FILLED" in status or "SHARED" in status or "TENANT" in status) else "✗"
            report_lines.append(f"    {mark}  Worksheets(\"{w}\") -> {status}")

    # --- targetSheet = "..." patterns ---
    tgt_sheets = RE_TARGET_SHEET.findall(bas_text)
    if tgt_sheets:
        report_lines.append(f"  targetXXXSheet = \"...\" assignments:")
        seen = set()
        for w in tgt_sheets:
            if w in seen: continue
            seen.add(w)
            status = check_in_config(w, profile_strs, spec_strs)
            if status.startswith("FILLED"): totals["filled"] += 1
            elif status.startswith("TENANT"): totals["tenant"] += 1
            elif status.startswith("GAP"): totals["gap"] += 1
            mark = "✓" if ("FILLED" in status or "TENANT" in status) else "✗"
            report_lines.append(f"    {mark}  sheet={w!r} -> {status}")

    # --- Company / currency hardcodes ---
    for pattern, label in ((RE_COMPANY_CONST, "company"), (RE_CURRENCY_CONST, "currency")):
        hits = pattern.findall(bas_text)
        if hits:
            report_lines.append(f"  {label} hardcodes:")
            seen = set()
            for w in hits:
                if w in seen: continue
                seen.add(w)
                status = check_in_config(w, profile_strs, spec_strs)
                if status.startswith("FILLED"): totals["filled"] += 1
                elif status.startswith("TENANT"): totals["tenant"] += 1
                elif status.startswith("GAP"): totals["gap"] += 1
                mark = "✓" if ("FILLED" in status or "TENANT" in status) else "✗"
                report_lines.append(f"    {mark}  {label}={w!r} -> {status}")

    # --- Magic numeric row assignments ---
    rows = RE_ROW_ASSIGN.findall(bas_text)
    if rows:
        report_lines.append(f"  Numeric row/col literals (potentially source_layout ints):")
        seen_counts = {}
        for r in rows:
            seen_counts[r] = seen_counts.get(r, 0) + 1
        # Profile has header_row / data_start_row
        sl = profile.get("source_layout", {}) if isinstance(profile, dict) else {}
        known_ints = set()
        for k in ("header_row", "data_start_row"):
            if sl.get(k) is not None: known_ints.add(str(sl[k]))
        # Also eib_defaults target_data_start_row for EIB type
        ep = profile.get("eib_population_type") if isinstance(profile, dict) else None
        if ep and ep in eib_defs.get("eib_populations", {}):
            tdsr = eib_defs["eib_populations"][ep].get("target_data_start_row")
            if tdsr is not None: known_ints.add(str(tdsr))
        for r, c in sorted(seen_counts.items()):
            status = "FILLED[profile.source_layout or eib_defaults]" if r in known_ints else f"NUMERIC_GAP literal={r} (x{c}) — not in profile/eib_defaults ints"
            if "FILLED" in status: totals["filled"] += 1
            else: totals["gap"] += 1
            mark = "✓" if "FILLED" in status else "✗"
            report_lines.append(f"    {mark}  row/col literal {r} x{c} -> {status}")

# --- Summary ---
report_lines.append("")
report_lines.append("=" * 100)
report_lines.append("SUMMARY")
report_lines.append("=" * 100)
report_lines.append(f"  FILLED (in YAML):          {totals['filled']}")
report_lines.append(f"  TENANT MAP (ingest JSON):  {totals['tenant']}")
report_lines.append(f"  SHARED ENGINE CONST:       {totals['shared_engine']}")
report_lines.append(f"  GAP (only in .bas):        {totals['gap']}")
report_lines.append("")
report_lines.append("  GAP items require YAML backfill per user's single-source-of-truth rule.")
report_lines.append("  If gap count > 0: extend schema + backfill profiles + add validator gates.")

out = "\n".join(report_lines)
print(out)
report_path = ROOT / "_reports" / "gap_sweep_report.txt"
report_path.parent.mkdir(parents=True, exist_ok=True)
report_path.write_text(out + "\n")
print(f"\nFull report saved to: {report_path}")
