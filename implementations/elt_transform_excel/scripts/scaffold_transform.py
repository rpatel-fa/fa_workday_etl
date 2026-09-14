#!/usr/bin/env python3
"""Create a skeleton new transform by forking the closest structural exemplar.

This is H2.1 of the cold-start-scale backlog. It does NOT infer column mappings
from source data (that's H3). It DOES solve the copy/paste/rename mechanics for
the 3 artifacts a new transform requires, so a coding agent (or a human) can
concentrate on the 2 judgement steps (pick exemplar, classify columns) and the
editing of the skeleton's TO_FILL_* markers.

3 artifacts per transform:
    1. config/transforms/<profile_key>.yaml      (Runner registration + shell paths)
    2. config/transform_specs/<profile_key>.yaml (column mapping spec, optional)
    3. vba_src/transforms/mod<Label><EIBType>.bas (actual VBA step list)

The script enforces the 31-character VBComponent limit BEFORE writing files,
uses CRLF line endings for .bas (Mac Excel VBA import requires CRLF), writes
pure ASCII, and marks every location a human/agent must edit with a sentinel
TO_FILL_<THING> token the user can grep for.

Usage examples:

    # Fork the ap_amex_trains (v1, simple per-row lookup) profile to a new one:
    uv run --with pyyaml --no-project python \\
        elt_transform_excel/scripts/scaffold_transform.py \\
        --fork-from ap_amex_trains \\
        --new-key ap_bacardi_uk \\
        --new-label "AP Bacardi UK" \\
        --eib-type "Supplier Invoice" \\
        --source-type "Bacardi UK GL Raw"

    # List known exemplar keys with their structural archetype:
    uv run --with pyyaml --no-project python \\
        elt_transform_excel/scripts/scaffold_transform.py --list
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

SCRIPT_PATH = Path(__file__).resolve()
MODULE_ROOT = SCRIPT_PATH.parents[1]
TRANSFORMS_DIR = MODULE_ROOT / "config" / "transforms"
TRANSFORM_SPECS_DIR = MODULE_ROOT / "config" / "transform_specs"
VBA_TRANSFORMS_DIR = MODULE_ROOT / "vba_src" / "transforms"

MAX_VBA_COMPONENT_NAME_CHARS = 31

# Archetype labels used by the 'Given a source file + EIB type' inference
# playbook in docs/todo/archive/cold-start-scale.md § I. Each exemplar key maps to the
# archetype a cold-start agent should consider when picking 'closest existing'
# to fork from.
ARCHETYPES: dict[str, str] = {
    "ap_amex_hotel": "Fixed constants / no per-row lookup (Supplier Inv archetype A)",
    "ap_amex_trains": "Per-row live mapping (Supplier Inv archetype B)",
    "ap_amex_trains_v2": "Raw derivation + per-row lookup (Amex Trains V2 pattern)",
    "ap_cfa_distributions": "Account-string coding lookup GL code -> worktags (Supplier Inv archetype C, MOST COMMON)",
    "ap_prize_funds": "Prize-fund specific, forks from fixed-constants archetype A",
    "ap_referees": "Referees supplier list, forks from archetype A",
    "ar_discipline_invoice": "Customer Invoice fixed-constants archetype",
    "ar_invoice": "Customer Invoice with header row at row 6 (varied source start)",
    "ar_invoice_run": "Customer Invoice - WNL/run archetype",
}

EXEMPLARY_PROFILE_KEYS = tuple(ARCHETYPES.keys())

EIB_ABBREVIATION: dict[str, str] = {
    # Display name -> suffix used in VBA module name + Run macro.
    # MUST match the pattern used in existing 9 transforms so grep-by-module
    # name conventions continue to hold after scaffolding.
    "Supplier Invoice": "SupplierInvoice",
    "Customer Invoice": "CustomerInvoice",
}

ALLOWED_EIB_TYPES = tuple(EIB_ABBREVIATION.keys())


@dataclass(frozen=True)
class RenamePlan:
    """All names a scaffold operation needs to rewrite, with their computed
    post-rename values. Enforces 31-char limit before any file is written."""

    # Source (fork-from) values:
    from_key: str
    from_label: str
    from_eib_type: str
    from_bas_stem: str
    from_step_macro: str

    # New (to-be-written) values:
    new_key: str
    new_label: str
    new_eib_type: str
    new_bas_stem: str
    new_step_macro: str


def create_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Scaffold a new elt_transform_excel transform by forking the "
        "closest structural exemplar and applying rename everywhere. Every location "
        "that requires human/agent judgement afterwards is left as a TO_FILL_* token."
    )
    parser.add_argument(
        "--fork-from",
        choices=EXEMPLARY_PROFILE_KEYS,
        help="Profile key of the exemplar to fork from. One of the 9 currently "
        "shipped profiles. Use --list to see them with archetype labels.",
    )
    parser.add_argument(
        "--new-key",
        help="New snake_case profile key, e.g. ap_bacardi_uk. Becomes the YAML "
        "filename stem and the transform key in registration.",
    )
    parser.add_argument(
        "--new-label",
        help="New transform label shown in the Runner dropdown, e.g. "
        "'AP Bacardi UK'. Keep < ~50 chars.",
    )
    parser.add_argument(
        "--eib-type",
        choices=ALLOWED_EIB_TYPES,
        help="Workday EIB type. Determines the VBA module-name suffix and the "
        "default eib_population_type value written into the profile YAML.",
    )
    parser.add_argument(
        "--source-type",
        help="Source Type value shown in registry/config. A short label for the "
        "source format, e.g. 'Bacardi UK GL Raw'.",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="List exemplar profile keys with their archetype and exit.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing files if the target artifacts already exist. "
        "Without --force the script errors out to avoid clobbering real work.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Compute rename plan, validate 31-char limits, print what would be "
        "written, but touch NO files.",
    )
    return parser


def _import_yaml():
    try:
        import yaml
    except ImportError as exc:  # pragma: no cover
        raise SystemExit(
            "PyYAML is required. Run with:\n"
            "uv run --with pyyaml --no-project python "
            "elt_transform_excel/scripts/scaffold_transform.py"
        ) from exc
    return yaml


def _camel_case(snake_or_spaces: str) -> str:
    """Convert 'ap_bacardi_uk' or 'AP Bacardi UK' -> 'ApBacardiUk'."""
    words = re.split(r"[\s_]+", snake_or_spaces.strip().lower())
    return "".join(w.capitalize() for w in words if w)


def _build_bas_stem(new_label: str, eib_type: str) -> str:
    """modApBacardiUkSupplierInvoice (31 chars max - build gates this)."""
    return f"mod{_camel_case(new_label)}{EIB_ABBREVIATION[eib_type]}"


def _build_step_macro(new_label: str, eib_type: str) -> str:
    """RunApBacardiUkSupplierInvoice."""
    return f"Run{_camel_case(new_label)}{EIB_ABBREVIATION[eib_type]}"


def _load_profile(yaml_mod, key: str) -> dict:
    path = TRANSFORMS_DIR / f"{key}.yaml"
    if not path.exists():
        raise SystemExit(f"Fork-from profile not found: {path}")
    with path.open() as f:
        data = yaml_mod.safe_load(f) or {}
    if not isinstance(data, dict):
        raise SystemExit(f"Fork-from profile is not a mapping: {path}")
    return data


def _resolve_existing_bas_stem(yaml_mod, from_key: str) -> str:
    profile = _load_profile(yaml_mod, from_key)
    step_macro = (
        profile.get("runtime", {}).get("step_list_macro", "") if isinstance(profile.get("runtime"), dict) else ""
    )
    if not step_macro:
        raise SystemExit(
            f"Fork-from profile {from_key!r} has no runtime.step_list_macro value "
            f"to locate its VBA module by."
        )
    # Scan vba_src/transforms for a .bas module whose source declares the
    # Public Sub named step_macro. The module filenames sometimes use
    # abbreviations (Distrib<...> not Distributions<...>) so the filename
    # stem isn't derivable purely from the Run* macro name; the declaration
    # text inside the file is the authority.
    import re as _re

    decl_re = _re.compile(
        rb"^\s*Public\s+Sub\s+" + _re.escape(step_macro.encode("ascii")) + rb"\s*\(",
        _re.MULTILINE,
    )
    matches: list[Path] = []
    for bas in VBA_TRANSFORMS_DIR.glob("mod*.bas"):
        raw = bas.read_bytes()
        if decl_re.search(raw):
            matches.append(bas)
    if not matches:
        # Fallback: try the shared parent dir (modDispatcher / modHelpers etc
        # are in vba_src/ not vba_src/transforms/).
        for bas in (VBA_TRANSFORMS_DIR.parent).glob("**/mod*.bas"):
            raw = bas.read_bytes()
            if decl_re.search(raw):
                matches.append(bas)
    if len(matches) == 1:
        return matches[0].stem
    if not matches:
        raise SystemExit(
            f"No VBA module under {VBA_TRANSFORMS_DIR} declares "
            f"Public Sub {step_macro}. Cannot locate fork-from module for "
            f"key {from_key!r}."
        )
    # Multiple matches: pick the transforms/ subdir one if available, else error.
    in_transforms = [p for p in matches if p.parent == VBA_TRANSFORMS_DIR]
    if len(in_transforms) == 1:
        return in_transforms[0].stem
    raise SystemExit(
        f"Ambiguous: {len(matches)} modules declare Public Sub {step_macro}: "
        f"{[str(p) for p in matches]}. Delete the stale duplicate."
    )


def _build_rename_plan(
    yaml_mod,
    fork_from: str,
    new_key: str,
    new_label: str,
    eib_type: str,
) -> RenamePlan:
    if not re.fullmatch(r"[a-z][a-z0-9_]{2,63}", new_key):
        raise SystemExit(
            f"--new-key {new_key!r} must be snake_case, 3-63 chars, start with a-z."
        )

    from_profile = _load_profile(yaml_mod, fork_from)
    transform_meta = from_profile.get("transform") or {}
    from_label = str(transform_meta.get("label") or fork_from)
    from_eib_type = str(transform_meta.get("eib_population_type") or "")
    if from_eib_type not in EIB_ABBREVIATION:
        from_eib_type = "Supplier Invoice"
    from_bas_stem = _resolve_existing_bas_stem(yaml_mod, fork_from)
    from_step_macro = (
        from_profile.get("runtime", {}).get("step_list_macro", "")
        if isinstance(from_profile.get("runtime"), dict)
        else ""
    )
    if not from_step_macro:
        from_step_macro = f"Run{from_bas_stem[3:]}"

    new_bas_stem = _build_bas_stem(new_label, eib_type)
    new_step_macro = _build_step_macro(new_label, eib_type)

    # HARD GATES BEFORE WRITING ANYTHING.
    violations: list[str] = []
    if len(new_bas_stem) > MAX_VBA_COMPONENT_NAME_CHARS:
        violations.append(
            f"Computed VBA module stem {new_bas_stem!r} is "
            f"{len(new_bas_stem)} chars > {MAX_VBA_COMPONENT_NAME_CHARS}. "
            f"Shorten --new-label (currently {new_label!r}) or fall back to a "
            f"manually chosen shorter abbreviation."
        )
    # VB_Name regex: no punctuation, ASCII letters/numbers only.
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,30}", new_bas_stem):
        violations.append(
            f"VBA module stem {new_bas_stem!r} is not a valid VBComponent name. "
            f"--new-label must only contain letters/numbers/spaces/underscores."
        )
    # step_macro matches Sub naming rules
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,254}", new_step_macro):
        violations.append(
            f"Computed step_list_macro {new_step_macro!r} is not a valid VBA Sub name."
        )
    # 31-char gate: "Attribute VB_Name = \"{stem}\"" is what VBA actually loads by
    if len(new_bas_stem) > MAX_VBA_COMPONENT_NAME_CHARS:
        pass  # already reported above
    if violations:
        raise SystemExit("Scaffold plan failed validation:\n  - " + "\n  - ".join(violations))

    return RenamePlan(
        from_key=fork_from,
        from_label=from_label,
        from_eib_type=from_eib_type,
        from_bas_stem=from_bas_stem,
        from_step_macro=from_step_macro,
        new_key=new_key,
        new_label=new_label,
        new_eib_type=eib_type,
        new_bas_stem=new_bas_stem,
        new_step_macro=new_step_macro,
    )


def _rewrite_text(text: str, plan: RenamePlan) -> str:
    """Apply string-level rewrites in a stable order (longest match first to
    avoid partial overlaps like modApAmexTrains being a prefix of
    modApAmexTrainsV2SupplierInvoice)."""

    pairs = [
        (plan.from_step_macro, plan.new_step_macro),
        (plan.from_bas_stem, plan.new_bas_stem),
        # Label text inside comments/descriptions: case-insensitive word swap,
        # but prefer exact-case first.
        (plan.from_label, plan.new_label),
        (plan.from_key, plan.new_key),
    ]
    # Longest-match-first ordering
    pairs.sort(key=lambda p: len(p[0]), reverse=True)

    for old, new in pairs:
        if old and old in text:
            text = text.replace(old, new)

    return text


def _scaffold_profile_yaml(yaml_mod, plan: RenamePlan, source_type: str, force: bool, dry_run: bool) -> Path:
    out = TRANSFORMS_DIR / f"{plan.new_key}.yaml"
    if out.exists() and not force:
        raise SystemExit(
            f"Target profile already exists: {out}. Pass --force to overwrite."
        )
    from_path = TRANSFORMS_DIR / f"{plan.from_key}.yaml"
    raw_text = from_path.read_text(encoding="utf-8")
    rewritten = _rewrite_text(raw_text, plan)

    # Now apply TO_FILL_* markers to the fields that can't be inferred from the
    # exemplar fork alone. These are the 2 judgement steps: the
    # source_layout.required_headers list, shell_defaults.* paths, and (if the
    # archetype requires it) any raw-derivation pre-processing description.
    lines = rewritten.splitlines()
    out_lines: list[str] = []
    in_shell_defaults = False
    in_source_layout = False
    indent = "  "
    for line in lines:
        stripped = line.lstrip()
        if stripped.startswith("transform:"):
            out_lines.append(line)
            continue
        if stripped.startswith("label:"):
            key_spaces = len(line) - len(stripped)
            out_lines.append(" " * key_spaces + f"label: {plan.new_label}")
            continue
        if stripped.startswith("key:"):
            key_spaces = len(line) - len(stripped)
            out_lines.append(" " * key_spaces + f"key: {plan.new_key}")
            continue
        if stripped.startswith("source_type:"):
            key_spaces = len(line) - len(stripped)
            val = source_type if source_type else "TO_FILL_SOURCE_TYPE"
            out_lines.append(" " * key_spaces + f"source_type: {val}")
            continue
        if stripped.startswith("eib_population_type:"):
            key_spaces = len(line) - len(stripped)
            out_lines.append(" " * key_spaces + f"eib_population_type: {plan.new_eib_type}")
            continue
        if stripped.startswith("step_list_macro:"):
            key_spaces = len(line) - len(stripped)
            out_lines.append(" " * key_spaces + f"step_list_macro: {plan.new_step_macro}")
            continue
        if stripped.startswith("shell_defaults:"):
            in_shell_defaults = True
            in_source_layout = False
            out_lines.append(line)
            continue
        if stripped.startswith("source_layout:"):
            in_source_layout = True
            in_shell_defaults = False
            out_lines.append(line)
            continue
        if in_shell_defaults and stripped and not line.startswith(" " * 2):
            in_shell_defaults = False
        if in_source_layout and stripped and not line.startswith(" " * 2):
            in_source_layout = False

        if in_shell_defaults and ":" in stripped:
            # Paths are machine-local. Replace them wholesale with sentinels so
            # the agent doesn't accidentally commit a forked path to a
            # different user's home directory.
            field, _, _old_val = stripped.partition(":")
            key_spaces = len(line) - len(stripped)
            sentinel = f"TO_FILL_{field.strip().upper()}_PATH"
            out_lines.append(" " * key_spaces + f"{field}: {sentinel}")
            continue
        if in_source_layout and stripped.startswith("required_headers:"):
            key_spaces = len(line) - len(stripped)
            out_lines.append(" " * key_spaces + stripped.split(":")[0] + ":")
            out_lines.append(
                " " * (key_spaces + 2) + "# TO_FILL_REQUIRED_HEADERS: list from actual source file headers"
            )
            out_lines.append(" " * (key_spaces + 2) + "- TO_FILL_HEADER_1")
            out_lines.append(" " * (key_spaces + 2) + "- TO_FILL_HEADER_2")
            # Skip existing required_headers list items below until next
            # top-level key.
            continue
        if in_source_layout and stripped.startswith("- ") and out_lines and "TO_FILL_REQUIRED_HEADERS" in str(out_lines[-3:-1]):
            # Skip trailing original list items we didn't emit above.
            # Detection heuristic: we just wrote 3 sentinel lines, so skip the
            # following `- ...` lines until non-list.
            continue
        out_lines.append(line)

    # Ensure trailing newline; use LF (YAML files are not .bas)
    final = "\n".join(out_lines).rstrip() + "\n"
    if dry_run:
        print(f"[dry-run] would write profile: {out} ({len(final)} bytes)")
    else:
        out.write_text(final, encoding="utf-8")
        print(f"Wrote profile: {out}")
    return out


def _scaffold_transform_spec_yaml(yaml_mod, plan: RenamePlan, force: bool, dry_run: bool) -> Path:
    from_path = TRANSFORM_SPECS_DIR / f"{plan.from_key}.yaml"
    out = TRANSFORM_SPECS_DIR / f"{plan.new_key}.yaml"
    if not from_path.exists():
        # Not every exemplar ships a transform_spec. Leave a minimal stub
        # pointing at the canonical inference playbook.
        stub = (
            f"# Column-mapping spec for {plan.new_label} (EIB: {plan.new_eib_type}).\n"
            f"#\n"
            f"# This is an agent/human-authored companion to:\n"
            f"#   ../transforms/{plan.new_key}.yaml\n"
            f"#\n"
            f"# Expected shape mirrors the 9 existing transform_specs/*.yaml files:\n"
            f"#   version: 1\n"
            f"#   columns:\n"
            f"#     - source: '<raw source header>'\n"
            f"#       target: '<EIB target column>'\n"
            f"#       mapping: 'passthrough | format_date_ddmmyyyy | lookup_company | ...'\n"
            f"#\n"
            f"# TO_FILL_MAPPING_SPEC: fill in every source -> target row. See\n"
            f"# docs/todo/archive/cold-start-scale.md § II. Column Classification for the\n"
            f"# 8 allowed classes (Pass-through / Format-translate / Per-row lookup /\n"
            f"# Account-string lookup / Raw derivation / Fixed constant / Cross-sheet\n"
            f"# source lookup / Workday IDs).\n"
            "version: 1\n"
            "\n"
            "columns:\n"
            "  - source: TO_FILL_SOURCE_COLUMN_1\n"
            "    target: TO_FILL_EIB_TARGET_COLUMN\n"
            "    mapping: TO_FILL_ONE_OF_THE_8_CLASSES\n"
        )
        final = stub
    else:
        raw = from_path.read_text(encoding="utf-8")
        rewritten = _rewrite_text(raw, plan)
        final = (
            "# TO_FILL_MAPPING_SPEC: This YAML is a structural copy of the\n"
            f"# {plan.from_key} exemplar. EVERY row below must be re-reviewed to\n"
            "# match the new source file headers and the new EIB target layout.\n"
            "# See docs/todo/archive/cold-start-scale.md § II. Column Classification.\n"
            + rewritten
        )
    if out.exists() and not force:
        raise SystemExit(
            f"Target transform_spec already exists: {out}. Pass --force to overwrite."
        )
    if dry_run:
        print(f"[dry-run] would write transform_spec: {out}")
    else:
        out.write_text(final, encoding="utf-8")
        print(f"Wrote transform_spec: {out}")
    return out


def _scaffold_bas(plan: RenamePlan, force: bool, dry_run: bool) -> Path:
    from_path = VBA_TRANSFORMS_DIR / f"{plan.from_bas_stem}.bas"
    if not from_path.exists():
        raise SystemExit(
            f"Cannot locate exemplar VBA module at {from_path}. Check fork-from "
            f"key -> bas_stem resolution. Known modules: "
            f"{sorted(p.stem for p in VBA_TRANSFORMS_DIR.glob('mod*.bas'))}"
        )
    out = VBA_TRANSFORMS_DIR / f"{plan.new_bas_stem}.bas"
    if out.exists() and not force:
        raise SystemExit(
            f"Target VBA module already exists: {out}. Pass --force to overwrite."
        )
    raw = from_path.read_bytes()
    # Exemplar files are always CRLF (build_runner.py hygiene gate). Rewrite as
    # text then re-encode to CRLF ASCII.
    try:
        text = raw.decode("ascii")
    except UnicodeDecodeError:
        # Legacy exemplars carry a handful (<20 bytes) of tolerated non-ASCII
        # characters like em dashes. Replace them with ASCII equivalents so
        # the scaffolded output is always pure-ASCII (the 31-char gate is
        # already enforced, and build_runner.py errors on >20 non-ASCII
        # bytes in any .bas file, so scaffolding must not accidentally ship
        # the exemplar's tolerance band).
        text = raw.decode("utf-8", errors="replace")
        # Common visual-character -> ASCII replacements.
        _replacements = (
            ("\u2014", "--"),   # em dash
            ("\u2013", "-"),    # en dash
            ("\u2018", "'"),    # left single quote
            ("\u2019", "'"),    # right single quote
            ("\u201c", '"'),    # left double quote
            ("\u201d", '"'),    # right double quote
            ("\u2026", "..."),  # ellipsis
            ("\ufffd", "?"),    # replacement char
        )
        for old, new in _replacements:
            text = text.replace(old, new)
    text_lf = text.replace("\r\n", "\n").replace("\r", "\n")

    rewritten = _rewrite_text(text_lf, plan)

    # Mark the two big judgement sections with TO_FILL sentinels so the user can
    # find them: (1) the required-headers constant array, (2) the per-row
    # column-writes inside whatever the exemplar's main "write one output row"
    # Sub is called.
    insertions: list[tuple[int, str]] = []
    lower = rewritten.lower()
    required_markers = [
        "requiredheaders",
        "required_headers",
        "required headers",
    ]
    # Just add a banner comment as the very first line after Option Explicit /
    # Attribute VB_Name. We don't attempt to surgically edit the VBA - markers
    # are grepable, which is sufficient for H2.1.
    banner_lines = (
        "' ============================================================\n"
        f"' SCAFFOLDED TRANSFORM: {plan.new_label} (EIB: {plan.new_eib_type})\n"
        f"' Forked from exemplar: {plan.from_key} ({plan.from_label})\n"
        "'\n"
        "' TO_FILL items - grep for them, then edit each location marked.\n"
        "'   TO_FILL_REQUIRED_HEADERS  -> Const REQUIRED_HEADERS array\n"
        "'   TO_FILL_OUTPUT_COLUMNS    -> Cells(row, col) = source(...) writes\n"
        "'   TO_FILL_MAPPING_CALLS     -> LookupMappingValue(...) entity/key args\n"
        "'   TO_FILL_SOURCE_HEADER_ROW -> If a source file has title rows before\n"
        "'                                 the real header row, adjust FindHeaderRow\n"
        "'                                 call or Const HEADER_ROW here.\n"
        "'\n"
        "' Archetype advice (see docs/todo/archive/cold-start-scale.md § I):\n"
        f"'   Exemplar archetype -> {ARCHETYPES.get(plan.from_key, 'unknown')}\n"
        "' ============================================================\n"
    )
    # Insert banner after Attribute VB_Name line (line 1) + Option Explicit (line 2).
    first_blank_or_option = rewritten.find("\n\n")
    if first_blank_or_option == -1:
        first_blank_or_option = rewritten.find("\nOption Explicit")
    if first_blank_or_option != -1:
        insert_point = rewritten.find("\n", first_blank_or_option + 1) + 1
        rewritten = rewritten[:insert_point] + banner_lines + rewritten[insert_point:]

    # Replace ASCII-only. Encode-decode pass forces any remaining non-ASCII char
    # to '?' so the scaffold never builds a file that fails the >20-byte
    # non-ASCII build gate.
    bytes_out = rewritten.encode("ascii", errors="replace")
    # CRLF everywhere (Mac Excel VBA import gate in build_runner.py).
    bytes_crlf = bytes_out.replace(b"\n", b"\r\n")
    if dry_run:
        print(
            f"[dry-run] would write VBA module: {out} "
            f"(module stem len={len(plan.new_bas_stem)} <= 31, {len(bytes_crlf)} bytes)"
        )
    else:
        out.write_bytes(bytes_crlf)
        print(f"Wrote VBA module: {out}")
    return out


def main(argv: list[str] | None = None) -> int:
    parser = create_parser()
    args = parser.parse_args(argv)

    if args.list:
        print("Known exemplar profile keys (--fork-from choices):\n")
        for key in EXEMPLARY_PROFILE_KEYS:
            label = ARCHETYPES[key]
            print(f"  {key:<24s} -> {label}")
        print()
        print(
            "Pick archetype guidance (docs/todo/archive/cold-start-scale.md § I):\n"
            "  - AP Supplier Invoice, per-row mapping_master rows for CC/Worktags ->\n"
            "    archetype C (ap_cfa_distributions) - MOST COMMON.\n"
            "  - AP Supplier Invoice, NO account-string coding, just fixed values ->\n"
            "    archetype A (ap_amex_hotel).\n"
            "  - AP Supplier Invoice that needs raw-derivation pre-step BEFORE\n"
            "    mapping (e.g. Ref 2 string-split to produce Cost Centre/VAT) ->\n"
            "    archetype ap_amex_trains_v2 pattern.\n"
            "  - Customer Invoice -> ar_invoice or ar_discipline_invoice.\n"
        )
        return 0

    missing = [
        f"--{name.replace('_', '-')}"
        for name, val in (
            ("fork_from", args.fork_from),
            ("new_key", args.new_key),
            ("new_label", args.new_label),
            ("eib_type", args.eib_type),
        )
        if not val
    ]
    if missing:
        parser.error("the following arguments are required: " + ", ".join(missing))
        return 2  # argparse exits 2; this line unreachable but typecheck likes it

    yaml_mod = _import_yaml()

    plan = _build_rename_plan(
        yaml_mod,
        args.fork_from,
        args.new_key,
        args.new_label,
        args.eib_type,
    )
    print("=== Rename plan ===")
    print(f"  fork-from key:          {plan.from_key}  ({plan.from_label})")
    print(f"  fork-from module:       {plan.from_bas_stem}  (Run macro: {plan.from_step_macro})")
    print(f"  new key:                {plan.new_key}  ({plan.new_label})")
    print(f"  new eib_population_type:{plan.new_eib_type}")
    print(f"  new module stem:        {plan.new_bas_stem}  (len={len(plan.new_bas_stem)} <= 31: OK)")
    print(f"  new step_list_macro:    {plan.new_step_macro}")
    print()

    _scaffold_profile_yaml(
        yaml_mod, plan, args.source_type or "", force=args.force, dry_run=args.dry_run
    )
    _scaffold_transform_spec_yaml(yaml_mod, plan, force=args.force, dry_run=args.dry_run)
    _scaffold_bas(plan, force=args.force, dry_run=args.dry_run)
    print()
    # ----- Conventions loaded directly from configs (NOT inferred) -----
    from_profile = _load_profile(yaml_mod, plan.from_key)
    eib_defaults_path = TRANSFORMS_DIR.parent / "eib_population_defaults.yaml"
    eib_defaults_doc: dict[str, dict[str, object]] = {}
    if eib_defaults_path.exists():
        with eib_defaults_path.open() as handle:
            raw_eib = yaml_mod.safe_load(handle) or {}
        eib_defaults_doc = (
            raw_eib.get("eib_populations")
            if isinstance(raw_eib.get("eib_populations"), dict)
            else {}
        )
    eib_entry = eib_defaults_doc.get(plan.new_eib_type, {})
    print("=== Target-EIB conventions (sourced from eib_population_defaults.yaml, not inferred) ===")
    if not eib_entry:
        print(f"  WARNING: {plan.new_eib_type!r} not listed in eib_population_defaults.yaml.")
        print("  Either add a new eib_populations entry, or the build_runner.py --validate step")
        print("  will fail for this transform BEFORE writing any workbook.")
    else:
        sheet = eib_entry.get("target_sheet_name", "<missing>")
        start_row = eib_entry.get("target_data_start_row", "<missing>")
        end_col = eib_entry.get("clear_published_rows_end_col", "<missing>")
        print(f"  target_eib_population_type: {plan.new_eib_type}")
        print(f"  target_sheet_name:          {sheet!r}")
        print(f"  target_data_start_row:      {start_row}")
        print(f"  clear_published_rows_end:   col {end_col!r}")
        print()
    src_layout = from_profile.get("source_layout") if isinstance(from_profile.get("source_layout"), dict) else {}
    resolution = src_layout.get("worksheet_resolution", "<missing>")
    detection = src_layout.get("header_detection", "<missing>")
    ws_name = src_layout.get("worksheet_name", "<first/only sheet>")
    hdr_row = src_layout.get("header_row", "<missing>")
    data_row = src_layout.get("data_start_row", "<missing>")
    print("=== Source-layout conventions (inherited from exemplar, NOT guessed) ===")
    print(f"  worksheet_resolution:  {resolution}  -> ws={ws_name!r}")
    print(f"  header_detection:      {detection}  -> header_row={hdr_row}, data_start_row={data_row}")
    print()
    print(
        "=== Post-scaffold TO_FILL checklist ==="
        + "\n  1. Edit "
        + str(TRANSFORMS_DIR / f"{plan.new_key}.yaml")
        + "\n     -> Replace every TO_FILL_*_PATH with the actual Finance-local source,"
        + "\n        EIB template, output, mapping, customer/supplier spreadsheet paths."
        + "\n     -> Replace TO_FILL_SOURCE_TYPE if not already set via --source-type."
        + "\n     -> Under source_layout.required_headers, list every header the"
        + "\n        actual source file uses (FindHeaderRow validates these at runtime)."
        + "\n  2. Edit "
        + str(TRANSFORM_SPECS_DIR / f"{plan.new_key}.yaml")
        + "\n     -> One row per source->EIB column mapping; pick one of the 8 classes"
        + "\n        from docs/todo/archive/cold-start-scale.md § II."
        + "\n  3. Edit "
        + str(VBA_TRANSFORMS_DIR / f"{plan.new_bas_stem}.bas")
        + "\n     -> Const REQUIRED_HEADERS array to match source file."
        + "\n     -> Per-output-row Cells(row, col) writes: replace TO_FILL_OUTPUT_COLUMNS."
        + "\n     -> If any per-row values need new ref_*.json rows (supplier names,"
        + "\n        GL codes, etc.) create them in config/ingest_lookups/ FIRST - the"
        + "\n        per-row generalization rule says: NEVER hardcode a per-row varying"
        + "\n        value in VBA; always ref_*.json first (transform-forever constants"
        + "\n        like the AME001 supplier ID are the only exception)."
        + "\n  4. Build + verify:"
        + "\n        uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project \\"
        + "\n            python implementations/elt_transform_excel/scripts/build_runner.py"
        + "\n        # (then run BootstrapImport macro in Excel and --verify)"
        + "\n  5. Inspect registry sheet state:"
        + "\n        uv run --with openpyxl --no-project python \\"
        + "\n            implementations/elt_transform_excel/scripts/inspect_registry.py --strict"
        + "\n"
        + "\n=== 6. Non-configurable tenant / user prompts (REQUIRED confirmation) ==="
        + "\n   These are the 6 judgement items that cannot be derived from config."
        + "\n   Confirm each one with Finance/user before the macro is considered done:"
        + "\n"
        + "\n   1. SOURCE WORKSHEET resolution:"
        + "\n      Inherited: worksheet_resolution=" + str(resolution) + " (exemplar ws=" + repr(str(ws_name)) + ")."
        + "\n      Confirm: is the real source tab name identical / really the first sheet?"
        + "\n"
        + "\n   2. COMPANY code + mapping-tab fallback text:"
        + "\n      Hardcoded tenant company (e.g. FA / WNSL) plus the 'THE FA GROUP - FA'"
        + "\n      style fallback used by ResolveMappedCompany singleton. MUST be set per"
        + "\n      transform; never guess a Workday company ref-id."
        + "\n"
        + "\n   3. FIXED tenant Workday ref-IDs that are NOT derived from source columns:"
        + "\n      List every hardcoded constant target-column ref-id you plan to use:"
        + "\n      (examples: Tax Applicability GBR_Outside_of_Scope, Tax Code"
        + "\n      GBR_Zero_Rated / GBR_200_VAT, Tax Recoverability Nonrecoverable_Allocate,"
        + "\n      Spend Category Hotel_Staff_Costs, Currency GBP, Company FA / WNSL,"
        + "\n      Document Type, Submit Y/N, Worktag IDs, etc.)"
        + "\n      Rule: Class 8 (tenant workday IDs) = NEVER guess; get Finance sign-off."
        + "\n"
        + "\n   4. UNUSED source columns — EXPLICIT CONFIRMATION:"
        + "\n      After classifying all source headers into the 8 mapping classes, list"
        + "\n      any header NOT yet assigned (e.g. 'Cost Code', 'FA WNL Code' not used)."
        + "\n      Confirm: 'These are intentionally NOT mapped — correct?' Otherwise a"
        + "\n      naive agent would mis-assign them to Cost Centre or similar and produce"
        + "\n      a silent wrong-coding EIB file at Workday-import time."
        + "\n"
        + "\n   5. REQUIRED_HEADERS list confirmation (header_detection=scan_required_headers):"
        + "\n      Replace TO_FILL_HEADER_* sentinels with the real headers from the source"
        + "\n      file. Confirm: 'Does every header below exist verbatim in the source?"
        + "\n      FindHeaderRow will error hard if any single header is missing.'"
        + "\n"
        + "\n   6. TAX mode per transform:"
        + "\n      Confirm: uniform (one tax code for all rows, e.g. Outside of Scope zero-AP),"
        + "\n      or per_row (driver column in source with explicit when_value -> code sets)."
        + "\n      If per_row, confirm the driver source column + every when_value + every"
        + "\n      target tax-code ref-id."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
