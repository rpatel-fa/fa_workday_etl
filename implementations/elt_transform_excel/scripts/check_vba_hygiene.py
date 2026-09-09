#!/usr/bin/env python3
"""Standalone VBA source hygiene pre-flight check.

Zero external deps (stdlib only, < 0.2s on 13 modules). Run BEFORE any Runner
build_runner.py, create_runner_shell.py, or any VBA edit session.

Same rules as build_runner.py _validate_vba_source_hygiene(), plus a --fix mode
that automatically rewrites common unicode to ASCII equivalents in place.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve()
MODULE_ROOT = SCRIPT_PATH.parents[1]
DEFAULT_VBA_SOURCE = MODULE_ROOT / "vba_src"
DEFAULT_PROFILE_DIR = MODULE_ROOT / "config" / "transforms"

MAX_VBA_COMPONENT_NAME_CHARS = 31
WARN_MAX_SOURCE_LINE_CHARS = 250
ERROR_MAX_SOURCE_LINE_CHARS = 299  # BUILD.md rule 4: Mac parser observed to fail on 258+ char lines; treat >= 300 as hard ERROR
WARN_LINE_CONTINUATION_RUN = 15     # >= 15 sequential line-continuations raises WARN
ERROR_LINE_CONTINUATION_RUN = 22   # >= 22 sequential line-continuations raises ERROR (hard Mac VBA parser continuation-stack failure observed at run=25; Stationery/Taxi bombed there)
NON_ASCII_ERROR_THRESHOLD = 20
NON_ASCII_WARN_THRESHOLD = 1
# For xattrs: only ERROR-class on the one xattr *proven* to cause generic
# 80040714 Import() failures on macOS: com.apple.quarantine.
# com.apple.provenance / com.apple.metadata:kMDItemWhereFroms are WARN only:
# they tag provenance/metadata but historically do NOT by themselves prevent
# VBComponents.Import() from succeeding.  build_runner.py pre-hygiene step
# strips them all anyway as belt-and-braces so they are zero-touch in normal
# workflow.
MACOS_XATTR_BLOCKLIST_ERROR = ("com.apple.quarantine",)
MACOS_XATTR_BLOCKLIST_WARN = (
    "com.apple.provenance",
    "com.apple.metadata:kMDItemWhereFroms",
)

# --fix rewrite rules. Order matters (left-applied.
UNICODE_TO_ASCII = [
    ("\u2014", "--"),  # em dash
    ("\u2013", "-"),   # en dash
    ("\u2010", "-"),   # hyphen
    ("\u2212", "-"),  # minus sign
    ("\u2192", "->"),  # rightwards arrow
    ("\u2190", "<-"),  # leftwards arrow
    ("\u21d2", "=>"),  # rightwards double arrow
    ("\u21d0", "<="),  # leftwards double arrow
    ("\u2018", "'"),   # left single quote
    ("\u2019", "'"),   # right single quote
    ("\u201c", '"'),   # left double quote
    ("\u201d", '"'),   # right double quote
    ("\u2026", "..."),  # ellipsis
    ("\u00a0", " "),   # non-breaking space
    ("\u2500", "-"),   # box drawings light horizontal
    ("\u2501", "-"),   # box drawings heavy horizontal
    ("\u2502", "|"),   # box drawings light vertical
    ("\u2503", "|"),   # box drawings heavy vertical
    ("\u251c", "+"),   # box drawings light vertical and right
    ("\u252c", "+"),   # box drawings light down and horizontal
    ("\u253c", "+"),   # box drawings light vertical and horizontal
    ("\u2514", "+"),   # box drawings light up and right
    ("\u2518", "+"),   # box drawings light up and left
    ("\u2510", "+"),   # box drawings light down and left
    ("\u00b7", "*"),   # middle dot
    ("\u2022", "*"),   # bullet
    ("\u2713", "v"),   # check mark
    ("\u2717", "x"),   # ballot cross mark
]

@dataclass
class Violation:
    level: str  # "WARN" | "ERROR"
    module: str
    rule: str
    detail: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "VBA source hygiene pre-flight: fail fast before any Runner build. "
            "Can auto-fix common unicode issues with --fix."
        )
    )
    parser.add_argument(
        "--vba-source", default=str(DEFAULT_VBA_SOURCE),
        help=f"Path to VBA source folder (default: {DEFAULT_VBA_SOURCE})",
    )
    parser.add_argument(
        "--profile-dir", default=str(DEFAULT_PROFILE_DIR),
        help=f"Transform YAML folder for contract checks (default: {DEFAULT_PROFILE_DIR})",
    )
    parser.add_argument(
        "--fix", action="store_true",
        help="Automatically rewrite common non-ASCII bytes (em dash -> --, arrow -> ->, curly quotes -> straight, etc.) then re-scan. Writes files in place.",
    )
    parser.add_argument(
        "--no-contract", dest="contract", action="store_false",
        help="Skip the YAML-contract checks; only run bytes/lines/names/CRLF.",
    )
    parser.add_argument(
        "-q", "--quiet", action="store_true",
        help="Only print violations + summary; suppress per-module pass lines.",
    )
    return parser.parse_args()


def _discover_modules(vba_source: Path) -> list[tuple[str, Path]]:
    """Return [(module_stem, file_path] for all importable modules, excluding BootstrapImporter itself."""
    modules: list[tuple[str, Path]] = []
    seen: set[str] = set()
    for candidate in sorted(vba_source.rglob("*")):
        if not candidate.is_file() or candidate.suffix.lower() not in {".bas", ".cls"}:
            continue
        stem = candidate.stem
        if stem.lower() == "bootstrapimporter":
            continue
        key = stem.lower()
        if key in seen:
            raise RuntimeError(f"Duplicate VBA component name '{stem}' under {vba_source}.")
        seen.add(key)
        modules.append((stem, candidate))
    return modules


def _fix_unicode_in_place(path: Path) -> int:
    raw_before = path.read_bytes()
    try:
        text = raw_before.decode("utf-8")
    except UnicodeDecodeError:
        print(f"  ! {path.name}: not UTF-8, skipping --fix")
        return 0
    changed = False
    for uni, asc in UNICODE_TO_ASCII:
        if uni in text:
            text = text.replace(uni, asc)
            changed = True
    if not changed:
        return 0
    # Preserve CRLF vs LF line endings as-is (only rewrite changed text)
    # Figure original sep
    sep = b"\r\n" if b"\r\n" in raw_before else b"\n"
    # ensure output has same line endings
    norm_lines = text.replace("\r\n", "\n").replace("\r", "\n")
    out_text = norm_lines.replace("\n", "\r\n") if sep == b"\r\n" else norm_lines
    path.write_bytes(out_text.encode("utf-8"))
    count_removed = sum(1 for b in raw_before if b >= 128) - sum(1 for b in path.read_bytes() if b >= 128)
    return max(0, count_removed)


def _check_bytes(module_name: str, path: Path) -> list[Violation]:
    violations: list[Violation] = []
    raw = path.read_bytes()

    # Component name lengths (filename stem + internal Attribute VB_Name)
    stem = path.stem
    if len(stem) > MAX_VBA_COMPONENT_NAME_CHARS:
        violations.append(Violation("ERROR", module_name, "COMPONENT_NAME",
            f"filename stem '{stem}' length={len(stem)} > {MAX_VBA_COMPONENT_NAME_CHARS}"))
    vb_name_re = re.compile(rb'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$', re.MULTILINE)
    attr_match = vb_name_re.search(raw[:4096])
    if attr_match:
        try:
            internal = attr_match.group(1).decode("ascii")
            if len(internal) > MAX_VBA_COMPONENT_NAME_CHARS:
                violations.append(Violation("ERROR", module_name, "COMPONENT_NAME",
                f"Attribute VB_Name '{internal}' length={len(internal)} > {MAX_VBA_COMPONENT_NAME_CHARS}"))
            if internal != stem:
                violations.append(Violation("WARN", module_name, "COMPONENT_NAME_MISMATCH",
                    f"stem '{stem}' != Attribute VB_Name '{internal}'"))
        except UnicodeDecodeError:
            violations.append(Violation("ERROR", module_name, "COMPONENT_NAME",
                "Attribute VB_Name contains non-ASCII bytes"))

    # Non-ASCII count
    non_ascii = sum(1 for b in raw if b >= 128)
    if non_ascii > NON_ASCII_ERROR_THRESHOLD:
        violations.append(Violation("ERROR", module_name, "NON_ASCII",
            f"{non_ascii} non-ASCII bytes (> {NON_ASCII_ERROR_THRESHOLD}). Mac Excel VBA Import fallback to Module1)"))
    elif non_ascii >= NON_ASCII_WARN_THRESHOLD:
        violations.append(Violation("WARN", module_name, "NON_ASCII",
            f"{non_ascii} non-ASCII bytes (<= {NON_ASCII_ERROR_THRESHOLD} tolerated legacy)"))

    # CRLF
    lone_lf = 0
    for i in range(len(raw)):
        if raw[i : i + 1] == b"\x0a" and (i == 0 or raw[i - 1 : i] != b"\x0d"):
            lone_lf += 1
    if lone_lf > 0:
        violations.append(Violation("ERROR", module_name, "LINE_ENDINGS",
            f"{lone_lf} LF-only line ending(s). Mac Excel requires CRLF everywhere."))

    # Overlong lines (WARN: 251..299 chars; ERROR: >=300 chars
    lines = raw.split(b"\r\n") if b"\r\n" in raw else raw.split(b"\x0a")
    long_warn_lines = [i + 1 for i, line in enumerate(lines) if WARN_MAX_SOURCE_LINE_CHARS < len(line) < ERROR_MAX_SOURCE_LINE_CHARS]
    long_err_lines = [i + 1 for i, line in enumerate(lines) if len(line) >= ERROR_MAX_SOURCE_LINE_CHARS]
    if long_err_lines:
        violations.append(Violation("ERROR", module_name, "LONG_LINES",
        f"{len(long_err_lines)} line(s) >= {ERROR_MAX_SOURCE_LINE_CHARS} chars (Mac VBA parser Import observed to fail mid-file on 258+; refactor with line-continuation): {', '.join(str(n) for n in long_err_lines[:5])}{' ...' if len(long_err_lines) > 5 else ''}"))
    if long_warn_lines:
        violations.append(Violation("WARN", module_name, "LONG_LINES",
        f"{len(long_warn_lines)} line(s) 251-299 chars (safe but consider continuation): {', '.join(str(n) for n in long_warn_lines[:5])}{' ...' if len(long_warn_lines) > 5 else ''}"))

    # Consecutive line-continuation runs (Mac VBA parser has a continuation-stack
    # ceiling empirically between 21-25; fail the build on >=15 and warn on 10-14.
    # Without this guardrail, new transforms with 20+ parameter signatures split
    # one-per-line fail the Import() in Excel with the generic 80040714 wall AFTER
    # a build succeeds -- impossible to triage without content-level inspection.)
    cont_runs_err: list[tuple[int, int]] = []  # (start_line, run_len)
    cont_runs_warn: list[tuple[int, int]] = []
    cur_start = -1
    cur_run = 0
    for idx, line in enumerate(lines, 1):
        if line.rstrip().endswith(b"_"):
            if cur_start < 0:
                cur_start = idx
            cur_run += 1
        else:
            if cur_run > 0:
                if cur_run >= ERROR_LINE_CONTINUATION_RUN:
                    cont_runs_err.append((cur_start, cur_run))
                elif cur_run >= WARN_LINE_CONTINUATION_RUN:
                    cont_runs_warn.append((cur_start, cur_run))
            cur_start = -1
            cur_run = 0
    if cur_run > 0:
        if cur_run >= ERROR_LINE_CONTINUATION_RUN:
            cont_runs_err.append((cur_start, cur_run))
        elif cur_run >= WARN_LINE_CONTINUATION_RUN:
            cont_runs_warn.append((cur_start, cur_run))
    if cont_runs_err:
        examples = "; ".join(f"lines {s}-{s+r-1} ({r} in a row)" for s, r in cont_runs_err[:3])
        violations.append(Violation("ERROR", module_name, "LINE_CONTINUATION_RUN",
            f"{len(cont_runs_err)} continuation run(s) >= {ERROR_LINE_CONTINUATION_RUN} sequential '_' markers ({examples}). Mac VBA Import parser continuation-stack crashes at 24+ and can fail as low as 21. Pack params denser into fewer lines (max 2-3 continuations per logical construct) or use a local params UDT / intermediate vars. Last observed Stationery/Taxi 80040714 at run=25."))
    if cont_runs_warn:
        examples = "; ".join(f"lines {s}-{s+r-1} ({r})" for s, r in cont_runs_warn[:3])
        violations.append(Violation("WARN", module_name, "LINE_CONTINUATION_RUN",
            f"{len(cont_runs_warn)} continuation run(s) {WARN_LINE_CONTINUATION_RUN}-{ERROR_LINE_CONTINUATION_RUN - 1} sequential '_' ({examples}). Safe but keep under {ERROR_LINE_CONTINUATION_RUN}."))

    # Index-based sheet access (not inside VBA comment)
    idx_lookup_re = re.compile(rb"\bWorksheets\s*\(\s*(\d+)\s*\)")
    for idx_match in idx_lookup_re.finditer(raw):
        start = idx_match.start()
        line_start = raw.rfind(b"\n", 0, start) + 1
        line_prefix = raw[line_start:start]
        if b"'" in line_prefix:
            continue
        ln = raw[:start].count(b"\n") + 1
        violations.append(Violation("ERROR", module_name, "WORKSHEETS_INDEX",
            f"L{ln}: Worksheets({idx_match.group(1).decode('ascii', errors='replace')}) -- use named constant instead"))

    # macOS extended attribute: com.apple.quarantine
    # When VBA .bas/.cls files are opened/saved through Mac Excel (or exported by
    # the VBE via File → Export), Apple sets the com.apple.quarantine xattr.
    # Subsequent calls to ThisWorkbook.VBProject.VBComponents.Import(<path>)
    # then fail with the opaque runtime error -2147219692 (80040714) and the
    # message "Method 'Import' of object '_VBComponents' failed" -- because
    # macOS Sandbox silently denies Excel read() on quarantined files when
    # called through the VBProject automation bridge (not via the normal File
    # Open dialog).
    #
    # The guardrail here treats the *presence* of com.apple.quarantine on any
    # importable source file as an ERROR: the user must remove it (or run the
    # build-runner, which also does it on every invocation) before attempting
    # BootstrapImportFromConfiguredSource.  We also check for a few related
    # "this file is from the internet" xattrs that trigger the same behaviour.
    XATTR_MACOS_INTERNET_QUARANTINE = "com.apple.quarantine"
    XATTR_MACOS_PROVENANCE = "com.apple.provenance"
    XATTR_MACOS_DOWNLOADED = "com.apple.metadata:kMDItemWhereFroms"
    BLOCKLIST_XATTRS = MACOS_XATTR_BLOCKLIST_ERROR + MACOS_XATTR_BLOCKLIST_WARN
    try:
        import xattr as _xattr_mod  # type: ignore[import-not-found]
        have_xattr_pkg = True
    except Exception:
        have_xattr_pkg = False
    # Fallback: walk the filesystem via `xattr` CLI (always available on macOS)
    error_hits: list[str] = []
    warn_hits: list[str] = []
    if have_xattr_pkg:
        try:
            xattrs_present = list(_xattr_mod.listxattr(str(path)))  # type: ignore[attr-defined]
        except Exception:
            xattrs_present = []
        for name in MACOS_XATTR_BLOCKLIST_ERROR:
            if name in xattrs_present:
                error_hits.append(name)
        for name in MACOS_XATTR_BLOCKLIST_WARN:
            if name in xattrs_present:
                warn_hits.append(name)
    else:
        import subprocess as _sp
        try:
            res = _sp.run(["xattr", str(path)], capture_output=True, text=True, check=False)
            if res.returncode == 0 and res.stdout:
                listed = [ln.strip() for ln in res.stdout.splitlines() if ln.strip()]
                for name in MACOS_XATTR_BLOCKLIST_ERROR:
                    if name in listed:
                        error_hits.append(name)
                for name in MACOS_XATTR_BLOCKLIST_WARN:
                    if name in listed:
                        warn_hits.append(name)
        except FileNotFoundError:
            pass  # not on macOS; can't happen
    if error_hits or warn_hits:
        error_human = ", ".join(error_hits)
        warn_human = ", ".join(warn_hits)
        hint_remove = (
            "Fix (CLI): xattr -d com.apple.quarantine " + str(path)
            + " ; or re-run the build-runner which strips quarantine/provenance/where-from xattrs on every build."
        )
        if XATTR_MACOS_INTERNET_QUARANTINE in error_hits:
            hint_explain = (
                "This xattr is set by macOS whenever a file is saved/exported through "
                "Mac Excel's VBE, downloaded via Safari, or touched by any quarantine-aware app. "
                "When BootstrapImporter calls VBComponents.Import() the macOS Sandbox denies "
                "the read mid-flight and Excel surfaces the generic 80040714 failure. "
                "This is the #1 cause of BootstrapImporter failures on macOS. "
            )
        else:
            hint_explain = "macOS has tagged this file with provenance/where-from metadata. It usually doesn't block Import() outright but build-runner strips it proactively."
        for xa in error_hits:
            violations.append(Violation("ERROR", module_name, "MACOS_XATTR_BLOCKLIST",
                f"{xa} extended attribute present on the source file. {hint_explain}{hint_remove}"))
        for xa in warn_hits:
            violations.append(Violation("WARN", module_name, "MACOS_XATTR_INFO",
                f"{xa} extended attribute present (informational, stripped next build). {hint_explain}{hint_remove}"))
    return violations


def _load_profile_step_list_macros(profile_dir: Path) -> set[str]:
    macro_names: set[str] = set()
    import yaml  # lazy; built-in via uv run --with pyyaml, but stdlib only here -> fail gracefully
    try:
        import yaml as _yaml
    except Exception as e:
        return macro_names
    for p in sorted(Path(profile_dir).glob("*.yaml")):
        try:
            d = _yaml.safe_load(p.read_text())
        except Exception:
            continue
        if isinstance(d, dict):
            step = d.get("runtime", {}).get("step_list_macro") if isinstance(d.get("runtime"), dict) else None
            if isinstance(step, str):
                macro_names.add(step)
    return macro_names


def _check_contract(module_name: str, path: Path, step_list_macros: set[str]) -> list[Violation]:
    viols: list[Violation] = []
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8", errors="replace")
    except Exception:
        return viols
    lines = text.replace("\r\n", "\n").split("\n")

    # ResolveSelectedTransformKey call sites arity
    rstk_call_re = re.compile(r"\bResolveSelectedTransformKey\s*\(([^()]*)\)", re.DOTALL)
    rstk_decl_re = re.compile(r"^\s*(Public|Private|Friend)\s+(Function|Sub|Property)\s+ResolveSelectedTransformKey\b", re.MULTILINE)

    for m in rstk_call_re.finditer(text):
        if rstk_decl_re.search(text[:m.start()]):
            continue  # skip declaration match
        args_str = m.group(1)
        depth = 0
        for ch in args_str:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
        # Count commas where depth == 0
        comma_count = 0
        depth = 0
        for ch in args_str:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "," and depth == 0:
                comma_count += 1
        arg_count = comma_count + 1 if args_str.strip() else 0
        if arg_count != 2:
            # Approximate line
            ln = text[:m.start()].count("\n") + 1
            viols.append(Violation("ERROR", module_name, "RSTK_CALL_ARITY",
                f"L{ln}: ResolveSelectedTransformKey({arg_count} args, expected 2: (runnerWorkbook, controlSheet)"))

    # Public Sub Run<Name> declaration signature must be exactly 2 args typed
    run_sub_re = re.compile(r"^\s*Public\s+Sub\s+(Run[A-Za-z0-9_]+)\s*\(\s*([^)]*)\s*\)\s*$", re.MULTILINE)
    for m in run_sub_re.finditer(text):
        sub_name = m.group(1)
        sig = m.group(2)
        # Skip top-level RunSelectedTransform (zero-arg button-dispatch entry point, not a transform)
        if sub_name == "RunSelectedTransform":
            continue
        # Expect exactly two ByVal ... As Workbook / As Worksheet, order-insensitive
        expected_wb = re.search(r"ByVal\s+\w+\s+As\s+Workbook", sig, re.IGNORECASE)
        expected_ws = re.search(r"ByVal\s+\w+\s+As\s+Worksheet", sig, re.IGNORECASE)
        # Count commas
        nargs = 1 + sum(1 for ch in sig if ch == ",")
        ok = bool(expected_wb and expected_ws and nargs == 2)
        if not ok:
            ln = text[:m.start()].count("\n") + 1
            viols.append(Violation("ERROR", module_name, "RUN_SUB_SIG",
                f"L{ln}: {sub_name} signature not 'Public Sub {sub_name}(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)'"))
    return viols


def main() -> int:
    args = parse_args()
    vba_source = Path(args.vba_source).expanduser().resolve()
    if not vba_source.exists():
        print(f"ERROR: VBA source folder not found: {vba_source}")
        return 2

    t0 = time.monotonic()

    modules = _discover_modules(vba_source)
    if not modules:
        print(f"ERROR: no .bas/.cls found under {vba_source}")
        return 2

    # --- Apply --fix pass first if requested
    fixed_total = 0
    if args.fix:
        for module_name, path in modules:
            fixed_total += _fix_unicode_in_place(path)
        if fixed_total:
            print(f"--fix: rewrote {fixed_total} non-ASCII bytes in-place across all modules.")

    # --- Collect all violations
    all_violations: list[Violation] = []
    per_module_pass: list[str] = []
    for module_name, path in modules:
        viols = _check_bytes(module_name, path)
        if args.contract:
            step_macros = set()
            try:
                import yaml as _y
                for p in Path(args.profile_dir).glob("*.yaml"):
                    try:
                        d = _y.safe_load(p.read_text())
                    except Exception:
                        continue
                    if isinstance(d, dict):
                        s = d.get("runtime", {}).get("step_list_macro") if isinstance(d.get("runtime"), dict) else None
                        if isinstance(s, str):
                            step_macros.add(s)
            except Exception:
                pass
            viols += _check_contract(module_name, path, step_macros)
        if not viols:
            per_module_pass.append(module_name)
        else:
            all_violations.extend(viols)

    # step_list_macro step-list-macro-declared-in-VBA cross-check: all declared YAML step_list_macro names appear as Public Sub names somewhere
    if args.contract:
        step_macros = set()
        try:
            import yaml as _y
            for p in Path(args.profile_dir).glob("*.yaml"):
                try:
                    d = _y.safe_load(p.read_text())
                except Exception:
                    continue
                if isinstance(d, dict):
                    s = d.get("runtime", {}).get("step_list_macro") if isinstance(d.get("runtime"), dict) else None
                    if isinstance(s, str):
                        step_macros.add(s)
        except Exception:
            pass
        all_public_run_subs: set[str] = set()
        run_sub_re = re.compile(r"^\s*Public\s+Sub\s+(Run[A-Za-z0-9_]+)\s*\(", re.MULTILINE)
        for mn, pth in modules:
            try:
                txt = pth.read_text(encoding="utf-8", errors="replace")
            except Exception:
                continue
            for m in run_sub_re.finditer(txt):
                all_public_run_subs.add(m.group(1))
        for macro in sorted(step_macros):
            if macro not in all_public_run_subs and macro != "RunSelectedTransform":
                all_violations.append(Violation("ERROR", "<yaml>", "STEP_LIST_MACRO_MISSING",
                    f"profile YAML runtime.step_list_macro '{macro}' has no matching Public Sub {macro} in any .bas module under vba_src/transforms"))

    # --- Render output
    errors = [v for v in all_violations if v.level == "ERROR"]
    warns = [v for v in all_violations if v.level == "WARN"]

    if not args.quiet and per_module_pass:
        print(f"OK {len(per_module_pass)} modules clean bytes+contract: {', '.join(sorted(per_module_pass))}")
    if args.fix and fixed_total:
        print(f"--fix applied: {fixed_total} non-ASCII bytes stripped across modules.")
    for v in all_violations:
        tag = f"[ERROR]" if v.level == "ERROR" else "[WARN ]"
        print(f"{tag} {v.module}: {v.rule} -- {v.detail}")

    elapsed = time.monotonic() - t0
    summary = (
        f"VBA hygiene: {len(modules)} modules in {elapsed*1000:.0f} ms — "
        f"{len(errors)} ERROR, {len(warns)} WARN "
        f"(checked: NON_ASCII/CRLF/COMPONENT_NAME_>31/LONG_LINES>250/WORKSHEETS(n)/RSTK_CALL_ARITY/RUN_SUB_SIG/STEP_LIST_MACRO_MISSING"
    )
    print(summary)
    if errors:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
