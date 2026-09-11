#!/usr/bin/env python3
"""Build the shipped Runner workbook from the generated shell and text VBA source."""

from __future__ import annotations

import argparse
import platform
import re
import shutil
import sys
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree


SCRIPT_PATH = Path(__file__).resolve()
MODULE_ROOT = SCRIPT_PATH.parents[1]
DEFAULT_RUNNER_SHELL = MODULE_ROOT / "runner_shell" / "RunnerShell.xlsm"
DEFAULT_VBA_SOURCE = MODULE_ROOT / "vba_src"
DEFAULT_PROFILE_DIR = MODULE_ROOT / "config" / "transforms"
DEFAULT_BUILD_MANIFEST = MODULE_ROOT / "config" / "build_manifest.yaml"
DEFAULT_REFERENCE_PATHS = MODULE_ROOT / "config" / "reference_paths.yaml"
DEFAULT_PATH_SUBSTITUTIONS = MODULE_ROOT / "config" / "root_path_substitutions.yaml"
DEFAULT_OUTPUT = MODULE_ROOT / "dist" / "Runner.xlsm"
CONFIG_SHEET_NAME = "config"
RUNNER_SHEET_NAME = "Runner"
STATUS_SOURCE = "Bootstrap Import Source"
STATUS_COUNT = "Bootstrap Import Count"
STATUS_AT = "Bootstrap Import At"
WORKBOOK_XML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REGISTRY_TITLE = "Transform Registry"
REGISTRY_SHEET_NAME = "registry"
REGISTRY_START_COLUMN = 1
REGISTRY_END_COLUMN = 14
REGISTRY_TITLE_ROW = 1
REGISTRY_HEADER_ROW = 2
REGISTRY_START_ROW = 3
RUNNER_SHEET_INDEX = 0
RUNNER_MAX_USER_COLUMN = 6
CONFIG_TITLE_ROW = 1
CONFIG_HEADER_ROW = 2
CONFIG_START_ROW = 3

# The Runner's runtime lookup sheet (read by VBA modHelpers.LookupMappingValue). Its rows
# are composed from ingest_lookups ref tables at build time — the same self-describing JSON
# mechanism every other reference dataset uses — so the mapping values live in config, not
# hand-authored inside the workbook. MAPPING_SHEET_NAME must match modHelpers
# MAPPING_LOOKUP_SHEET_NAME; the (entity, source, target) column order is the VBA contract.
MAPPING_SHEET_NAME = "mapping"
MAPPING_HEADER = ("entity", "source", "target")
# entity label -> (lookup name, source key, target key) for each ref table folded into the sheet.
MAPPING_LOOKUP_SOURCES = (
    ("company", "ref_company", "source", "target"),
    ("default", "ref_default", "domain", "value"),
)


@dataclass(frozen=True)
class BootstrapStatus:
    """Status values written back into the Runner workbook by the build step."""

    source: str
    import_count: int
    imported_at: str


@dataclass
class TransformBinding:
    """Build-time transform registration sourced from YAML profiles."""

    profile_path: Path
    profile_key: str
    profile_label: str
    source_type: str
    source_data: str
    target_eib: str
    output_file_path: str
    mapping_spreadsheet: str
    mapping_sheet: str
    customer_spreadsheet: str
    supplier_spreadsheet: str
    workday_reference_ids: str
    eib_population_type: str
    step_list_macro: str
    start_row: int | None
    # Extended source_layout fields: cold-start agent MUST read these from
    # the YAML profile, NOT infer them from the .bas. See also
    # config/eib_population_defaults.yaml for per-EIB-type shared constants
    # (target_sheet_name / target_data_start_row / clear_end_col).
    worksheet_name: str | None
    worksheet_resolution: str | None
    header_row: int | None
    data_start_row: int | None
    header_detection: str | None
    required_headers: list[str]
    # Per-transform constants that were previously ONLY in .bas Private Consts.
    # Cold-start agents MUST read this block rather than inferring from .bas
    # source code.  Schema enforced by _validate_transform_constants_schema().
    transform_constants_dict: dict | None


def create_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Copy the generated Runner shell to dist/Runner.xlsm, replace every "
            "importable .bas/.cls module from vba_src by running the workbook's own "
            "BootstrapImporter inside a live Excel host, and verify the saved "
            "workbook reflects the expected module set."
        )
    )
    parser.add_argument(
        "--runner-shell",
        default=str(DEFAULT_RUNNER_SHELL),
        help=f"Path to the generated Runner shell workbook (default: {DEFAULT_RUNNER_SHELL})",
    )
    parser.add_argument(
        "--vba-source",
        default=str(DEFAULT_VBA_SOURCE),
        help=f"Path to the text-source VBA folder (default: {DEFAULT_VBA_SOURCE})",
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT),
        help=f"Path to the built Runner workbook (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--profile-dir",
        default=str(DEFAULT_PROFILE_DIR),
        help=f"Path to the YAML transform profile folder (default: {DEFAULT_PROFILE_DIR})",
    )
    parser.add_argument(
        "--build-manifest",
        default=str(DEFAULT_BUILD_MANIFEST),
        help=(
            "Path to build_manifest.yaml (single source of truth for active/defunct "
            f"transform keys and the default profile key; default: {DEFAULT_BUILD_MANIFEST})"
        ),
    )
    parser.add_argument(
        "--verify",
        action="store_true",
        help=(
            "Skip prep and verify an already-built workbook instead. Use this after "
            "manually running BootstrapImporter.BootstrapImportFromConfiguredSource "
            "in Excel (see the instructions printed by the default mode)."
        ),
    )
    parser.add_argument(
        "--substitute-root-paths",
        action="store_true",
        help=(
            "Skip the build; instead patch the already-built workbook's "
            "reference/transform data-root path cells (mapping/customer/"
            "supplier_spreadsheet, source_data, target_eib, output_file_path) in "
            "place using root_path_substitutions.yaml, without re-copying the shell "
            "or touching anything else. A separate step from the build itself -- run "
            "after the default build, before the manual Import step "
            "(docs/BUILD_WIN.md Step 3)."
        ),
    )
    parser.add_argument(
        "--hygiene-only",
        action="store_true",
        help=(
            "Run ONLY the VBA source hygiene checks (ASCII, CRLF, 31-char component "
            "names, 250-char line ceiling, no index-based Worksheets(n) lookups, "
            "ResolveSelectedTransformKey/Run<name> arity contracts) and exit. "
            "Does not require Excel, pyOpenVBA, or a Runner shell. Fail fast on any "
            "rule violation so bad edits never reach the manual Bootstrap Import step."
        ),
    )
    parser.add_argument(
        "--path-substitutions",
        default=str(DEFAULT_PATH_SUBSTITUTIONS),
        help=(
            "Path to root_path_substitutions.yaml — build-time filesystem-path rewrites "
            f"applied to paths written into the workbook (default: {DEFAULT_PATH_SUBSTITUTIONS})"
        ),
    )
    parser.add_argument(
        "--path-profile",
        default=None,
        help=(
            "Named profile from root_path_substitutions.yaml to apply to build-written "
            "paths. Default: auto-selected from the OS via that file's "
            "os_default_profile (Windows -> forward-slash rewrite so "
            "BootstrapImporter can read the vba_src path; macOS/Linux -> none)."
        ),
    )
    return parser


def main() -> int:
    args = create_parser().parse_args()

    runner_shell = Path(args.runner_shell).expanduser().resolve()
    vba_source = Path(args.vba_source).expanduser().resolve()
    profile_dir = Path(args.profile_dir).expanduser().resolve()
    build_manifest_path = Path(args.build_manifest).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()
    path_substitution = _load_path_substitution(
        Path(args.path_substitutions).expanduser().resolve(), args.path_profile
    )

    module_names = _discover_importable_modules(vba_source)
    if not module_names:
        raise RuntimeError(f"No importable .bas/.cls files found under {vba_source}")
    transform_bindings = _load_transform_bindings(profile_dir, build_manifest_path)

    # ------------------------------------------------------------------
    # Pre-hygiene step: strip macOS "internet quarantine" extended attributes
    # from every importable VBA source file.
    #
    # RATIONALE (cold-start lesson, see docs/BUILD.md Bootstrap section):
    # Whenever Mac Excel/VBE saves or exports a .bas/.cls (or the file is
    # downloaded via any quarantine-aware app), Apple sets
    # com.apple.quarantine on the file.  When BootstrapImporter later calls
    # ThisWorkbook.VBProject.VBComponents.Import(<path>) via the VBA
    # automation bridge, the macOS Sandbox silently denies the read()
    # mid-flight and Excel surfaces the *generic* runtime error
    #   -2147219692 (80040714): Method 'Import' of object '_VBComponents' failed
    # with NO mention of quarantine or permissions.  This is the #1 cause of
    # BootstrapImporter import failures on macOS, and it is triggered by
    # ANY prior Excel-sourced edit to a .bas.  To avoid chasing our tails on
    # a symptom that isn't a code problem, we strip these xattrs here and
    # now, every build, before the hygiene pass or any Excel-side step.
    #
    # We also clear two related xattrs that historically caused the same
    # generic VBE Import() failure in earlier macOS builds:
    #   com.apple.provenance (set by newer Office save paths)
    #   com.apple.metadata:kMDItemWhereFroms (downloaded-from metadata)
    #
    # The guardrail in check_vba_hygiene.py _check_bytes also ERROR-classes
    # these xattrs so a user running only `--hygiene-only` still sees a
    # crisp diagnostic instead of the VBA 80040714 wall.
    # ------------------------------------------------------------------
    MACOS_XATTR_BLOCKLIST = (
        "com.apple.quarantine",
        "com.apple.provenance",
        "com.apple.metadata:kMDItemWhereFroms",
    )
    try:
        import xattr as _xattr  # type: ignore[import-not-found]
        have_xattr_pkg = True
    except Exception:
        have_xattr_pkg = False
    stripped_total = 0
    for module_name in module_names:
        matches = sorted(vba_source.rglob(f"{module_name}.*"))
        for pmp in matches:
            if not pmp.is_file() or pmp.suffix.lower() not in {".bas", ".cls"}:
                continue
            if have_xattr_pkg:
                try:
                    present = set(_xattr.listxattr(str(pmp)))  # type: ignore[attr-defined]
                    removed_here = 0
                    for xa in MACOS_XATTR_BLOCKLIST:
                        if xa in present:
                            try:
                                _xattr.removexattr(str(pmp), xa)  # type: ignore[attr-defined]
                                removed_here += 1
                            except Exception:
                                pass
                    stripped_total += removed_here
                except Exception:
                    pass
            else:
                import subprocess as _sp
                for xa in MACOS_XATTR_BLOCKLIST:
                    try:
                        r = _sp.run(["xattr", "-d", xa, str(pmp)],
                                    capture_output=True, check=False)
                        if r.returncode == 0:
                            stripped_total += 1
                    except Exception:
                        pass
    if stripped_total:
        print(f"[pre-hygiene] stripped macOS quarantine/provenance/where-from xattrs on {stripped_total} VBA source file(s).")

    # _validate_vba_source_hygiene prints warnings directly to stdout + raises on
    # ERROR-class violations. For --hygiene-only we catch-and-print then return 1.
    if args.hygiene_only:
        try:
            _validate_vba_source_hygiene(vba_source, module_names, transform_bindings, profile_dir)
        except SystemExit:
            raise
        except Exception as exc:
            print(f"HYGIENE FAILED: {exc}")
            return 1
        print(f"HYGIENE OK — {len(module_names)} modules, {len(transform_bindings)} bindings, 0 ERROR-class violations.")
        return 0
    _validate_vba_source_hygiene(vba_source, module_names, transform_bindings, profile_dir)

    if args.verify:
        return _verify_built_workbook(output_path, vba_source, module_names, transform_bindings)

    if args.substitute_root_paths:
        return _substitute_root_paths(output_path, transform_bindings, path_substitution)

    _validate_inputs(runner_shell, vba_source, profile_dir, output_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(runner_shell, output_path)

    print(f"Runner shell: {runner_shell}")
    print(f"VBA source:   {vba_source}")
    print(f"Profiles:     {profile_dir}")
    print(f"Output:       {output_path}")
    print(f"Modules:      {len(module_names)}")
    print(f"Bindings:     {len(transform_bindings)}")
    print(f"Path profile: {path_substitution.profile_name}")
    print()

    source_folder_value = path_substitution.rewrite(
        "bootstrap_source_folder", str(vba_source)
    )
    if source_folder_value != str(vba_source):
        print(
            f"Path substitution [{path_substitution.profile_name}] on bootstrap "
            f"source folder: {vba_source} -> {source_folder_value}"
        )

    # NOTE: this step does NOT rewrite the data-root paths sourced from
    # config/reference_paths.yaml and each transform's shell_defaults (source_data,
    # target_eib, output_file_path, mapping/customer/supplier_spreadsheet) -- they are
    # written here as authored (one mac path root), unlike bootstrap_source_folder
    # above. That is a separate, deliberate step: run
    # `build_runner.py --substitute-root-paths` after this build, before the manual
    # Import step. See docs/BUILD_WIN.md Step 3 and _substitute_root_paths() below.
    _write_label_value(output_path, SOURCE_FOLDER_INPUT_LABEL, source_folder_value)
    _clear_runner_tab_technical_columns(output_path)
    _write_transform_registry_block(output_path, transform_bindings)
    _write_shell_defaults_config_block(output_path, transform_bindings)
    _rewrite_runner_formulas_block(output_path, transform_bindings)
    _write_mapping_sheet_block(output_path)

    print("Prep complete. Next, apply the root-path substitution (Step 3):")
    print()
    print(f"  uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python "
          "elt_transform_excel/scripts/build_runner.py --substitute-root-paths")
    print()
    print("Then the module import itself must be run manually in Excel —")
    print("driving it headlessly via AppleScript proved unreliable (it breaks mid-import")
    print("unpredictably; see docs/RUNBOOK.md). To finish the build:")
    print()
    print(f"  1. Open {output_path} in Excel.")
    print("  2. Tools > Macro > Macros, run: BootstrapImporter.BootstrapImportFromConfiguredSource")
    print("     (the source folder is already filled in on the Runner sheet).")
    print("  3. The macro saves the workbook itself when it finishes.")
    print(f"  4. Verify with: uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project "
          "python elt_transform_excel/scripts/build_runner.py --verify")
    return 0


_DATA_ROOT_FIELDS = (
    "mapping_spreadsheet",
    "customer_spreadsheet",
    "supplier_spreadsheet",
    "source_data",
    "target_eib",
    "output_file_path",
)


def _substitute_root_paths(
    output_path: Path,
    transform_bindings: list["TransformBinding"],
    path_substitution: "PathSubstitution",
) -> int:
    """Patch the already-built workbook's data-root path cells in place -- a distinct
    step from the build itself (docs/BUILD_WIN.md Step 3: "Root-path substitution"),
    run after Step 2's `build_runner.py` build and before the manual Step 4 import.

    Does NOT copy the shell, does not touch the transform registry's non-path columns,
    and does not touch bootstrap_source_folder (written once, by Step 2's build; it
    does not depend on the data root -- see config/root_path_substitutions.yaml
    "WHY THIS EXISTS", job 2). Rewrites only mapping/customer/supplier_spreadsheet,
    source_data, target_eib, and output_file_path across every transform binding, via
    the same path_substitution profile used for job 2, then re-writes just the blocks
    those fields live in."""
    if not output_path.is_file():
        raise FileNotFoundError(
            f"Built workbook not found: {output_path}\n"
            "Run the build first (docs/BUILD_WIN.md Step 2): uv run --with pyOpenVBA "
            "--with openpyxl --with pyyaml --no-project python "
            "elt_transform_excel/scripts/build_runner.py"
        )
    print(f"Output:       {output_path}")
    print(f"Bindings:     {len(transform_bindings)}")
    print(f"Path profile: {path_substitution.profile_name}")
    print()

    data_root_rewrites = 0
    for binding in transform_bindings:
        for field_name in _DATA_ROOT_FIELDS:
            original = getattr(binding, field_name)
            rewritten = path_substitution.rewrite(field_name, original)
            if rewritten != original:
                setattr(binding, field_name, rewritten)
                data_root_rewrites += 1
    if data_root_rewrites:
        print(
            f"Path substitution [{path_substitution.profile_name}] on data-root "
            f"fields: {data_root_rewrites} value(s) rewritten across "
            f"{len(transform_bindings)} binding(s)."
        )
    else:
        print(
            f"Path substitution [{path_substitution.profile_name}]: no data-root "
            "values needed rewriting (already matched)."
        )

    _clear_runner_tab_technical_columns(output_path)
    _write_transform_registry_block(output_path, transform_bindings)
    _write_shell_defaults_config_block(output_path, transform_bindings)
    _rewrite_runner_formulas_block(output_path, transform_bindings)

    print()
    print("Root-path substitution complete. Continue with the manual Import step (Step 4).")
    return 0


def _verify_built_workbook(
    output_path: Path,
    vba_source: Path,
    module_names: list[str],
    transform_bindings: list[TransformBinding],
) -> int:
    if not output_path.exists():
        raise FileNotFoundError(f"Built workbook not found: {output_path}")

    status = _read_bootstrap_status(output_path)
    _verify_bootstrap_status(status, vba_source, module_names, output_path)
    _verify_transform_registry(output_path, transform_bindings)
    _verify_shell_defaults_config(output_path, transform_bindings)
    _verify_mapping_sheet(output_path)
    _verify_workbook_modules(output_path, module_names)

    print("Verification passed.")
    print(f"Built workbook: {output_path}")
    print(f"Imported modules: {len(module_names)}")
    print(f"Status timestamp: {status.imported_at}")
    return 0


def _validate_inputs(
    runner_shell: Path,
    vba_source: Path,
    profile_dir: Path,
    output_path: Path,
) -> None:
    if not runner_shell.exists():
        raise FileNotFoundError(
            f"Runner shell not found: {runner_shell}\n"
            "Create it first with:\n"
            "uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python "
            "elt_transform_excel/scripts/create_runner_shell.py"
        )
    if runner_shell.suffix.lower() != ".xlsm":
        raise ValueError(f"Runner shell must be an .xlsm workbook: {runner_shell}")
    if not vba_source.exists() or not vba_source.is_dir():
        raise FileNotFoundError(f"VBA source folder not found: {vba_source}")
    if not profile_dir.exists() or not profile_dir.is_dir():
        raise FileNotFoundError(f"Transform profile folder not found: {profile_dir}")
    if output_path == runner_shell:
        raise ValueError("--output must be different from --runner-shell")


def _discover_importable_modules(vba_source: Path) -> list[str]:
    module_names: list[str] = []
    seen_names: set[str] = set()

    for candidate in sorted(vba_source.rglob("*")):
        if not candidate.is_file() or candidate.suffix.lower() not in {".bas", ".cls"}:
            continue

        module_name = candidate.stem
        if module_name.lower() == "bootstrapimporter":
            continue

        key = module_name.lower()
        if key in seen_names:
            raise RuntimeError(
                f"Duplicate VBA component name '{module_name}' found under {vba_source}. "
                "Component names must remain unique across all importable files."
            )
        seen_names.add(key)
        module_names.append(module_name)

    return module_names


# House rule: ASCII + CRLF + 31-char component-name limit. Mac Excel's
# VBComponents.Import silently aborts on multi-byte UTF-8 in volume, on LF-only
# line endings, and on component names > 31 chars, falling back to a default
# "Module1" placeholder -- you only find out after running the manual import
# step in Excel. Fail fast here so the prep step can't pass a bad source set.
MAX_VBA_COMPONENT_NAME_CHARS = 31
WARN_MAX_SOURCE_LINE_CHARS = 250
ERROR_MAX_SOURCE_LINE_CHARS = 299  # BUILD.md rule 4: Mac VBA parser observed to fail mid-file on 258+ char lines; treat >= 300 as hard ERROR
WARN_LINE_CONTINUATION_RUN = 15    # >= 15 sequential line-continuations raises WARN
ERROR_LINE_CONTINUATION_RUN = 22  # >= 22 sequential line-continuations raises ERROR (hard Mac VBA parser continuation-stack failure observed at run=25)

# Guard (vii) PORTABLE_VBA_NO_ACTIVEX. Banned cross-reference / byte patterns that
# reference COM-only / ActiveX / Type-Library-only objects. These ship ONLY on Windows
# VBA hosts and raise Run-time error '429' "ActiveX component can't create object" on
# Mac Excel (and on locked-down Windows laptops with Trust Center "Disable all ActiveX").
# ARCHITECTURE.md § "Native VBA only — no COM-only libraries" is the canonical rule;
# this tuple is the mechanical guardrails implementation. Each pair is:
#   (short_label, bytes_regex_with_word_boundaries_as_appropriate).
# Patterns are case-insensitive via re.IGNORECASE at match time. COMMENT-only lines are
# excluded (first non-space char is "'" on the same line before the match) so modules
# can still name the banned libraries in rationale comments.
_ACTIVEX_BANNED_REF_PATTERNS: tuple[tuple[str, bytes], ...] = (
    # Banned type-library CLASS name references (use anywhere except in a comment line is ERROR)
    ("Scripting.Dictionary", rb"\bScripting\.Dictionary\b"),
    ("Scripting.FileSystemObject", rb"\bScripting\.FileSystemObject\b"),
    ("Scripting.Drive", rb"\bScripting\.Drive\b"),
    ("Scripting.Drives", rb"\bScripting\.Drives\b"),
    ("Scripting.File", rb"\bScripting\.File\b"),
    ("Scripting.Files", rb"\bScripting\.Files\b"),
    ("Scripting.Folder", rb"\bScripting\.Folder\b"),
    ("Scripting.Folders", rb"\bScripting\.Folders\b"),
    ("Scripting.TextStream", rb"\bScripting\.TextStream\b"),
    ("ADODB.Connection", rb"\bADODB\.Connection\b"),
    ("ADODB.Recordset", rb"\bADODB\.Recordset\b"),
    ("ADODB.Record", rb"\bADODB\.Record\b"),
    ("ADODB.Command", rb"\bADODB\.Command\b"),
    ("ADODB.Field", rb"\bADODB\.Field\b"),
    ("ADODB.Fields", rb"\bADODB\.Fields\b"),
    ("ADODB.Parameter", rb"\bADODB\.Parameter\b"),
    ("ADODB.Parameters", rb"\bADODB\.Parameters\b"),
    ("ADODB.Stream", rb"\bADODB\.Stream\b"),
    ("ADODB.Error", rb"\bADODB\.Error\b"),
    ("ADODB.Errors", rb"\bADODB\.Errors\b"),
    ("VBScript.RegExp", rb"\bVBScript\.RegExp\b"),
    ("VBScript.Match", rb"\bVBScript\.Match\b"),
    ("VBScript.Matches", rb"\bVBScript\.Matches\b"),
    ("VBScript.MatchCollection", rb"\bVBScript\.MatchCollection\b"),
    ("VBScript.SubMatches", rb"\bVBScript\.SubMatches\b"),
    ("ADOX.Catalog", rb"\bADOX\.Catalog\b"),
    ("ADOX.Table", rb"\bADOX\.Table\b"),
    ("DAO.DBEngine", rb"\bDAO\.DBEngine\b"),
    ("DAO.Database", rb"\bDAO\.Database\b"),
    ("DAO.Recordset", rb"\bDAO\.Recordset\b"),
    # Banned CreateObject() PROGID strings. CreateObject("Scripting.Dictionary") and
    # CreateObject('Scripting.Dictionary') variants, both case styles. These are the
    # late-bound form that sneaks past the reference list because it requires no
    # Tools > References tick. Both quote styles are guarded.
    ("CreateObject(\"Scripting.", rb'CreateObject\(\s*"Scripting\.'),
    ("CreateObject('Scripting.", rb"CreateObject\(\s*'Scripting\."),
    ("CreateObject(\"ADODB.", rb'CreateObject\(\s*"ADODB\.'),
    ("CreateObject('ADODB.", rb"CreateObject\(\s*'ADODB\."),
    ("CreateObject(\"VBScript.", rb'CreateObject\(\s*"VBScript\.'),
    ("CreateObject('VBScript.", rb"CreateObject\(\s*'VBScript\."),
    ("CreateObject(\"ADOX.", rb'CreateObject\(\s*"ADOX\.'),
    ("CreateObject('ADOX.", rb"CreateObject\(\s*'ADOX\."),
    ("CreateObject(\"DAO.", rb'CreateObject\(\s*"DAO\.'),
    ("CreateObject('DAO.", rb"CreateObject\(\s*'DAO\."),
)

# --- Spec <-> .bas column-agreement gate (mapping-SOT anti-drift) ---
# Parsed from the RAW spec text (not yaml.safe_load) because YAML 1.1 coerces bare
# Y/N/yes/no/on/off to booleans -- a `target: N` would silently become False.
_SPEC_LIST_TARGET_RE = re.compile(r'^\s*-\s*target:\s*"?([A-Z]{1,3})"?\s*$', re.MULTILINE)
_SPEC_INLINE_TARGET_RE = re.compile(r'^\s*"([A-Z]{1,3})"\s*:\s*\{', re.MULTILINE)


def _parse_spec_mapping_targets(spec_text: str) -> set[str]:
    """Uppercase EIB column letters declared as mapping targets in a transform_spec's
    `mapping:` block (both `- target: X` list form and `"X": { ... }` inline form).
    Scoped to the mapping block so unrelated keys elsewhere are never matched. Transforms
    that express their mapping differently (AMEX derived_fields, referee named blocks) yield
    an empty set and are simply not covered by this gate."""
    lines = spec_text.split("\n")
    start = None
    for i, line in enumerate(lines):
        if re.match(r'^mapping:\s*$', line):
            start = i
            break
    if start is None:
        return set()
    block_lines = []
    for line in lines[start + 1:]:
        if re.match(r'^[A-Za-z_][\w]*:', line):  # next top-level section ends the mapping block
            break
        block_lines.append(line)
    block = "\n".join(block_lines)
    return set(_SPEC_LIST_TARGET_RE.findall(block)) | set(_SPEC_INLINE_TARGET_RE.findall(block))


def _load_safe_shared_columns(defaults_path: Path) -> dict[str, frozenset]:
    """Per-EIB-population 'safe shared' plumbing columns: the generic ColumnIndex() writes that are
    NOT driven by a profile's worktag pipeline and so need not appear in a profile's eib_columns
    block. SINGLE SOURCE OF TRUTH is config/eib_population_defaults.yaml `safe_shared_columns:`
    (moved out of hard-coded frozensets in this file, for parity with customer_invoice_eib_columns).
    Letters are quoted in the YAML so PyYAML never coerces "N"/"Y" to booleans; str() defends anyway."""
    yaml = _import_yaml()
    data = yaml.safe_load(defaults_path.read_text(encoding="utf-8")) or {}
    out: dict[str, frozenset] = {}
    for pop, entry in (data.get("eib_populations") or {}).items():
        cols = (entry or {}).get("safe_shared_columns") or []
        out[str(pop)] = frozenset(str(c).strip() for c in cols if str(c).strip())
    return out


def _validate_vba_source_hygiene(
    vba_source: Path,
    module_names: list[str],
    transform_bindings: list[TransformBinding] | None = None,
    profile_dir: Path | None = None,
) -> None:
    import re

    # MULTILINE + search (not match on line 1): a .bas carries Attribute VB_Name
    # on its first line, but a .cls opens with VERSION / BEGIN / END header lines
    # and the attribute follows below -- anchoring to line 1 would silently skip
    # the internal-name checks for every class module. \s* before $ absorbs the
    # trailing CR on CRLF files.
    vb_name_re = re.compile(rb'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$', re.MULTILINE)
    errors: list[str] = []
    warnings: list[str] = []

    for module_name in module_names:
        # Locate the file on disk: rglob *.bas / *.cls
        candidate = None
        for ext in (".bas", ".cls"):
            for match in vba_source.rglob(module_name + ext):
                candidate = match
                break
            if candidate is not None:
                break
        if candidate is None:
            # Should not happen (module_names comes from the same rglob), but
            # don't abort the build for a missing hygiene input -- the later
            # Excel-side failure will be clearer.
            continue

        raw = candidate.read_bytes()

        # 1. Component name <= 31 chars, BOTH from filename stem AND from the
        #    internal Attribute VB_Name. VBA uses the Attribute value as the
        #    canonical component name; the filename is how the build discovers
        #    the file. Flag either as a violation.
        stem_name = candidate.stem
        if len(stem_name) > MAX_VBA_COMPONENT_NAME_CHARS:
            errors.append(
                f"filename stem '{stem_name}' is {len(stem_name)} chars, "
                f"exceeds the VBA component-name limit of {MAX_VBA_COMPONENT_NAME_CHARS}."
            )
        # Search a header window rather than only line 1: a class module's
        # VB_Name sits a few lines below its VERSION/BEGIN/END preamble. 4 KB
        # comfortably covers both a .bas first line and a .cls attribute block.
        attr_match = vb_name_re.search(raw[:4096])
        if attr_match:
            internal_name_bytes = attr_match.group(1)
            try:
                internal_name = internal_name_bytes.decode("ascii")
            except UnicodeDecodeError:
                errors.append(
                    f"Attribute VB_Name in '{candidate.name}' contains non-ASCII bytes; "
                    f"must be plain ASCII to import reliably."
                )
                internal_name = None
            if internal_name is not None:
                if len(internal_name) > MAX_VBA_COMPONENT_NAME_CHARS:
                    errors.append(
                        f"Attribute VB_Name '{internal_name}' in '{candidate.name}' is "
                        f"{len(internal_name)} chars, exceeds the VBA component-name limit "
                        f"of {MAX_VBA_COMPONENT_NAME_CHARS}."
                    )
                if internal_name != stem_name:
                    # Attribute VB_Name and filename stem disagree. The Attribute wins
                    # inside VBA but the build globs use the filename; flag as a warning
                    # because it can cause confusing verify-failures later.
                    warnings.append(
                        f"module filename stem '{stem_name}' != Attribute VB_Name "
                        f"'{internal_name}' in '{candidate.name}'. Keep them aligned."
                    )

        # 2. Pure ASCII (no multi-byte UTF-8). modHelpers / modFinanceCommon carry
        #    a tiny amount (< 20 bytes = 2 em dashes) and work; anything larger
        #    reliably fails import. Enforce pure ASCII on new files and flag
        #    > 20 non-ASCII bytes as an error, > 0 as a warning.
        non_ascii_count = sum(1 for b in raw if b >= 128)
        if non_ascii_count == 0:
            pass
        elif non_ascii_count <= 20:
            warnings.append(
                f"{non_ascii_count} non-ASCII bytes in '{candidate.name}' "
                f"(tolerated legacy: <= 20 bytes -- e.g. a handful of em dashes "
                f"-- but consider converting to ASCII equivalents)."
            )
        else:
            errors.append(
                f"{non_ascii_count} non-ASCII bytes in '{candidate.name}' "
                f"(> 20 byte threshold). Multi-byte UTF-8 in volume causes "
                f"VBComponents.Import to silently fall back to 'Module1' on Mac "
                f"Excel. Convert to ASCII: em dash -> --, box-rule -> -, arrow -> "
                f"->, curly quotes -> straight quotes."
            )

        # 3. CRLF line endings (not LF-only). Zero lone-LF means every LF has a
        #    preceding CR. count lone_LF = LF where byte before is not CR.
        lone_lf = 0
        for i in range(len(raw)):
            if raw[i : i + 1] == b"\x0a" and (i == 0 or raw[i - 1 : i] != b"\x0d"):
                lone_lf += 1
        if lone_lf > 0:
            errors.append(
                f"{lone_lf} LF-only line ending(s) in '{candidate.name}'. "
                f"Mac Excel's VBA importer requires CRLF everywhere. Fix: "
                f"sed -i '' 's/$/\\\\r/' {candidate} or a Python-based writer."
            )

        # 4. Line-length ceiling: WARN on 251..299 chars; ERROR on >= 300 chars.
        #    Anecdotally Mac Excel's VBA parser has failed import mid-file on
        #    258+ char source lines even when line-continuation is formally valid.
        lines = raw.split(b"\r\n") if b"\r\n" in raw else raw.split(b"\x0a")
        long_warn_lines: list[int] = []
        long_err_lines: list[int] = []
        for idx, line in enumerate(lines, 1):
            if len(line) >= ERROR_MAX_SOURCE_LINE_CHARS:
                long_err_lines.append(idx)
            elif len(line) > WARN_MAX_SOURCE_LINE_CHARS:
                long_warn_lines.append(idx)
        if long_err_lines:
            errors.append(
                f"{len(long_err_lines)} source line(s) >= {ERROR_MAX_SOURCE_LINE_CHARS} chars in "
                f"'{candidate.name}' (lines {', '.join(str(n) for n in long_err_lines[:5])}"
                f"{' ...' if len(long_err_lines) > 5 else ''}). Mac Excel's VBA importer "
                f"has been observed to abort import mid-file on 258+ char lines; "
                f"refactor with line-continuation ( _)."
            )
        if long_warn_lines:
            warnings.append(
                f"{len(long_warn_lines)} source line(s) 251-299 chars in "
                f"'{candidate.name}' (lines {', '.join(str(n) for n in long_warn_lines[:5])}"
                f"{' ...' if len(long_warn_lines) > 5 else ''}). Safe but consider "
                f"line-continuation ( _) to stay under the 250-char comfort ceiling."
            )

        # 4c. Consecutive line-continuation runs. Mac VBA parser has an internal
        #     continuation-merge stack that empirically fails somewhere between
        #     21 and 25 sequential '_' lines (Stationery/Taxi signatures failed
        #     at run=25 with the generic 80040714 Import() wall; CFA at 21 still
        #     worked by luck). New authors write 24+ param signatures with one
        #     parameter per line and hit this at Excel import time, at which
        #     point the error path is opaque unless someone walks the source
        #     with a continuation counter. Fail fast at build time instead:
        #     ERROR on >= 15 (10-under observed), WARN on 10-14.
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
            errors.append(
                f"{len(cont_runs_err)} continuation run(s) >= {ERROR_LINE_CONTINUATION_RUN} "
                f"sequential '_' markers in '{candidate.name}' ({examples}). "
                f"Mac VBA Import parser continuation-stack crashes at 24+ and "
                f"can fail as low as 21. Pack params denser into fewer lines "
                f"(max 2-3 continuations per logical construct) or use a local "
                f"params UDT / intermediate vars. Last observed Stationery/Taxi "
                f"80040714 at run=25."
            )
        if cont_runs_warn:
            examples = "; ".join(f"lines {s}-{s+r-1} ({r})" for s, r in cont_runs_warn[:3])
            warnings.append(
                f"{len(cont_runs_warn)} continuation run(s) {WARN_LINE_CONTINUATION_RUN}-"
                f"{ERROR_LINE_CONTINUATION_RUN - 1} sequential '_' in "
                f"'{candidate.name}' ({examples}). Safe but keep under "
                f"{ERROR_LINE_CONTINUATION_RUN}."
            )

        # 4b. No index-based sheet access. `Worksheets(1)`, `Worksheets(0)`, or
        #     `Worksheets(n)` where n is a literal integer are fragile -- a tab
        #     drag by Finance reorders the workbook and every hardcoded index
        #     silently points to the wrong sheet. All sheet lookups MUST use a
        #     named constant instead (Worksheets(RUNNER_SHEET_NAME), etc.).
        #
        #     Exception (legacy): modDispatcher still carries an inline comment
        #     reference ("Worksheets(1) is fragile") -- allow text inside a
        #     VBA single-line comment (starting with ') so the explanatory
        #     warning isn't itself a build blocker.
        idx_lookup_re = re.compile(rb"\bWorksheets\s*\(\s*(\d+)\s*\)")
        for idx_match in idx_lookup_re.finditer(raw):
            match_start = idx_match.start()
            line_start = raw.rfind(b"\n", 0, match_start) + 1
            line_prefix = raw[line_start:match_start]
            # Line has a VBA single-line comment (apostrophe) anywhere before
            # the match: skip.
            if b"'" in line_prefix:
                continue
            line_no = raw[:match_start].count(b"\n") + 1
            errors.append(
                f"Index-based sheet lookup Worksheets({idx_match.group(1).decode('ascii')}) "
                f"in {candidate.name}:{line_no}. Never reference sheets by numeric "
                f"index -- a tab drag reorders the workbook and silently breaks "
                f"lookups. Use a named constant: Worksheets(RUNNER_SHEET_NAME), "
                f"Worksheets(REGISTRY_SHEET_NAME), etc. See modHelpers.bas for "
                f"the shared constants."
            )

        # 4d. Guard (vii) PORTABLE_VBA_NO_ACTIVEX. COM-only / ActiveX / Type-Library-only
        #     references are strictly banned. On Mac Excel and on Trust-Center-locked
        #     Windows laptops, CreateObject("Scripting.Dictionary") raises Run-time
        #     error '429' "ActiveX component can't create object" even when Tools >
        #     References shows nothing ticked (it's a late-bound call, no reference).
        #     The canonical rule list is ARCHITECTURE.md § "Native VBA only — no COM-only
        #     libraries"; the mechanical ban list lives at _ACTIVEX_BANNED_REF_PATTERNS
        #     above.
        #     As with 4b (Worksheets(n)), comment-only lines are excluded: if the first
        #     non-space byte of the source line BEFORE the match offset is a VBA
        #     apostrophe comment marker (0x27 = "'"), skip it. This lets modules name
        #     the banned libraries in rationale comments without triggering the guard.
        for short_label, banned_pattern in _ACTIVEX_BANNED_REF_PATTERNS:
            compiled = re.compile(banned_pattern, re.IGNORECASE)
            for ax_match in compiled.finditer(raw):
                match_start = ax_match.start()
                line_start = raw.rfind(b"\n", 0, match_start) + 1
                line_prefix = raw[line_start:match_start]
                line_stripped_prefix = line_prefix.lstrip()
                # Comment-only line (apostrophe is the first non-space char on
                # the line before this match):
                if line_stripped_prefix.startswith(b"'"):
                    continue
                line_no = raw[:match_start].count(b"\n") + 1
                # Try to grab the matched snippet in human-readable form
                try:
                    snippet = ax_match.group(0).decode("ascii", errors="replace")
                except Exception:  # noqa: BLE001 -- defensive, bytes decode never fails with replace
                    snippet = repr(ax_match.group(0))
                errors.append(
                    f"[Guard (vii): PORTABLE_VBA_NO_ACTIVEX] COM-only / ActiveX "
                    f"reference '{short_label}' -> snippet '{snippet}' found in "
                    f"{candidate.name}:{line_no}. These libraries ship ONLY on "
                    f"Windows VBA hosts and raise Run-time error '429' 'ActiveX "
                    f"component can't create object' on Mac Excel (engineering "
                    f"authoring host) AND on locked-down Windows laptops with "
                    f"Trust Center 'Disable all ActiveX' policy. Mac-safe "
                    f"replacements: VBA Collection for dictionary semantics "
                    f"(see CollectionKeyExists / CollectionSetItemLong / "
                    f"CollectionGetItemLong helpers in modApCfaClaimsSupplierInvoice.bas); "
                    f"Open / Print # / Input # for file I/O (no "
                    f"Scripting.FileSystemObject); Worksheet array reads + "
                    f"hand-rolled tokenization instead of ADODB SQL / "
                    f"VBScript.RegExp. See canonical rule in ARCHITECTURE.md "
                    f"§ 'Native VBA only — no COM-only libraries'."
                )

    # 5. Signature contract checks. The shared-dispatcher contract between the
    #    per-transform .bas modules and the shared helpers/dispatcher evolves
    #    (e.g. ResolveSelectedTransformKey grew a runnerWorkbook argument when
    #    the registry moved off Worksheets(1) to its own VeryHidden sheet).
    #    Cold-start sessions copying old 1-arg call patterns from memory or
    #    from an early exemplar silently build-clean and then fail at runtime
    #    inside a user-pressed Run Transform. Fail the build on any drift.
    #
    #    The contract for .bas modules (transforms and the dispatcher) is:
    #    a. EVERY CALL SITE to ResolveSelectedTransformKey passes exactly 2
    #       untyped positional args: (runnerWorkbook, controlSheet). The
    #       FUNCTION DECLARATION has typed params and is excluded from this
    #       check (we inspect declaration shape separately below).
    #    b. EVERY public entry-point Sub Run<Name> in a TRANSFORM .bas (not
    #       the top-level RunSelectedTransform button-dispatcher macro, which
    #       receives zero args because Form Controls bind to zero-arg macros)
    #       has the exact 2-arg signature:
    #       Public Sub Run<Name>(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet)
    #    c. EVERY profile-YAML runtime.step_list_macro name appears as a
    #       Public Sub name somewhere in vba_src/.

    # a. ResolveSelectedTransformKey CALL SITES (not declarations) arity.
    #    Declaration: "Public Function ResolveSelectedTransformKey(ByVal ... As Workbook, ByVal ... As Worksheet) As String"
    #    Call site (untyped): "ResolveSelectedTransformKey(runnerWorkbook, controlSheet)"
    rstk_decl_re = re.compile(
        rb"^\s*(Public|Private|Friend)\s+(Function|Sub|Property)\s+ResolveSelectedTransformKey\b",
        re.MULTILINE,
    )
    rstk_call_re = re.compile(rb"\bResolveSelectedTransformKey\s*\(([^()]*)\)", re.DOTALL)
    rstk_call_sites: list[tuple[Path, int, bytes]] = []
    # b. Public Sub Run<Name> declaration captures.
    run_sub_decl_re = re.compile(
        rb"^\s*Public\s+Sub\s+(Run[A-Za-z0-9_]+)\s*\(\s*([^)]*)\s*\)\s*$", re.MULTILINE
    )
    public_run_subs: dict[str, tuple[Path, bytes]] = {}
    # Tracks declaration line ranges so we skip them when matching call sites
    # below. (We can't just negative-lookbehind the declaration because the
    # capture for args sits on its own).
    module_decl_lines: dict[Path, set[int]] = {}

    for module_name in module_names:
        candidate = None
        for ext in (".bas", ".cls"):
            for match in vba_source.rglob(module_name + ext):
                candidate = match
                break
            if candidate is not None:
                break
        if candidate is None:
            continue
        raw = candidate.read_bytes()
        # Shared LF-normalised index map: `cr_before[i]` == number of \r bytes
        # present in raw[0:i]. With this, `lf_index = raw_index - cr_before[raw_index]`
        # translates any raw byte offset into the equivalent offset in an
        # LF-only view of the file, and line counting on `raw_lf_only[:lf_index]`
        # is stable against: (a) CRLF vs LF mix, (b) regex ^ anchored at \n that
        # would capture a preceding \r in leading \s*, (c) inline matches that
        # start mid-line.
        raw_lf_only = raw.replace(b"\r", b"")
        cr_before: list[int] = [0] * (len(raw) + 1)
        cr_count = 0
        for i in range(len(raw)):
            if raw[i : i + 1] == b"\r":
                cr_count += 1
            cr_before[i + 1] = cr_count

        def _line_no_at(raw_index: int) -> int:
            lf_index = raw_index - cr_before[min(raw_index, len(raw))]
            return raw_lf_only[:lf_index].count(b"\n") + 1

        decl_lines: set[int] = set()
        for m in rstk_decl_re.finditer(raw):
            line_no = _line_no_at(m.start())
            decl_lines.add(line_no)
            decl_lines.add(line_no + 1)
        module_decl_lines[candidate] = decl_lines

        for call_match in rstk_call_re.finditer(raw):
            args_bytes = call_match.group(1)
            line_no = _line_no_at(call_match.start())
            if line_no in decl_lines:
                continue
            rstk_call_sites.append((candidate, line_no, args_bytes))

        for sub_match in run_sub_decl_re.finditer(raw):
            sub_name_bytes = sub_match.group(1)
            args_bytes = sub_match.group(2)
            try:
                sub_name = sub_name_bytes.decode("ascii")
            except UnicodeDecodeError:
                sub_name = None
                errors.append(
                    f"Public Sub name contains non-ASCII bytes in '{candidate.name}'. "
                    f"Must be ASCII as it is invoked by string name via Application.Run."
                )
            if sub_name:
                if sub_name in public_run_subs:
                    errors.append(
                        f"Duplicate Public Sub declaration '{sub_name}': appears in "
                        f"both {public_run_subs[sub_name][0].name} and {candidate.name}. "
                        f"The dispatcher resolves macros by bare name, so duplicates are a race."
                    )
                public_run_subs[sub_name] = (candidate, args_bytes)

    for file_path, line_no, args_bytes in rstk_call_sites:
        args_stripped = args_bytes.replace(b"\r", b"").replace(b"\n", b"").replace(b" ", b"")
        expected_stripped = b"runnerWorkbook,controlSheet"
        if args_stripped != expected_stripped:
            errors.append(
                f"{file_path.name}:{line_no} ResolveSelectedTransformKey called with args "
                f"{args_bytes!r}. The enforced 2-arg call-site contract is "
                f"ResolveSelectedTransformKey(runnerWorkbook, controlSheet). "
                f"Shared-helper signature drift breaks runtime dispatch."
            )

    # RunSelectedTransform is the ONLY zero-arg allowed exception: it's the
    # Form-Control button binding target (buttons pass no args). All other
    # Public Sub Run<...> must match the 2-arg contract.
    DISPATCHER_BUTTON_ENTRYPOINT = "RunSelectedTransform"
    EXPECTED_RUN_SUB_ARGS_STRIPPED = (
        b"ByValrunnerWorkbookAsWorkbook,ByValcontrolSheetAsWorksheet"
    )
    for sub_name, (file_path, args_bytes) in public_run_subs.items():
        if sub_name == DISPATCHER_BUTTON_ENTRYPOINT:
            args_stripped = args_bytes.replace(b"\r", b"").replace(b"\n", b"").replace(b" ", b"")
            if args_stripped != b"":
                errors.append(
                    f"{file_path.name} Public Sub {sub_name}(...) should be zero-arg because it "
                    f"is the Form-Control button macro target (buttons pass no args). Got: {args_bytes!r}."
                )
            continue
        args_stripped = args_bytes.replace(b"\r", b"").replace(b"\n", b"").replace(b" ", b"")
        if args_stripped != EXPECTED_RUN_SUB_ARGS_STRIPPED:
            errors.append(
                f"{file_path.name} Public Sub {sub_name}(...) has unexpected signature. "
                f"The transform entry-point contract requires exactly: "
                f"Public Sub {sub_name}(ByVal runnerWorkbook As Workbook, ByVal controlSheet As Worksheet). "
                f"(Zero-arg button-dispatcher is only '{DISPATCHER_BUTTON_ENTRYPOINT}'.) "
                f"Got args (after stripping whitespace/line-continuations): {args_bytes!r}."
            )

    if transform_bindings is not None:
        macro_names_seen = {sub_name for sub_name in public_run_subs.keys()}
        for binding in transform_bindings:
            expected_macro = binding.step_list_macro
            if not expected_macro:
                errors.append(
                    f"profile '{binding.profile_key}' has empty runtime.step_list_macro. "
                    f"The dispatcher needs it to find the entry-point Sub via the registry."
                )
                continue
            if expected_macro not in macro_names_seen:
                errors.append(
                    f"profile '{binding.profile_key}' declares step_list_macro "
                    f"'{expected_macro}', but no Public Sub with that exact name was "
                    f"found under {vba_source}. (Found: {sorted(macro_names_seen)}). "
                    f"This transform will fail at runtime with 'macro not found' when the "
                    f"user presses Run Transform."
                )

    # ======================================================================
    # STEP 6. Config ↔ VBA bi-directional guardrails.
    # CANONICAL RULE: the CODING AGENT drives .bas implementation FROM the
    # YAML config only.  The VBA runtime NEVER parses or reads the YAML
    # profiles — instead, the agent authors the .bas modules by reading the
    # declared keys/columns/derivations and translating them to Private
    # Const WBO_* declarations, ResolveReportingTypeFromFinalCode() calls,
    # and ColumnIndex("XX") writes.  The .bas never reads config.
    #
    # If a profile has an ACTIVE worktag pipeline (worktag_business_objects
    # or fixed_literal_worktags declared in transform_constants) the WBO
    # strings, derived-worktag resolution calls, and ColumnIndex() write
    # targets must 1:1 match what the .bas module declares/calls.  Any
    # mismatch == build ERROR — the guardrails stop the agent from authoring
    # a .bas that contradicts the config (or, conversely, declaring config
    # that the .bas doesn't honour).
    # ======================================================================
    if transform_bindings is not None:
        # Map the transform profile_key -> module_path containing the Public Sub entry.
        profile_to_module: dict[str, Path] = {}
        for binding in transform_bindings:
            expected_macro = binding.step_list_macro
            if not expected_macro:
                continue
            for sub_name, (file_path, _args_bytes) in public_run_subs.items():
                if sub_name == expected_macro:
                    profile_to_module[binding.profile_key] = file_path
                    break

        # --- Tier-1 hoisting: WBO + TAX constants live as PUBLIC in modFinanceCommon.bas ---
        # AP transform modules (modAp*) now reference bare names that resolve to Public scope.
        # AR transform modules (modAr*) are out of scope for this rollout and keep Private Const
        # copies (existing logic for AR is preserved below).
        #
        # Extract the canonical source ONCE from modFinanceCommon.bas, then per-AP-module:
        #   (NEG) the module MUST NOT contain Private Const copies of WBO_* or TAX_*
        #   (POS) the module does reference the hoisted bare names (e.g. WBO_COST_CENTER token)
        #   (SYNC) the public modFinanceCommon values match the YAML-declared values 1:1
        vba_src_root = vba_source
        if isinstance(vba_source, list) and vba_source:
            vba_src_root = vba_source[0].parent if hasattr(vba_source[0], "parent") else vba_source[0]
        mod_finance_common_path = vba_src_root / "modFinanceCommon.bas"
        mfc_raw = mod_finance_common_path.read_bytes() if mod_finance_common_path.exists() else b""

        def _extract_public_string_consts(payload: bytes, prefix: str) -> dict[str, str]:
            """Return {const_name: literal_value} for every `Public Const <prefix>X As String = "..."`."""
            rgx = re.compile(
                rf'^\s*Public\s+Const\s+({re.escape(prefix)}[A-Za-z0-9_]+)\s+As\s+String\s*=\s*"([^"]+)"\s*$'.encode(
                    "ascii"
                ),
                re.MULTILINE,
            )
            out: dict[str, str] = {}
            for m in rgx.finditer(payload):
                try:
                    out[m.group(1).decode("ascii")] = m.group(2).decode("ascii")
                except UnicodeDecodeError:
                    pass
            return out

        public_wbo_consts: dict[str, str] = _extract_public_string_consts(mfc_raw, "WBO_")
        public_tax_consts: dict[str, str] = _extract_public_string_consts(mfc_raw, "TAX_")

        # Map concern -> bare VBA helper names that MUST NOT exist as Private Function copies
        HOISTED_HELPER_PRIVATE_FN_NAMES: tuple[str, ...] = (
            "NormaliseToWordBoundaries",
            "NormalizeCostCodeSeparators",
            "NormalizeSourceDate",
            "RowIsTrulyEmpty",
            "RowIsEmptyOrTrailer",
            "ResolveEffectiveNet",
        )
        # For each concern/variant -> list of bare helper names that MUST be referenced (POS)
        # when the variant != "none". NOTE: only helpers DIRECTLY CALLED in the transform
        # .bas body (not internally inside a hoisted helper) are checked. "NormaliseToWordBoundaries"
        # is an internal detail of the hoisted RowIsEmptyOrTrailer() implementation in modSourceTable;
        # "FinalCodeSegment" is an internal modFinanceCommon primitive used by the public
        # ResolveReportingTypeFromFinalCode / ResolveLowestHierarchyLookupAcrossFinalCodeSegments
        # entry points — the transform .bas never calls them by bare name directly.
        CONCERN_VARIANT_TO_BARE_HELPER_HINTS: dict[tuple[str, str], tuple[str, ...]] = {
            # supplier_resolution: the split-lookup variant MUST invoke the shared resolver.
            # A module declaring this while doing an inline passthrough (0 real calls) is the
            # exact drift that shipped the 3 Workday-ID transforms mislabelled. The
            # direct_workday_id_passthrough variant is inline (no shared fn) and therefore has
            # NO hint here, so the POS guard correctly skips it.
            ("supplier_resolution", "per_row_split_company_creditor"): (
                "ResolveSupplierWithCompanyPrecedence",
            ),
            ("accounting_string_derivation", "direct_column_separator_normalized"): (
                "NormalizeCostCodeSeparators",
            ),
            ("row_net_resolution", "effective_net_gross_tax_fallback"): ("ResolveEffectiveNet",),
            ("row_trailer_detection", "text_cols_and_amount_cols_parameterized"): (
                "RowIsTrulyEmpty",
                "RowIsEmptyOrTrailer",
            ),
            # source_date_normalization: every transform that declares this concern MUST call
            # NormalizeSourceDate at runtime. The REVERSE guard below catches the opposite
            # direction: modules that CALL it but FORGOT to declare the concern.
            ("source_date_normalization", "excel_serial_and_dmy_text"): ("NormalizeSourceDate",),
            # customer_name_resolution: every transform selecting the two-level GBE cascade
            # MUST invoke both the lookup builder AND the resolver at runtime. The bare-name
            # POS guard catches drift where a module declares the shared variant but still
            # carries its own inline lookup (the pre-hoist pattern in modArGbeInvoiceCustInvoice).
            # NormalizeClubName / ApplySqlNameTransforms are internal primitives consumed by
            # the public resolver; the POS guard intentionally only enforces the two public
            # entry points (builder + resolver) so future profile-specific optimizations may
            # call the primitives directly without tripping the guard.
            ("customer_name_resolution", "two_level_l1_l2_sql_migration"): (
                "BuildCustomerNameLookupFromWorkbook",
                "ResolveCustomerIdMultiLevel",
            ),
        }
        # Reverse-direction anti-drift map: for a shared helper that is (a) public and
        # (b) used across all/most transforms, if the bare helper NAME is referenced in
        # the transform .bas body then the listed concern MUST be present and declare a
        # variant != "none" in the transform spec. This is the guard that catches the
        # exact incident we had on NormalizeSourceDate: every AR module called the helper
        # at runtime (and had done for months), yet the 5 AR specs had no shared_helpers:
        # block at all and the registry didn't even declare source_date_normalization —
        # so a cold-start author reading only config/docs could have authoritatively
        # said "no shared date coercion exists" and then incorrectly re-implemented the
        # same behavior inline in a new AR module.
        BARE_HELPER_TO_REQUIRED_CONCERN: dict[str, str] = {
            "NormalizeSourceDate": "source_date_normalization",
        }
        # --- Stage 1: shared_helper_registry.yaml + transform_specs shared_helpers blocks ---
        # Declarative anti-drift: each transform spec declares which shared helper variant
        # it uses for each concern. Step 6 enforces:
        #   (1) the declaration EXISTS (AP + AR both required; AR was deferred until v46.1
        #       EIB templates were correct and AR specs carried a shared_helpers block)
        #   (2) each selected variant RESOLVES to a valid key in the single catalog registry
        #   (3) for each concern/variant declared in the registry that carries
        #       vba_module/vba_public_fn, the corresponding Public export must appear as
        #       a declared signature in the target VBA module (registry→VBA drift guard)
        #   (4) NEG: transform .bas (AP+AR) does NOT contain Private Function copies of
        #       any hoisted helper
        #   (5) POS (forward): if variant != "none", the bare helper hint names are
        #       referenced in the .bas
        #   (6) POS (reverse): BARE_HELPER_TO_REQUIRED_CONCERN helper name seen in .bas →
        #       the concern MUST be declared in the spec (and variant != "none")
        registry_catalog: dict[str, dict[str, dict]] = {}
        transform_specs_dir: Path | None = None
        safe_shared_by_pop: dict[str, frozenset] = {}
        _yaml_lib = _import_yaml()
        if profile_dir is not None:
            safe_shared_by_pop = _load_safe_shared_columns(
                profile_dir.parent / "eib_population_defaults.yaml"
            )
            registry_path = profile_dir.parent / "shared_helper_registry.yaml"
            if registry_path.exists():
                with registry_path.open("r", encoding="utf-8") as _h:
                    _registry_root = _yaml_lib.safe_load(_h) or {}
                if isinstance(_registry_root, dict) and isinstance(_registry_root.get("shared_helpers"), dict):
                    for _concern, _variants in _registry_root["shared_helpers"].items():
                        if isinstance(_variants, dict):
                            registry_catalog[str(_concern)] = {
                                str(k): v if isinstance(v, dict) else {} for k, v in _variants.items()
                            }
            transform_specs_dir = profile_dir.parent / "transform_specs"

        def _strip_vba_comments(raw_module: bytes) -> bytes:
            """Remove VBA line comments (`'` to end-of-line, respecting double-quoted strings)
            so a DEPRECATED mention of a helper in a comment cannot satisfy the POS anti-drift
            guard. Without this, a `'  ... ResolveSupplierWithCompanyPrecedence (RETIRED)` comment
            counts as a call site and lets a mislabelled shared_helpers declaration pass."""
            out: list[bytes] = []
            for line in raw_module.split(b"\n"):
                in_str = False
                cut = len(line)
                for i in range(len(line)):
                    ch = line[i : i + 1]
                    if ch == b'"':
                        in_str = not in_str
                    elif ch == b"'" and not in_str:
                        cut = i
                        break
                out.append(line[:cut])
            return b"\n".join(out)

        def _bare_name_is_referenced(raw_module: bytes, bare_name: str) -> bool:
            """True if the .bas module references `bare_name` as a bare identifier token at a
            call site (not in a `Const <bare_name>` declaration line and not inside a comment),
            so we know it is consumed at runtime. VBA resolves unqualified standard-module Public
            names via normal project scope. Hybrid scope (local Private Const shadowing the shared
            Public Const with the same bare name): correctly counts as runtime use — we do NOT short-circuit
            on the declaration line; we DELETE declaration lines first then check the remaining body for
            the runtime token reference."""
            raw_module = _strip_vba_comments(raw_module)
            decl_stripper = re.compile(
                rf'^\s*(?:Private|Public|Dim|Global|Const|Type)\s+.*\b{re.escape(bare_name)}\b.*$\n?'.encode(
                    "ascii"
                ),
                re.MULTILINE,
            )
            raw_body = decl_stripper.sub(b"", raw_module)
            token = re.compile(rf"\b{re.escape(bare_name)}\b".encode("ascii"))
            return bool(token.search(raw_body))

        wbo_const_re = re.compile(
            rb'^\s*Private\s+Const\s+(WBO_[A-Za-z0-9_]+)\s+As\s+String\s*=\s*"([^"]+)"\s*$',
            re.MULTILINE,
        )
        reportingtype_call_re = re.compile(rb'\bResolveReportingTypeFromFinalCode\s*\(')
        column_index_letter_re = re.compile(rb'ColumnIndex\s*\(\s*"([A-Z]{1,3})"\s*\)')

        # Shell-variant spec cache: multiple bindings derived from the same base
        # profile (via variants) share ONE module AND ONE transform_spec file.
        # The base transform_spec is named after the BASE profile_key, not the
        # variant-substituted one (e.g. ap_stationery_addison_lee_fa shares the
        # ap_stationery.yaml spec). Cache: module_path -> shared_helpers dict.
        module_to_spec_shared_helpers: dict[Path, dict[str, object]] = {}

        # ======================================================================
        # [PRE-LOOP GUARD: registry → VBA linkage]
        # Runs EXACTLY ONCE per hygiene pass (not per binding, not per module).
        # For each concern/variant in the registry that declares vba_module and
        # vba_public_fn:
        #   (a) the target VBA module file must exist under {vba_source}
        #   (b) every identifier listed in vba_public_fn must appear as a Public
        #       Function|Sub|Property signature inside that module.
        #
        # vba_public_fn accepts flexible shapes (to match the existing mixed
        # style in shared_helper_registry.yaml):
        #   - YAML string, single name:        "NormalizeSourceDate"
        #   - YAML string, multi-name prose:   "BuildReferenceLookupFromWorkbook
        #                                        + LookupCollectionValue"
        #                                        "SplitSupplierIdRow (...) tuple"
        #   - YAML list of names/prose strings: ["BuildBusinessObjectHierarchyLookup",
        #                                        "ResolveLowestHierarchy...(...)"]
        # From any of these we extract bare identifier tokens and validate each
        # one. Parenthesised signatures, trailing commentary text, and ' + ' /
        # ' and ' separators are all stripped silently.
        # ======================================================================
        _registry_vba_verified: set[tuple[str, str]] = set()
        _module_paths_by_stem: dict[str, Path] = {}
        for _nm in module_names:
            for _ext in (".bas", ".cls"):
                for _m in vba_source.rglob(_nm + _ext):
                    _module_paths_by_stem[_m.stem] = _m
                    break

        def _extract_bare_public_names(raw_value: object) -> list[str]:
            """Given a raw vba_public_fn value, return bare identifier tokens.

            Accepted input shapes (to match the mixed style in the registry):
              - YAML string, single name:        "NormalizeSourceDate"
              - YAML string, multi-name prose:   "BuildReferenceLookupFromWorkbook
                                                   + LookupCollectionValue"
                                                   "SplitSupplierIdRow (...) tuple"
              - YAML list of names/prose strings: ["BuildBusinessObjectHierarchyLookup",
                                                   "ResolveLowestHierarchy...(params)"]

            We only accept identifiers that look like VBA *function/helper names*,
            i.e. PascalCase / CamelCase tokens that START WITH A CAPITAL LETTER and
            contain at least one more letter/number. This deliberately drops
            signature parameter names (amountCheckCols, labelCheckCols, textCheckCols,
            table, row, vendorId, company, ...) which are always lowercase-start
            or all-lowercase inside the parenthesised signature portion of the YAML.
            """
            candidates: list[str] = []
            if isinstance(raw_value, list):
                for elem in raw_value:
                    candidates.extend(_extract_bare_public_names(elem))
                return candidates
            if not isinstance(raw_value, str):
                return []
            text = raw_value
            # Normalise multi-name separators
            text = text.replace(" + ", " ").replace(" and ", " ")
            # Strip parenthesised signature material (nested "FnName(arg, arg2())" →
            # the inner () pairs are handled by repeated substitution until stable).
            prev = None
            while prev != text:
                prev = text
                text = re.sub(r"\([^()]*\)", " ", text)
            tokens = re.findall(r"[A-Za-z_][A-Za-z0-9_]*", text)
            accepted: list[str] = []
            STOPWORDS = {
                "returns", "tuple", "and", "or", "with", "from", "as", "is",
                "return", "Table", "Row", "List", "Array", "Dict", "Set",
                "Lookup", "Collection", "Value", "Values", "Workbook", "Worksheet",
            }
            for tok in tokens:
                # PascalCase/CamelCase function-style gate: starts with capital,
                # has at least one alphanumeric char after it. Rejects lowercase
                # params (amountCheckCols, labelCheckCols, vendorId, company, ...)
                # even when they contain internal capitals.
                if len(tok) < 3:
                    continue
                if not tok[:1].isupper():
                    continue
                if tok in STOPWORDS:
                    continue
                accepted.append(tok)
            return accepted

        for _reg_concern, _reg_variants in registry_catalog.items():
            for _reg_variant, _reg_payload in _reg_variants.items():
                _pair_key = (_reg_concern, _reg_variant)
                if _pair_key in _registry_vba_verified:
                    continue
                _registry_vba_verified.add(_pair_key)
                if not isinstance(_reg_payload, dict):
                    continue
                _target_module = str(_reg_payload.get("vba_module") or "").strip()
                _raw_public_fn = _reg_payload.get("vba_public_fn")
                if not _target_module:
                    continue  # null / empty vba_module == inline-variant, skip guard
                # Inline-variant indicator: the registry value is prose like
                # "local transform (...)" or "local transform" that explicitly
                # says the variant is implemented inline in each transform rather
                # than as a shared Public export. These are semantically the
                # same as a null vba_module (the author marked them as inline).
                _target_lower = _target_module.lower()
                if (
                    "local transform" in _target_lower
                    or "one-off pattern" in _target_lower
                    or (" " in _target_module and not any(_target_module.endswith(_ext) for _ext in (".bas", ".cls")))
                ):
                    continue
                _bare_names = _extract_bare_public_names(_raw_public_fn)
                if not _bare_names and _raw_public_fn not in (None, ""):
                    # vba_public_fn is present but yielded zero identifiers —
                    # likely a prose string the parser rejected. Surface as a
                    # warning so the author can tighten the shape in YAML.
                    warnings.append(
                        f"shared_helper_registry.yaml: concern={_reg_concern!r} "
                        f"variant={_reg_variant!r} has vba_public_fn=... but no "
                        "bare identifier tokens could be extracted. Consider "
                        "writing vba_public_fn as a simple name, a 'A + B' string, "
                        "or a YAML list [\"NameA\", \"NameB\"] so the guard can "
                        "validate runtime exports."
                    )
                _mp: Path | None = None
                _stem = Path(_target_module).stem
                if _stem in _module_paths_by_stem:
                    _mp = _module_paths_by_stem[_stem]
                else:
                    for _ext in (".bas", ".cls"):
                        for _cand in vba_source.rglob(_stem + _ext):
                            _mp = _cand
                            break
                        if _mp is not None:
                            break
                if _mp is None:
                    errors.append(
                        f"shared_helper_registry.yaml: concern={_reg_concern!r} "
                        f"variant={_reg_variant!r} declares vba_module="
                        f"{_target_module!r} but no such VBA module file exists "
                        f"under {vba_source}. REGISTRY→VBA GUARD: the SSO contract "
                        "names the implementing module/function; either the YAML "
                        "vba_module value is wrong or the implementation file is missing."
                    )
                    continue
                _mod_raw = _mp.read_bytes()
                for _fn_name in _bare_names:
                    _esc_fn = re.escape(_fn_name.encode("ascii"))
                    _export_re = re.compile(
                        rb'^\s*Public\s+(Function|Sub|Property\s+(Get|Let|Set))\s+'
                        + _esc_fn + rb'\b',
                        re.MULTILINE,
                    )
                    if not _export_re.search(_mod_raw):
                        errors.append(
                            f"{_mp.name}: shared_helper_registry.yaml concern="
                            f"{_reg_concern!r} variant={_reg_variant!r} declares "
                            f"vba_public_fn identifier {_fn_name!r} but this module "
                            "has NO matching Public (Function|Sub|Property) export "
                            "signature. REGISTRY→VBA GUARD: either the YAML "
                            "vba_public_fn value is wrong (try a simpler shape: "
                            "\"NameA + NameB\" or YAML list [\"NameA\", \"NameB\"]) "
                            "or the public export has been renamed/deleted in VBA "
                            "without updating the SSO contract."
                        )

        validated_module_pairs: set[tuple[int, int]] = set()
        for binding in transform_bindings:
            module_path = profile_to_module.get(binding.profile_key)
            if module_path is None:
                continue
            constants_dict = binding.transform_constants_dict or {}
            pair_key = (id(module_path), id(constants_dict))
            if pair_key in validated_module_pairs:
                continue
            validated_module_pairs.add(pair_key)
            raw = module_path.read_bytes()
            worktag_bos = constants_dict.get("worktag_business_objects")
            derived_wts = constants_dict.get("derived_worktags")
            eib_type = str(binding.eib_population_type or "")

            # NON-CANONICAL PIPELINE EXEMPTION.
            # When transform_constants.worktag_resolution.resolution ==
            # "course_name_direct_from_source_lookup_sheet", the transform does NOT use
            # Mapping MASTER / Final Code segmentation / canonical ResolveReportingTypeFromFinalCode.
            # All worktag values come verbatim from a secondary source-lookup sheet in the same
            # source workbook, matched on Course Name (or equivalent driver).  Skip the 3 Stage-6
            # congruence guards that are tightly coupled to the canonical Mapping-MASTER pipeline:
            #   (6b)  AR Private Const WBO_* declaration congruence
            #   (6c)  derived_worktags.reporting_type <-> ResolveReportingTypeFromFinalCode() call
            #   (6e)  Customer-Invoice per_row_tax AR Private Const congruence (when the profile
            #         uses its own tax_triple_per_row naming instead of the legacy per_row_tax)
            worktag_res = constants_dict.get("worktag_resolution") if isinstance(constants_dict, dict) else None
            is_noncanonical_course_lookup = (
                isinstance(worktag_res, dict) and
                str(worktag_res.get("resolution", "")) == "course_name_direct_from_source_lookup_sheet"
            )

            # AP/AR scope dispatch:
            #   modAp* = Tier-1 hoisted scope. WBO + TAX consts live as PUBLIC in modFinanceCommon.
            #   modAr* = legacy scope (AR out of this rollout). Keep original Private Const check.
            is_ap_module = module_path.name.startswith("modAp")
            is_ar_module = module_path.name.startswith("modAr")

            # --- Stage 1 shared_helpers block validation (AP + AR required) ---
            # Shell-variant spec dedupe: spec-existence + shared_helpers declaration are
            # a property of the module (.bas), not each shell-variant binding. Run ONLY
            # once per unique module_path (cache result). E.g. ap_stationery_addison_lee_fa
            # is a shell-variant of the ap_stationery base profile and reuses the SAME
            # ap_stationery.yaml transform_spec (both compile to modApStationerySupplierInvoice.bas).
            if profile_dir is not None and transform_specs_dir is not None:
                spec_shared_helpers: dict[str, object] | None = module_to_spec_shared_helpers.get(module_path)
                spec_path_for_error = transform_specs_dir / f"{binding.profile_key}.yaml"
                if spec_shared_helpers is None:
                    # (i) find the transform_spec file associated with this module.
                    # Spec lookup order:
                    #   1. exact <profile_key>.yaml for the binding we are currently iterating
                    #   2. ANY <other_profile_key>.yaml for any binding that maps to same module_path
                    #   3. existing <spec_name>.yaml in transform_specs_dir whose stem is a
                    #      PREFIX of the current binding's profile_key. This handles the case
                    #      where variants: replace the base binding (base_binding is discarded
                    #      by _expand_shell_variants when variants: is non-empty) but the
                    #      build profile file's stem is the correct shared-spec owner.
                    spec_candidates: list[Path] = [transform_specs_dir / f"{binding.profile_key}.yaml"]
                    for other_key, other_module in profile_to_module.items():
                        if other_module == module_path and other_key != binding.profile_key:
                            spec_candidates.append(transform_specs_dir / f"{other_key}.yaml")
                    try:
                        all_existing_specs = sorted(transform_specs_dir.glob("*.yaml"))
                    except (OSError, ValueError):
                        all_existing_specs = []
                    for existing_spec in all_existing_specs:
                        existing_stem = existing_spec.stem
                        if binding.profile_key == existing_stem:
                            continue
                        if binding.profile_key.startswith(f"{existing_stem}_"):
                            spec_candidates.append(existing_spec)
                    resolved_spec_path: Path | None = None
                    for _cand in spec_candidates:
                        if _cand.exists():
                            resolved_spec_path = _cand
                            break
                    if resolved_spec_path is None:
                        errors.append(
                            f"{module_path.name}: no transform_specs/*.yaml file found. "
                            f"Tried candidates: {[str(p.name) for p in spec_candidates]}. "
                            "STAGE-1 RULE: every AP module needs exactly one transform_specs "
                            "YAML declaring a shared_helpers: block so drift between helper "
                            "variants used in .bas and declared variants is mechanically prevented."
                        )
                        module_to_spec_shared_helpers[module_path] = {}
                    else:
                        spec_path_for_error = resolved_spec_path
                        _spec_text = resolved_spec_path.read_text(encoding="utf-8")
                        with resolved_spec_path.open("r", encoding="utf-8") as _sh:
                            _spec_root = _yaml_lib.safe_load(_sh) or {}
                        # --- Spec<->.bas column-agreement gate (mapping-SOT anti-drift) ---
                        # Every EIB column the spec declares as a mapping target MUST be written by
                        # the .bas via ColumnIndex(). This is the gate that would have caught the
                        # v46.1->v46.2 drift (spec said GF/GQ, code wrote the identically-named split
                        # twins IK/IL). Transforms whose spec expresses mapping differently yield no
                        # targets and are simply not covered here.
                        _spec_targets = _parse_spec_mapping_targets(_spec_text)
                        if _spec_targets:
                            _bas_cols = {
                                mm.group(1).decode("ascii")
                                for mm in column_index_letter_re.finditer(raw)
                            }
                            _spec_only = sorted(_spec_targets - _bas_cols)
                            if _spec_only:
                                errors.append(
                                    f"{module_path.name}: transform_spec '{resolved_spec_path.name}' "
                                    f"declares mapping target column(s) {_spec_only} that the .bas never "
                                    "writes via ColumnIndex(). SPEC<->CODE DRIFT: the spec mapping is the "
                                    "authoring source of truth the .bas mirrors -- either the .bas must "
                                    "write those columns, or the spec target is wrong. (This is the gate "
                                    "that catches the v46.1->v46.2 class where the spec named one column "
                                    "and the code wrote its identically-named twin.)"
                                )
                        if isinstance(_spec_root, dict) and isinstance(_spec_root.get("shared_helpers"), dict):
                            spec_shared_helpers = {
                                str(k): v for k, v in _spec_root["shared_helpers"].items()
                            }
                        else:
                            errors.append(
                                f"{resolved_spec_path.name}: shared_helpers: block is MISSING or non-dict. "
                                f"Module {module_path.name}. "
                                "STAGE-1 RULE: every transform spec (AP and AR) must declare shared_helpers with "
                                "one variant key per concern the runtime invokes. Minimum: any helper referenced "
                                "in BARE_HELPER_TO_REQUIRED_CONCERN must be present (e.g. source_date_normalization "
                                "→ NormalizeSourceDate). AP additionally requires: supplier_resolution, "
                                "company_resolution, accounting_string_derivation, worktag_hierarchy_lookup, "
                                "row_net_resolution, row_trailer_detection. See "
                                "config/shared_helper_registry.yaml for valid concern/variant keys."
                            )
                            spec_shared_helpers = {}
                        module_to_spec_shared_helpers[module_path] = spec_shared_helpers
                # (ii) every declared concern/variant must resolve in the registry catalog:
                for _concern, _variant in (spec_shared_helpers or {}).items():
                    variant_str = str(_variant) if _variant is not None else ""
                    concern_variants = registry_catalog.get(str(_concern))
                    if concern_variants is None:
                        errors.append(
                            f"{spec_path_for_error.name}: shared_helpers declares concern "
                            f"{_concern!r} (variant {variant_str!r}) which is NOT a valid "
                            f"concern in config/shared_helper_registry.yaml. Valid concerns: "
                            f"{sorted(registry_catalog.keys())}."
                        )
                    elif variant_str not in concern_variants:
                        errors.append(
                            f"{spec_path_for_error.name}: shared_helpers.{_concern} = "
                            f"{variant_str!r} is NOT a valid variant for concern "
                            f"{_concern!r}. Valid variants: {sorted(concern_variants.keys())}."
                        )
                # (iii) NEG guard — transform .bas (AP+AR) MUST NOT carry Private Function
                # copies of any hoisted helper (they are public in modSourceTable/modFinanceCommon).
                private_fn_re = re.compile(
                    rb'^\s*Private\s+(?:Function|Sub)\s+((?:NormaliseToWordBoundaries|NormalizeCostCodeSeparators'
                    rb'|NormalizeSourceDate|RowIsTrulyEmpty|RowIsEmptyOrTrailer'
                    rb'|ResolveEffectiveNet'
                    rb'|BuildCustomerNameLookup(?:FromWorkbook)?'
                    rb'|NormalizeClubName'
                    rb'|ApplySqlNameTransforms'
                    rb'|ResolveCustomerIdMultiLevel)\b)',
                    re.MULTILINE,
                )
                local_private_fns = private_fn_re.findall(raw)
                if local_private_fns:
                    local_names = sorted({m.decode("ascii") for m in local_private_fns if m})
                    errors.append(
                        f"{module_path.name}: module declares {len(local_names)} "
                        f"local Private Function copies of hoisted shared helpers: "
                        f"{local_names}. TIER-1 STAGE-1 RULE: these helpers are defined as "
                        "Public in shared engine modules (NormaliseToWordBoundaries, "
                        "NormalizeCostCodeSeparators, NormalizeSourceDate, RowIsTrulyEmpty, "
                        "RowIsEmptyOrTrailer, ResolveEffectiveNet live in modSourceTable.bas; "
                        "supplier resolution + customer_name_resolution helpers "
                        "(BuildCustomerNameLookupFromWorkbook, NormalizeClubName, "
                        "ApplySqlNameTransforms, ResolveCustomerIdMultiLevel) live in "
                        "modFinanceCommon.bas). DELETE these local Private (Function|Sub) "
                        "declarations — VBA resolves bare-name calls to Public scope "
                        "automatically."
                    )
                # (iv) POS guard (forward direction) — for declared variants (not "none"),
                # ensure the bare helper names are referenced at call sites in the .bas.
                # This validates that the shared_helpers declaration matches what the
                # runtime actually invokes.
                for _concern, _variant in (spec_shared_helpers or {}).items():
                    variant_str = str(_variant) if _variant is not None else ""
                    if variant_str == "none":
                        continue
                    hint_bare_names = CONCERN_VARIANT_TO_BARE_HELPER_HINTS.get(
                        (str(_concern), variant_str), ()
                    )
                    for bare_name in hint_bare_names:
                        if not _bare_name_is_referenced(raw, bare_name):
                            errors.append(
                                f"{module_path.name}: shared_helpers.{_concern} = "
                                f"{variant_str!r} declares runtime use of helper {bare_name} "
                                "but the .bas body does NOT reference it as a bare identifier. "
                                "STAGE-1 ANTI-DRIFT RULE: if a shared helper variant is "
                                "declared in the YAML spec, the transform module must actually "
                                "invoke the corresponding public helper by bare-name call "
                                "(or set the variant to 'none' / a correct key for what the "
                                ".bas actually does)."
                            )
                # (v) POS guard (reverse direction) — for any helper in
                # BARE_HELPER_TO_REQUIRED_CONCERN that is referenced by bare name in the
                # .bas body at runtime, the transform spec MUST declare the matching
                # concern with a variant != "none". This is the guard that catches the
                # specific gap we shipped where every AR module called NormalizeSourceDate
                # yet the 5 AR specs didn't have a shared_helpers: block at all.
                for bare_name, required_concern in BARE_HELPER_TO_REQUIRED_CONCERN.items():
                    if not _bare_name_is_referenced(raw, bare_name):
                        continue
                    declared_variant = (spec_shared_helpers or {}).get(required_concern)
                    if declared_variant is None:
                        errors.append(
                            f"{module_path.name}: runtime references shared helper "
                            f"{bare_name!r} (via bare-name call) but the transform spec "
                            f"'{spec_path_for_error.name}' does NOT declare concern "
                            f"{required_concern!r} under 'shared_helpers:'. STAGE-1 "
                            f"CONFIG-FIRST RULE: if a transform module invokes a shared "
                            f"helper, the canonical contract for that helper MUST be "
                            f"declared in the spec before the VBA can be considered "
                            "authoritative. Add 'shared_helpers: "
                            f"{required_concern}: <valid_variant>' to the spec (see "
                            "config/shared_helper_registry.yaml for valid variants), "
                            "OR stop calling the helper if the concern is genuinely unused."
                        )
                    elif str(declared_variant) == "none":
                        errors.append(
                            f"{module_path.name}: runtime references shared helper "
                            f"{bare_name!r} but spec '{spec_path_for_error.name}' declares "
                            f"shared_helpers.{required_concern} = 'none'. STAGE-1 "
                            "CONFIG-FIRST RULE: a variant of 'none' means the transform "
                            "promises not to call the corresponding helper at runtime. "
                            "Either: (a) remove the bare-name call to the helper from the "
                            ".bas, or (b) set the shared_helpers variant in the spec to a "
                            "non-'none' key matching what the .bas actually does."
                        )

            # Collect any Private Const WBO_* in the local module (regardless of scope).
            wbo_consts_in_bas: list[tuple[str, str]] = []
            for const_match in wbo_const_re.finditer(raw):
                try:
                    const_name = const_match.group(1).decode("ascii")
                    const_value = const_match.group(2).decode("ascii")
                except UnicodeDecodeError:
                    const_name = ""
                    const_value = ""
                wbo_consts_in_bas.append((const_name, const_value))

            if worktag_bos is None:
                # --- 6a. Pipeline INACTIVE in YAML ---
                if is_ap_module:
                    # (NEG) AP modules: if YAML is inactive, they should not reference ANY
                    # public WBO_* at call sites, AND must not have local Private Const copies.
                    any_wbo_reference = any(
                        _bare_name_is_referenced(raw, name) for name in public_wbo_consts
                    )
                    if wbo_consts_in_bas:
                        errors.append(
                            f"{module_path.name}: profile '{binding.profile_key}' has NO "
                            "worktag_business_objects in the YAML (pipeline inactive), but this "
                            "AP module STILL declares Private Const WBO_* copies: "
                            f"{wbo_consts_in_bas}. TIER-1 RULE: AP modules hoist WBO strings to "
                            "Public Const in modFinanceCommon.bas and reference bare names. "
                            "Remove the local Private Const copies. If the pipeline is truly "
                            "inactive, also remove the bare-name call-site references to WBO_*."
                        )
                    if any_wbo_reference:
                        errors.append(
                            f"{module_path.name}: profile '{binding.profile_key}' has NO "
                            "worktag_business_objects in YAML (pipeline inactive), yet the AP "
                            "module body still references hoisted WBO_* public constants from "
                            "modFinanceCommon.bas. TIER-1 RULE: if the YAML pipeline is inactive, "
                            "the .bas must not invoke worktag lookups either. Either populate "
                            "the YAML worktag_business_objects block, or remove the WBO_* bare-"
                            "name references in the .bas."
                        )
                else:
                    # AR / legacy: original Private Const negative guard.
                    if wbo_consts_in_bas:
                        errors.append(
                            f"{module_path.name}: profile '{binding.profile_key}' has NO "
                            "worktag_business_objects or fixed_literal_worktags in the YAML "
                            "(pipeline inactive), but the .bas module declares Private Const "
                            f"WBO_* strings: {wbo_consts_in_bas}. CANONICAL RULE: the CODING AGENT "
                            "drives .bas implementation FROM the YAML config only. "
                            "If you want the account-string mapping, fill the COMPLETE pipeline "
                            "companion blocks FIRST in the YAML (worktag_business_objects / "
                            "fixed_literal_worktags / derived_worktags / eib_columns / "
                            "unmapped_handling), THEN have the agent author corresponding Private "
                            "Const WBO_* declarations in the .bas. No WBO strings in the .bas without "
                            "matching YAML keys."
                        )
                # ReportingType is also profile-controlled: no config declaration
                # means the .bas must NOT call ResolveReportingTypeFromFinalCode.
                if reportingtype_call_re.search(raw):
                    errors.append(
                        f"{module_path.name}: profile '{binding.profile_key}' has NO "
                        "derived_worktags.reporting_type in the YAML, yet the .bas calls "
                        "ResolveReportingTypeFromFinalCode(). CANONICAL RULE: the CODING AGENT "
                        "drives .bas implementation FROM config only. Declare "
                        "derived_worktags.reporting_type in the YAML FIRST, THEN "
                        "have the agent call ResolveReportingTypeFromFinalCode() in the .bas "
                        "(or remove the call if reporting type is not wanted)."
                    )
                # Pipeline inactive => no column-coverage validation needed; skip rest.
                continue

            # --- 6b. Pipeline ACTIVE in YAML ---
            yaml_wbo_values: dict[str, str] = {
                _canonicalize_worktag_key(str(k)): str(v) for k, v in worktag_bos.items()
            }
            yaml_keys = set(yaml_wbo_values.keys())

            def _yaml_key_from_wbo_const_name(const_name: str) -> str:
                key_stem = const_name[len("WBO_"):].lower()
                if key_stem == "cost_center":
                    key_stem = "cost_centre"
                return _canonicalize_worktag_key(key_stem)

            if is_ap_module:
                # =========================================================================
                # AP TIER-1 validation (hoisted scope). 3 gates:
                #   (1) NEG: NO local Private Const WBO_* copies exist in the transform .bas
                #   (2) SYNC: Public WBO_* in modFinanceCommon 1:1 match the YAML display names
                #   (3) POS: Every YAML-declared key's corresponding public const is referenced
                #            as a bare name in the transform .bas (not silently dropped)
                # =========================================================================
                # (1) NEG guard: zero local private copies in AP transform .bas
                if wbo_consts_in_bas:
                    local_priv_names = sorted({n for n, _v in wbo_consts_in_bas})
                    errors.append(
                        f"{module_path.name}: profile '{binding.profile_key}' is AP scope and "
                        "declares worktag_business_objects in the YAML, but the .bas STILL has "
                        f"local Private Const WBO_* copies: {local_priv_names}. TIER-1 RULE: "
                        "AP modules hoist WBO strings to Public Const WBO_* in modFinanceCommon.bas. "
                        "DELETE these local Private Const copies from the transform .bas — call "
                        "sites resolve to the shared public constants automatically via VBA "
                        "standard-module scope."
                    )
                # (2) SYNC guard: public consts in modFinanceCommon match YAML values exactly
                bas_public_values: dict[str, str] = {}
                for const_name, const_value in public_wbo_consts.items():
                    key_stem = _yaml_key_from_wbo_const_name(const_name)
                    bas_public_values[key_stem] = const_value
                public_keys = set(bas_public_values.keys())
                missing_public_keys = sorted(yaml_keys - public_keys)
                # NOTE: We intentionally do NOT flag extra_public_keys
                # (shared keys present in modFinanceCommon but not in this
                # profile). The Public Const block in modFinanceCommon is the
                # UNION of all worktags used by ANY AP profile. It is normal
                # (and correct) for profiles like AMEX to not use spend_category.
                # Anti-drift on orphan public constants is covered by a separate
                # cross-profile check: any key in modFinanceCommon that is
                # referenced by ZERO profiles should be removed.
                if missing_public_keys:
                    errors.append(
                        f"modFinanceCommon.bas is MISSING Public Const WBO_* declarations for "
                        f"YAML worktag_business_objects keys {missing_public_keys} required by "
                        f"profile '{binding.profile_key}'. TIER-1 SYNC RULE: every YAML "
                        "worktag_business_objects key must have a corresponding Public Const "
                        "WBO_<UpperCamelName> string in modFinanceCommon.bas matching the "
                        "display-name value. Add the missing Public Const declarations to "
                        "modFinanceCommon.bas."
                    )
                shared_wbo_keys = yaml_keys & public_keys
                for key in sorted(shared_wbo_keys):
                    yaml_val = yaml_wbo_values[key]
                    public_val = bas_public_values[key]
                    if yaml_val != public_val:
                        errors.append(
                            f"WBO display-name VALUE MISMATCH for profile "
                            f"'{binding.profile_key}' worktag_business_objects.{key}: YAML "
                            f"value = {yaml_val!r} but modFinanceCommon.bas Public Const value = "
                            f"{public_val!r}. TIER-1 SYNC RULE: modFinanceCommon.bas is the "
                            "canonical runtime source for all AP modules; its Public Const WBO_* "
                            "strings must 1:1 match every YAML profile that uses them. Update "
                            "modFinanceCommon.bas Public Const to the YAML value exactly (or fix "
                            "the YAML if the canonical value should actually change)."
                        )
                # (3) POS guard: every YAML key's bare name is actually referenced in the .bas body
                for yaml_key in sorted(yaml_keys):
                    # Reverse-map yaml_key -> const name (e.g. cost_centre -> WBO_COST_CENTER)
                    const_for_key = None
                    for const_name, _v in public_wbo_consts.items():
                        if _yaml_key_from_wbo_const_name(const_name) == yaml_key:
                            const_for_key = const_name
                            break
                    if const_for_key is None:
                        continue  # missing-public already reported above
                    if not _bare_name_is_referenced(raw, const_for_key):
                        errors.append(
                            f"{module_path.name}: profile '{binding.profile_key}' declares "
                            f"worktag_business_objects.{yaml_key} in the YAML, yet the AP .bas "
                            f"module body does NOT reference the hoisted Public Const "
                            f"{const_for_key} as a bare identifier. TIER-1 RULE: if the YAML "
                            "declares a worktag key, the transform MUST actually use it at "
                            "runtime. Add the corresponding BuildBusinessObjectHierarchyLookup / "
                            f"ResolveLowestHierarchyLookup call that passes {const_for_key}, OR "
                            "remove the worktag key from the YAML if it's genuinely unused."
                        )
            else:
                # =========================================================================
                # AR / legacy validation: original Private Const checks preserved.
                # NON-CANONICAL COURSE-LOOKUP EXEMPTION: profiles that resolve worktags directly
                # from a source-lookup sheet (e.g. Codes tab) skip this block because the .bas
                # never opens Mapping MASTER, so declaring Private Const WBO_* strings
                # would be dead misleading code.
                # =========================================================================
                if not is_noncanonical_course_lookup:
                    bas_wbo_values: dict[str, str] = {}
                    for const_name, const_value in wbo_consts_in_bas:
                        key_stem = _yaml_key_from_wbo_const_name(const_name)
                        bas_wbo_values[key_stem] = const_value
                    bas_keys = set(bas_wbo_values.keys())
                    extra_bas_keys = sorted(bas_keys - yaml_keys)
                    missing_bas_keys = sorted(yaml_keys - bas_keys)
                    if extra_bas_keys:
                        errors.append(
                            f"{module_path.name}: profile '{binding.profile_key}' .bas declares "
                            f"WBO Private Const keys {extra_bas_keys} that are NOT present in the "
                            "YAML worktag_business_objects block. CANONICAL RULE: the CODING AGENT "
                            "drives .bas implementation FROM the YAML config only. Declare every "
                            "business object name in YAML worktag_business_objects FIRST, THEN have "
                            "the agent author the corresponding Private Const WBO_* in the .bas. "
                            "Never add an ad-hoc Private Const WBO_* in the .bas without first "
                            "declaring the key in the YAML profile."
                        )
                    if missing_bas_keys:
                        errors.append(
                            f"{module_path.name}: profile '{binding.profile_key}' YAML declares "
                            f"worktag_business_objects keys {missing_bas_keys} but the .bas does NOT "
                            "have matching Private Const WBO_* declarations for them. "
                            "CANONICAL RULE: the CODING AGENT drives .bas implementation FROM "
                            "the YAML config. If a key is declared in YAML worktag_business_objects, "
                            "the agent MUST author the corresponding Private Const WBO_* so the "
                            ".bas actually uses the Mapping Master hierarchy for that key."
                        )
                    shared_keys = yaml_keys & bas_keys
                    for key in sorted(shared_keys):
                        yaml_val = yaml_wbo_values[key]
                        bas_val = bas_wbo_values[key]
                        if yaml_val != bas_val:
                            errors.append(
                                f"{module_path.name}: WBO display-name VALUE mismatch: profile "
                                f"worktag_business_objects.{key} value = {yaml_val!r} but .bas "
                                f"Private Const WBO_* value = {bas_val!r}. CANONICAL RULE: the "
                                "YAML config is the SINGLE SOURCE OF TRUTH that the CODING AGENT "
                                "reads when authoring the .bas. Fix the .bas Private Const to match "
                                "the YAML value exactly — the agent should have written it that way."
                            )

            # --- 6c. Reporting Type IFF derived_worktags.reporting_type is in YAML.
            # NON-CANONICAL COURSE-LOOKUP EXEMPTION: for profiles that resolve Reporting Type
            # directly from a source-lookup sheet (Codes tab), the Reporting Type comes from
            # the Codes sheet row, not from Final Code seg2 — so the canonical
            # ResolveReportingTypeFromFinalCode() helper is intentionally never called.
            if derived_wts is not None and "reporting_type" in derived_wts:
                if not is_noncanonical_course_lookup and not reportingtype_call_re.search(raw):
                    errors.append(
                        f"{module_path.name}: profile '{binding.profile_key}' YAML declares "
                        "derived_worktags.reporting_type as ACTIVE but the .bas does NOT call "
                        "ResolveReportingTypeFromFinalCode(). CANONICAL RULE: the CODING AGENT "
                        "drives .bas implementation FROM config. Since the YAML declares the "
                        "Reporting Type derivation, the agent must have added the matching "
                        "ResolveReportingTypeFromFinalCode() call in the .bas — add it."
                    )
            else:
                if reportingtype_call_re.search(raw):
                    errors.append(
                        f"{module_path.name}: profile '{binding.profile_key}' YAML does NOT "
                        "declare derived_worktags.reporting_type, yet the .bas calls "
                        "ResolveReportingTypeFromFinalCode(). CANONICAL RULE: the CODING AGENT "
                        "drives .bas implementation FROM config only. Either add the "
                        "derived_worktags.reporting_type block to the YAML FIRST (so the call "
                        "is justified), OR remove that call from the .bas."
                    )

            # --- 6e. Tax triple bidirectional check (Supplier Invoice + Customer Invoice archetypes).
            # CANONICAL RULE mirroring 6a/6b for worktags: every YAML tax declaration MUST
            # have a corresponding set of Private Const strings in the .bas, and each VALUE
            # string must match EXACTLY, otherwise we have two copies of the truth that can
            # silently drift.  The three archetypes + their mapping:
            #
            #   SUPPLIER INVOICE
            #   (A) uniform_tax
            #         YAML applicability -> .bas Private Const OUTPUT_TAX_APPLICABILITY
            #         YAML code          -> .bas Private Const OUTPUT_TAX_CODE
            #         YAML recoverability-> .bas Private Const OUTPUT_TAX_RECOVERABILITY
            #       NOTE: bas also writes those consts verbatim via ColumnIndex("DS") etc.
            #   (B) per_row_tax {rate_0, rate_20} (AMEX Trains V1/V2, Stationery, Taxi)
            #         YAML rate_0.{a,c,r} -> .bas TAX_APPLICABILITY_ZERO / TAX_CODE_ZERO / TAX_RECOVERABILITY_ZERO
            #         YAML rate_20.{a,c,r} -> .bas TAX_APPLICABILITY_STANDARD / TAX_CODE_STANDARD / TAX_RECOVERABILITY_STANDARD
            #       NOTE: bas dispatches via common helper TaxValueForRate(vatRate>=20, ZERO, STANDARD)
            #             or inline P0/PZ/P20 branch on taxCodeColumn; same const names.
            #   (C) conditional_on_tax_amount {tax_zero_rated, tax_standard_rated} (CFA Distributions)
            #         YAML tax_zero_rated.{a,c,r}     -> TAX_APPLICABILITY_ZERO / TAX_CODE_ZERO / TAX_RECOVERABILITY_ZERO
            #         YAML tax_standard_rated.{a,c,r} -> TAX_APPLICABILITY_STANDARD / TAX_CODE_STANDARD / TAX_RECOVERABILITY_STANDARD
            #       BONUS: bas body MUST contain TaxAmount switch driver (ResolveAmount(taxAmount) > 0#
            #              branch or equivalent pattern so we know driver is correct not just consts.
            #
            #   CUSTOMER INVOICE (archetype-limited — per_row_tax.code_zero_rated + code_standard only):
            #   (D) per_row_tax (AR Invoice Run, fines driver)
            #         YAML per_row_tax.code_zero_rated -> .bas TAX_CODE_ZERO_RATED
            #         YAML per_row_tax.code_standard    -> .bas TAX_CODE_STANDARD
            #         YAML per_row_tax.driver_invoice_type_value_fines -> .bas TAX_DRIVER_FINES
            #
            # Guards: if YAML declares NONE of these, .bas must NOT contain ANY of the 9
            # supplier-side tax consts (OUTPUT_TAX_* / TAX_*_ZERO / TAX_*_STANDARD) as
            # Private Const strings (excepting the archetype-invariant shared
            # OUTPUT_TAX_POINT_DATE_TYPE / OUTPUT_TAX_OPTION which are in every SI module
            # and therefore live in archetype safe-list logic).
            SUPPLIER_SHARED_TAX_POINT_DATE = b"OUTPUT_TAX_POINT_DATE_TYPE"
            SUPPLIER_SHARED_TAX_OPTION = b"OUTPUT_TAX_OPTION"
            uniform_tax = constants_dict.get("uniform_tax") if isinstance(constants_dict, dict) else None
            per_row_tax = constants_dict.get("per_row_tax") if isinstance(constants_dict, dict) else None
            tax_mode = constants_dict.get("tax_mode") if isinstance(constants_dict, dict) else None
            tax_zero_rated = constants_dict.get("tax_zero_rated") if isinstance(constants_dict, dict) else None
            tax_standard_rated = constants_dict.get("tax_standard_rated") if isinstance(constants_dict, dict) else None
            has_supplier_arch = eib_type == "Supplier Invoice"
            has_customer_arch = eib_type == "Customer Invoice"
            yml_declares_any_tax = bool(uniform_tax or per_row_tax or tax_mode == "conditional_on_tax_amount")
            if has_customer_arch and isinstance(per_row_tax, dict):
                cust_keys = {"code_zero_rated", "code_standard", "driver_invoice_type_value_fines"}
                if any(k in per_row_tax for k in cust_keys):
                    yml_declares_any_tax = True

            supplier_const_regex_map: dict[str, re.Pattern[bytes]] = {
                "OUTPUT_TAX_APPLICABILITY": re.compile(rb'^\s*Private\s+Const\s+OUTPUT_TAX_APPLICABILITY\s+As\s+String\s*=\s*"([^"]+)"\s*$', re.MULTILINE),
                "OUTPUT_TAX_CODE": re.compile(rb'^\s*Private\s+Const\s+OUTPUT_TAX_CODE\s+As\s+String\s*=\s*"([^"]+)"\s*$', re.MULTILINE),
                "OUTPUT_TAX_RECOVERABILITY": re.compile(rb'^\s*Private\s+Const\s+OUTPUT_TAX_RECOVERABILITY\s+As\s+String\s*=\s*"([^"]+)"\s*$', re.MULTILINE),
                "TAX_APPLICABILITY_ZERO": re.compile(rb'^\s*Private\s+Const\s+TAX_APPLICABILITY_ZERO\s+As\s+String\s*=\s*"([^"]+)"\s*$', re.MULTILINE),
                "TAX_CODE_ZERO": re.compile(rb'^\s*Private\s+Const\s+TAX_CODE_ZERO\s+As\s+String\s*=\s*"([^"]+)"\s*$', re.MULTILINE),
                "TAX_RECOVERABILITY_ZERO": re.compile(rb'^\s*Private\s+Const\s+TAX_RECOVERABILITY_ZERO\s+As\s+String\s*=\s*"([^"]+)"\s*$', re.MULTILINE),
                "TAX_APPLICABILITY_STANDARD": re.compile(rb'^\s*Private\s+Const\s+TAX_APPLICABILITY_STANDARD\s+As\s+String\s*=\s*"([^"]+)"\s*$', re.MULTILINE),
                "TAX_CODE_STANDARD": re.compile(rb'^\s*Private\s+Const\s+TAX_CODE_STANDARD\s+As\s+String\s*=\s*"([^"]+)"\s*$', re.MULTILINE),
                "TAX_RECOVERABILITY_STANDARD": re.compile(rb'^\s*Private\s+Const\s+TAX_RECOVERABILITY_STANDARD\s+As\s+String\s*=\s*"([^"]+)"\s*$', re.MULTILINE),
            }
            customer_const_regex_map: dict[str, re.Pattern[bytes]] = {
                "TAX_DRIVER_FINES": re.compile(rb'^\s*Private\s+Const\s+TAX_DRIVER_FINES\s+As\s+String\s*=\s*"([^"]+)"\s*$', re.MULTILINE),
                "TAX_CODE_ZERO_RATED": re.compile(rb'^\s*Private\s+Const\s+TAX_CODE_ZERO_RATED\s+As\s+String\s*=\s*"([^"]+)"\s*$', re.MULTILINE),
                "TAX_CODE_STANDARD_AR": re.compile(rb'^\s*Private\s+Const\s+TAX_CODE_STANDARD\s+As\s+String\s*=\s*"([^"]+)"\s*$', re.MULTILINE),
            }

            def _extract_bas_value(regex: re.Pattern[bytes], payload: bytes) -> str:
                m = regex.search(payload)
                if not m:
                    return ""
                try:
                    return m.group(1).decode("ascii")
                except UnicodeDecodeError:
                    return "__NON_ASCII__"

            # Quick: list all supplier consts that appear in bas (for negative guard)
            supplier_consts_found: dict[str, str] = {}
            for k, rgx in supplier_const_regex_map.items():
                v = _extract_bas_value(rgx, raw)
                if v:
                    supplier_consts_found[k] = v
            customer_consts_found: dict[str, str] = {}
            for k, rgx in customer_const_regex_map.items():
                v = _extract_bas_value(rgx, raw)
                if v:
                    customer_consts_found[k] = v

            if not yml_declares_any_tax:
                # Negative guard: .bas has no tax consts at all for this archetype.
                if has_supplier_arch and supplier_consts_found:
                    errors.append(
                        f"{module_path.name}: profile '{binding.profile_key}' has NO tax "
                        "declaration in YAML (uniform_tax / per_row_tax / tax_mode:"
                        " conditional_on_tax_amount are all absent), yet the .bas module "
                        f"still declares Supplier-Invoice tax Private Consts: "
                        f"{sorted(supplier_consts_found.keys())}. CANONICAL RULE: YAML "
                        "is the SINGLE SOURCE OF TRUTH for tax applicability/code/recoverability"
                        " strings. Either declare the matching tax block in the YAML "
                        "(if the module genuinely uses it), OR remove the stray Private "
                        "Const TAX_* / OUTPUT_TAX_* strings from the .bas that the "
                        "profile doesn't declare."
                    )
                # NON-CANONICAL COURSE-LOOKUP EXEMPTION: profiles with their own
                # tax_triple_per_row driver (AR Integration) deliberately avoid the
                # legacy Customer-Invoice per_row_tax shape.  Skip the negative guard
                # because the presence of profile-local Private Const TAX_* strings is
                # intentional and correctly declared via the other tax schema key.
                if (
                    has_customer_arch and customer_consts_found
                    and not is_noncanonical_course_lookup
                ):
                    errors.append(
                        f"{module_path.name}: profile '{binding.profile_key}' has NO "
                        "Customer-Invoice per_row_tax.code_zero_rated / code_standard "
                        "in the YAML, yet the .bas declares AR tax Private Consts: "
                        f"{sorted(customer_consts_found.keys())}. CANONICAL RULE: declare "
                        "the matching per_row_tax block in YAML FIRST, THEN have the "
                        "agent author corresponding Private Const strings in the .bas."
                    )
            else:
                # --- Positive check: Supplier Invoice Archetypes A/B/C ---
                if has_supplier_arch:
                    # Classification of supplier-side tax consts for AP Tier-1 hoisting:
                    #   - HOISTED (6-tuple, archetypes B/C): TAX_APPLICABILITY_ZERO, TAX_CODE_ZERO,
                    #     TAX_RECOVERABILITY_ZERO, TAX_APPLICABILITY_STANDARD, TAX_CODE_STANDARD,
                    #     TAX_RECOVERABILITY_STANDARD — must be Public in modFinanceCommon.
                    #   - LOCAL (3-flat, archetype A): OUTPUT_TAX_APPLICABILITY, OUTPUT_TAX_CODE,
                    #     OUTPUT_TAX_RECOVERABILITY — genuinely per-transform uniform consts; keep
                    #     Private Const in each transform that declares uniform_tax.
                    HOISTED_TAX_KEYS = {
                        "TAX_APPLICABILITY_ZERO", "TAX_CODE_ZERO", "TAX_RECOVERABILITY_ZERO",
                        "TAX_APPLICABILITY_STANDARD", "TAX_CODE_STANDARD", "TAX_RECOVERABILITY_STANDARD",
                    }
                    LOCAL_TAX_KEYS = {
                        "OUTPUT_TAX_APPLICABILITY", "OUTPUT_TAX_CODE", "OUTPUT_TAX_RECOVERABILITY",
                    }
                    def _split_local_hoisted_found(
                        d: dict[str, str],
                    ) -> tuple[dict[str, str], dict[str, str]]:
                        local = {k: v for k, v in d.items() if k in LOCAL_TAX_KEYS}
                        hoisted = {k: v for k, v in d.items() if k in HOISTED_TAX_KEYS}
                        return local, hoisted

                    if isinstance(uniform_tax, dict):
                        # Archetype A: uniform_tax — always LOCAL (3-flat Private Const per module)
                        expected = {
                            "OUTPUT_TAX_APPLICABILITY": str(uniform_tax.get("applicability", "")),
                            "OUTPUT_TAX_CODE":          str(uniform_tax.get("code", "")),
                            "OUTPUT_TAX_RECOVERABILITY": str(uniform_tax.get("recoverability", "")),
                        }
                        local_found, hoisted_found = _split_local_hoisted_found(supplier_consts_found)
                        # Uniform archetype invariant: no hoisted TAX_*_ZERO / TAX_*_STANDARD local copies
                        stray_per_row = sorted(hoisted_found.keys())
                        if stray_per_row:
                            errors.append(
                                f"{module_path.name}: profile '{binding.profile_key}' uses "
                                "uniform_tax (single applicability/code/recoverability for "
                                "all rows), but the .bas also declares per-row tax Private "
                                f"Consts {stray_per_row}. UNIFORM archetype consts live in "
                                "OUTPUT_TAX_APPLICABILITY / OUTPUT_TAX_CODE / "
                                "OUTPUT_TAX_RECOVERABILITY only. Remove the stray TAX_*_ZERO / "
                                "TAX_*_STANDARD Private Consts (they belong to archetype B/C: "
                                "per_row_tax or conditional_on_tax_amount)."
                            )
                        # AP uniform: if AP, OUTPUT_TAX_* still validated as local Private Const
                        # (archetype A is genuinely per-module). Same validation as legacy.
                        for k, yaml_val in expected.items():
                            bas_val = local_found.get(k, "")
                            if not bas_val:
                                errors.append(
                                    f"{module_path.name}: profile '{binding.profile_key}' "
                                    "declares transform_constants.uniform_tax but the .bas "
                                    f"does NOT contain Private Const {k} (expected "
                                    f"value {yaml_val!r}). CANONICAL RULE: the CODING "
                                    "AGENT drives .bas Private Consts FROM the YAML tax "
                                    "block. Add the missing const."
                                )
                            elif yaml_val != bas_val:
                                errors.append(
                                    f"{module_path.name}: uniform_tax VALUE MISMATCH on "
                                    f"{k}: YAML uniform_tax.{k[len('OUTPUT_TAX_'):].lower()} = "
                                    f"{yaml_val!r} but .bas Private Const value = {bas_val!r}. "
                                    "CANONICAL RULE: the YAML config is the SINGLE SOURCE "
                                    "OF TRUTH. Fix the .bas Private Const to match exactly."
                                )

                    if isinstance(per_row_tax, dict) and any(
                        k in per_row_tax for k in ("rate_0", "rate_20")
                    ):
                        # Archetype B: per_row_tax.rate_0 / rate_20 (AMEX / Stationery / Taxi)
                        rate_0 = per_row_tax.get("rate_0")
                        rate_20 = per_row_tax.get("rate_20")
                        expected = {}
                        if isinstance(rate_0, dict):
                            expected.update({
                                "TAX_APPLICABILITY_ZERO":  str(rate_0.get("applicability", "")),
                                "TAX_CODE_ZERO":           str(rate_0.get("code", "")),
                                "TAX_RECOVERABILITY_ZERO": str(rate_0.get("recoverability", "")),
                            })
                        if isinstance(rate_20, dict):
                            expected.update({
                                "TAX_APPLICABILITY_STANDARD":  str(rate_20.get("applicability", "")),
                                "TAX_CODE_STANDARD":           str(rate_20.get("code", "")),
                                "TAX_RECOVERABILITY_STANDARD": str(rate_20.get("recoverability", "")),
                            })
                        _local_found, hoisted_found = _split_local_hoisted_found(supplier_consts_found)
                        if is_ap_module:
                            # =====================================================================
                            # AP archetype B: HOISTED 6-tuple validation — HYBRID SCOPE
                            #   DEFAULT (no local override): shared Public Const in modFinanceCommon
                            #   OPT-IN PER-PROFILE ISOLATION: if transform .bas declares its own
                            #     Private Const of the SAME bare name inside its body, that key
                            #     validates LOCALLY against the .bas Private Const value instead of
                            #     the shared public one. Keys NOT locally overridden fall back to
                            #     the shared public const. This hybrid enables "Isolation over
                            #     DRY" per the project contract: e.g. a single profile like taxi
                            #     can carry finance-validated ref-IDs independently from siblings
                            #     until a second profile also proves the same triple.
                            #   (1) NEG guard REMOVED: local Private Copies are now OPT-IN not BANNED.
                            #   (2) SYNC guard: split LOCAL_OVERRIDDEN keys vs PUBLIC_FALLBACK keys
                            #   (3) POS guard: every expected key's bare name is referenced in .bas
                            # =====================================================================
                            locally_overridden = {k: v for k, v in hoisted_found.items() if k in expected}
                            public_fallback = {k: expected[k] for k in expected if k not in locally_overridden}
                            # (2a) LOCAL scope SYNC: Private Const value == YAML value for overridden keys
                            for k, yaml_val in expected.items():
                                if k in locally_overridden:
                                    local_val = locally_overridden[k]
                                    if yaml_val != local_val:
                                        errors.append(
                                            f"{module_path.name}: per_row_tax LOCAL VALUE MISMATCH "
                                            f"on {k}: YAML value = {yaml_val!r} but .bas Private Const "
                                            f"value = {local_val!r}. SINGLE SOURCE OF TRUTH RULE: the "
                                            "YAML per_row_tax block drives this profile's isolated "
                                            "values. Fix the .bas Private Const to match exactly, or "
                                            "if the shared public value is actually correct, DELETE "
                                            f"the local Private Const {k} copy from the .bas so it "
                                            "falls back to the modFinanceCommon.bas Public Const."
                                        )
                            # (2b) PUBLIC fallback scope SYNC: only for keys NOT locally overridden
                            for k, yaml_val in public_fallback.items():
                                public_val = public_tax_consts.get(k, "")
                                if not public_val:
                                    errors.append(
                                        f"modFinanceCommon.bas is MISSING Public Const {k} "
                                        f"(required by profile '{binding.profile_key}' "
                                        f"per_row_tax.rate_0/rate_20 block; expected YAML value = "
                                        f"{yaml_val!r}). TIER-1 SYNC RULE: the 6-tuple tax strings "
                                        "for archetype B/C are declared either as Public Const in "
                                        "modFinanceCommon.bas OR as an isolated local Private Const "
                                        "copy inside the transform .bas. Add whichever is missing."
                                    )
                                elif yaml_val != public_val:
                                    errors.append(
                                        f"TAX VALUE MISMATCH (fallback public scope) for profile "
                                        f"'{binding.profile_key}' per_row_tax on {k}: YAML value = "
                                        f"{yaml_val!r} but modFinanceCommon.bas Public Const value = "
                                        f"{public_val!r}. TIER-1 SYNC RULE: either update "
                                        "modFinanceCommon.bas to the YAML value, OR isolate this "
                                        f"profile by adding its own Private Const {k} copy in the "
                                        ".bas (it will then be validated under the LOCAL scope rule "
                                        "instead of the shared public fallback)."
                                    )
                            # (3) POS guard: every expected key's bare name is referenced
                            for k in sorted(expected.keys()):
                                if not _bare_name_is_referenced(raw, k):
                                    errors.append(
                                        f"{module_path.name}: profile '{binding.profile_key}' "
                                        f"declares per_row_tax needing 6-tuple string {k}, yet "
                                        "the AP .bas module body does NOT reference the hoisted "
                                        f"Public Const {k} as a bare identifier. TIER-1 RULE: "
                                        "if the YAML declares a tax triple, the transform MUST "
                                        "actually use it at runtime. Add the corresponding "
                                        f"TaxValueForRate() call or If/Else branch passing {k}, "
                                        "OR remove the tax block from the YAML if it's unused."
                                    )
                        else:
                            # AR / legacy archetype B: original Private Const validation
                            for k, yaml_val in expected.items():
                                bas_val = hoisted_found.get(k, "")
                                if not bas_val:
                                    errors.append(
                                        f"{module_path.name}: profile '{binding.profile_key}' "
                                        f"declares transform_constants.per_row_tax but the .bas "
                                        f"does NOT contain Private Const {k} (expected YAML "
                                        f"value {yaml_val!r}). CANONICAL RULE: the CODING "
                                        "AGENT drives .bas Private Consts FROM the YAML "
                                        "per_row_tax block. Add the missing const."
                                    )
                                elif yaml_val != bas_val:
                                    errors.append(
                                        f"{module_path.name}: per_row_tax VALUE MISMATCH on "
                                        f"Private Const {k}: YAML value = {yaml_val!r} but "
                                        f".bas Private Const value = {bas_val!r}. CANONICAL "
                                        "RULE: the YAML config is the SINGLE SOURCE OF TRUTH. "
                                        "Fix the .bas Private Const to match exactly."
                                    )
                        # Archetype invariant: both TAX_APPLICABILITY_ZERO/STANDARD exist
                        # iff bas calls TaxValueForRate() or has inline P0/P20 branch.
                        # (We don't mandate a specific dispatch pattern; we only enforce
                        #  the const values themselves match the YAML triples.)

                    if tax_mode == "conditional_on_tax_amount":
                        # Archetype C: conditional_on_tax_amount.zero_rated / standard_rated (CFA)
                        branch_zero = tax_zero_rated if isinstance(tax_zero_rated, dict) else {}
                        branch_std = tax_standard_rated if isinstance(tax_standard_rated, dict) else {}
                        expected = {
                            "TAX_APPLICABILITY_ZERO":  str(branch_zero.get("applicability", "")),
                            "TAX_CODE_ZERO":           str(branch_zero.get("code", "")),
                            "TAX_RECOVERABILITY_ZERO": str(branch_zero.get("recoverability", "")),
                            "TAX_APPLICABILITY_STANDARD":  str(branch_std.get("applicability", "")),
                            "TAX_CODE_STANDARD":           str(branch_std.get("code", "")),
                            "TAX_RECOVERABILITY_STANDARD": str(branch_std.get("recoverability", "")),
                        }
                        _local_found, hoisted_found = _split_local_hoisted_found(supplier_consts_found)
                        if is_ap_module:
                            # =====================================================================
                            # AP archetype C: HOISTED 6-tuple validation — HYBRID SCOPE
                            #   (mirrors archetype B: default = shared Public, opt-in via local
                            #    Private Const with same bare name isolates per-profile values.)
                            # =====================================================================
                            locally_overridden = {k: v for k, v in hoisted_found.items() if k in expected}
                            public_fallback = {k: expected[k] for k in expected if k not in locally_overridden}
                            # (2a) LOCAL scope SYNC for overridden keys
                            for k, yaml_val in expected.items():
                                if k in locally_overridden:
                                    local_val = locally_overridden[k]
                                    if yaml_val != local_val:
                                        errors.append(
                                            f"{module_path.name}: conditional_on_tax_amount LOCAL "
                                            f"VALUE MISMATCH on {k}: YAML value = {yaml_val!r} but "
                                            f".bas Private Const value = {local_val!r}. SINGLE "
                                            "SOURCE OF TRUTH RULE: fix the .bas Private Const to "
                                            "match the YAML value, or remove the local copy so it "
                                            "falls back to modFinanceCommon.bas Public Const."
                                        )
                            # (2b) PUBLIC fallback scope SYNC for keys NOT locally overridden
                            for k, yaml_val in public_fallback.items():
                                public_val = public_tax_consts.get(k, "")
                                if not public_val:
                                    errors.append(
                                        f"modFinanceCommon.bas is MISSING Public Const {k} "
                                        f"(required by profile '{binding.profile_key}' "
                                        "tax_mode=conditional_on_tax_amount; expected YAML "
                                        f"value = {yaml_val!r}). TIER-1 SYNC RULE: either "
                                        "add the missing Public Const declaration to "
                                        "modFinanceCommon.bas, or isolate this profile with "
                                        "a local Private Const copy in the transform .bas."
                                    )
                                elif yaml_val != public_val:
                                    errors.append(
                                        f"TAX VALUE MISMATCH (fallback public scope) for "
                                        f"profile '{binding.profile_key}' "
                                        f"conditional_on_tax_amount on {k}: YAML value = "
                                        f"{yaml_val!r} but modFinanceCommon.bas Public Const "
                                        f"value = {public_val!r}. TIER-1 SYNC RULE: either "
                                        "update modFinanceCommon.bas to the YAML value, or "
                                        f"isolate this profile by adding its own Private "
                                        f"Const {k} copy in the .bas."
                                    )
                            # (3) POS guard: each expected key's bare name referenced
                            for k in sorted(expected.keys()):
                                if not _bare_name_is_referenced(raw, k):
                                    errors.append(
                                        f"{module_path.name}: profile '{binding.profile_key}' "
                                        f"declares tax_mode=conditional_on_tax_amount needing "
                                        f"6-tuple string {k}, yet the AP .bas module body does "
                                        f"NOT reference the hoisted Public Const {k} as a bare "
                                        "identifier. TIER-1 RULE: add the runtime driver branch "
                                        f"that consumes {k}, or remove the tax block from YAML."
                                    )
                        else:
                            # AR / legacy archetype C: original Private Const validation
                            for k, yaml_val in expected.items():
                                bas_val = hoisted_found.get(k, "")
                                if not bas_val:
                                    errors.append(
                                        f"{module_path.name}: profile '{binding.profile_key}' "
                                        "declares transform_constants.tax_mode="
                                        "'conditional_on_tax_amount' but the .bas is MISSING "
                                        f"Private Const {k} (expected YAML value = {yaml_val!r}). "
                                        "CANONICAL RULE: the CODING AGENT drives .bas Private "
                                        "Consts FROM the YAML tax_zero_rated / tax_standard_rated "
                                        "blocks. Add the missing const."
                                    )
                                elif yaml_val != bas_val:
                                    errors.append(
                                        f"{module_path.name}: conditional_on_tax_amount VALUE "
                                        f"MISMATCH on Private Const {k}: YAML value = "
                                        f"{yaml_val!r} but .bas Private Const value = "
                                        f"{bas_val!r}. CANONICAL RULE: the YAML config is the "
                                        "SINGLE SOURCE OF TRUTH. Fix the .bas Private Const "
                                        "to match exactly."
                                    )
                        # BONUS driver check: the .bas body MUST contain an actual
                        # TaxAmount>0 branch somewhere (ResolveAmount(taxAmount) > 0#
                        # OR taxAmountColumn / taxIsStandard = (ResolveAmount(...) > 0#)
                        # or equivalent. Without this, the YAML archetype is declared but
                        # the module still runs with a uniform branch (common drift bug).
                        driver_regexes = [
                            re.compile(rb'ResolveAmount\s*\([^\n]*taxAmount[^\n]*\)\s*>\s*0#', re.IGNORECASE),
                            re.compile(rb'taxIsStandard\s*=\s*\([^)]*>\s*0#\)', re.IGNORECASE),
                            re.compile(rb'taxAmountColumn\s*[+\-*/]', re.IGNORECASE),
                        ]
                        driver_hit = any(r.search(raw) for r in driver_regexes)
                        # (Also accept an inline reference to taxAmount > 0 inside the
                        #  WriteOutputRow body if column naming differs.)
                        if not driver_hit:
                            errors.append(
                                f"{module_path.name}: profile '{binding.profile_key}' "
                                "declares tax_mode='conditional_on_tax_amount' (archetype "
                                "C: switch on Integration TaxAmount > 0) but the .bas body "
                                "does NOT appear to branch on TaxAmount>0 when writing DS "
                                "(Applicability) / DT (Code) / DY (Recoverability). "
                                "CANONICAL RULE: the CODING AGENT must author a runtime "
                                "driver that matches the YAML archetype. Expected pattern: "
                                "taxIsStandard = (ResolveAmount(taxAmountColumn) > 0#) "
                                "followed by an If/Else branch selecting ZERO vs STANDARD "
                                "triples. Either add the driver switch in the .bas, OR "
                                "correct the YAML archetype (e.g. use uniform_tax / "
                                "per_row_tax if that's what this module actually does)."
                            )
                # --- Positive check: Customer Invoice Archetype D (AR fines switch) ---
                if has_customer_arch and isinstance(per_row_tax, dict):
                    cust_yml_map = {
                        "code_zero_rated": ("TAX_CODE_ZERO_RATED", customer_const_regex_map["TAX_CODE_ZERO_RATED"]),
                        "code_standard":    ("TAX_CODE_STANDARD_AR", customer_const_regex_map["TAX_CODE_STANDARD_AR"]),
                        "driver_invoice_type_value_fines": ("TAX_DRIVER_FINES", customer_const_regex_map["TAX_DRIVER_FINES"]),
                    }
                    for yml_key, (label, rgx) in cust_yml_map.items():
                        if yml_key not in per_row_tax:
                            continue
                        expected_val = str(per_row_tax.get(yml_key) or "")
                        m = rgx.search(raw)
                        if not m:
                            errors.append(
                                f"{module_path.name}: profile '{binding.profile_key}' "
                                f"declares per_row_tax.{yml_key} = {expected_val!r} "
                                f"but the .bas does NOT contain Private Const {label}. "
                                "CANONICAL RULE: the CODING AGENT drives .bas AR tax "
                                "strings FROM the YAML per_row_tax block. Add the "
                                "missing Private Const, or correct the YAML."
                            )
                            continue
                        try:
                            bas_val = m.group(1).decode("ascii")
                        except UnicodeDecodeError:
                            bas_val = "__NON_ASCII__"
                        if expected_val != bas_val:
                            errors.append(
                                f"{module_path.name}: Customer Invoice per_row_tax "
                                f"VALUE MISMATCH on Private Const {label}: YAML "
                                f"per_row_tax.{yml_key} = {expected_val!r} but .bas "
                                f"Private Const value = {bas_val!r}. CANONICAL RULE: "
                                "the YAML config is the SINGLE SOURCE OF TRUTH. Fix the "
                                ".bas Private Const to match exactly."
                            )

            # --- 6d. ColumnIndex(\"XX\") writes must be covered by archetype_eib_columns or safe-list.
            # Safe-shared plumbing columns from config/eib_population_defaults.yaml
            # (safe_shared_columns:), keyed by eib_population_type. Single source of truth.
            archetype_safe = safe_shared_by_pop.get(eib_type, frozenset())
            eib_column_map_keys: set[str] = set()
            if eib_type == "Customer Invoice":
                cust_cols = constants_dict.get("customer_invoice_eib_columns") or {}
                for top_key in ("revenue_category", "billable_project_header", "billable_project_line"):
                    if isinstance(cust_cols.get(top_key), str):
                        eib_column_map_keys.add(str(cust_cols[top_key]))
                worktags_plus = cust_cols.get("worktags_plus_order") or {}
                if isinstance(worktags_plus, dict):
                    for col_letter in worktags_plus.values():
                        if isinstance(col_letter, str):
                            eib_column_map_keys.add(col_letter)
            elif eib_type == "Supplier Invoice":
                supp_cols = constants_dict.get("supplier_eib_columns") or {}
                for col_letter in supp_cols.values():
                    if isinstance(col_letter, str):
                        eib_column_map_keys.add(col_letter)
            for col_match in column_index_letter_re.finditer(raw):
                try:
                    col = col_match.group(1).decode("ascii")
                except UnicodeDecodeError:
                    continue
                if col in archetype_safe:
                    continue
                if col in COLUMN_INDEX_READ_ALLOWLIST:
                    continue
                if col in EIB_CLEAR_END_COL_ALLOWLIST:
                    continue
                if col in eib_column_map_keys:
                    continue
                errors.append(
                    f"{module_path.name}: .bas writes to ColumnIndex(\"{col}\") but "
                    f"profile '{binding.profile_key}' does NOT declare column \"{col}\" in "
                    "its eib_columns YAML block AND it is not in the archetype safe-shared list. "
                    f"CANONICAL RULE: the CODING AGENT drives .bas implementation FROM the YAML "
                    f"config only. EITHER add column '{col}' to the correct eib_columns block for "
                    "this transform in the YAML profile FIRST (so the agent is justified in writing "
                    "to it), OR add it to the archetype SAFE_SHARED_COLUMNS tuple in build_runner.py "
                    "if it is genuinely shared plumbing like Invoice Number or Submit flags. The .bas "
                    "must NEVER write to a column that the config doesn't know about."
                )

    if warnings:
        print("VBA hygiene:")
        for w in warnings:
            print(f"  ! {w}")
    if errors:
        raise RuntimeError(
            "VBA source hygiene failed -- fix before opening Excel:\n  - "
            + "\n  - ".join(errors)
        )


SOURCE_FOLDER_INPUT_LABEL = "Bootstrap Source Folder"


def _write_label_value(workbook_path: Path, label_text: str, value: str) -> None:
    load_workbook = _import_openpyxl()

    workbook = load_workbook(workbook_path, keep_vba=True)
    try:
        if RUNNER_SHEET_NAME not in workbook.sheetnames:
            raise RuntimeError(
                f"Workbook {workbook_path} is missing worksheet named "
                f"{RUNNER_SHEET_NAME!r} (sheetnames={workbook.sheetnames}). "
                "Recreate the shell with scripts/create_runner_shell.py and rebuild."
            )
        worksheet = workbook[RUNNER_SHEET_NAME]
        max_row = max(worksheet.max_row, 1)
        label_cell = None
        for row_index in range(1, max_row + 1):
            cell = worksheet.cell(row=row_index, column=1)
            if str(cell.value or "").strip() == label_text:
                label_cell = cell
                break
        if label_cell is None:
            label_cell = worksheet.cell(row=max_row + 1, column=1)
            label_cell.value = label_text

        worksheet.cell(row=label_cell.row, column=label_cell.column + 1).value = value
        workbook.save(workbook_path)
    finally:
        archive = getattr(workbook, "vba_archive", None)
        if archive is not None:
            archive.close()
        workbook.close()


def _load_reference_paths(reference_path: Path) -> dict:
    """Load the shared reference-data paths (customer/supplier/mapping) that every transform
    inherits unless its profile overrides the key. Single source of truth so an updated Workday
    extract or finance mapping file is changed once, not in every profile.

    Authored with one person's mac path root; a build-time root-path substitution (see
    config/root_path_substitutions.yaml, applied in main() via PathSubstitution.rewrite for the
    mapping_spreadsheet/customer_spreadsheet/supplier_spreadsheet targets) rewrites that root
    for other build hosts. This function itself does no host-specific rewriting."""
    yaml = _import_yaml()
    if not reference_path.is_file():
        raise RuntimeError(
            f"Shared reference-paths file not found: {reference_path}. "
            "Transform profiles inherit customer/supplier/mapping paths from it."
        )
    with reference_path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}
    paths = data.get("reference_paths")
    if not isinstance(paths, dict):
        raise RuntimeError(f"'{reference_path}' must contain a 'reference_paths' mapping.")
    return {key: str(value or "") for key, value in paths.items()}


@dataclass(frozen=True)
class PathSubstitution:
    """The build-time path-rewrite profile selected for this run.

    Build-time only: the VBA runtime never reads root_path_substitutions.yaml (same
    status as config/transforms/*.yaml) — see its header for the full rationale.

    A profile does exactly three things, in this order, to any target field listed
    in apply_to: (1) replace `root_path_substitution` with `new_root_path` (a
    literal string swap; no-op if `root_path_substitution` is empty), then (3)
    normalize every remaining "/" to "\\" if `convert_forward_slash_to_backslash`
    is true. There is no per-field targeting — every field in apply_to gets the
    same treatment.
    """

    profile_name: str
    root_path_substitution: str
    new_root_path: str
    convert_forward_slash_to_backslash: bool
    apply_to: frozenset[str]

    def rewrite(self, target: str, value: str) -> str:
        """Return `value` rewritten if `target` is in this profile's apply_to set."""
        if target not in self.apply_to:
            return value
        result = value
        if self.root_path_substitution:
            result = result.replace(self.root_path_substitution, self.new_root_path)
        if self.convert_forward_slash_to_backslash:
            result = result.replace("/", "\\")
        return result


def _load_path_substitution(substitutions_path: Path, profile_override: str | None) -> PathSubstitution:
    """Resolve the active path-substitution profile from root_path_substitutions.yaml.

    Selection order: --path-profile if given, else os_default_profile for the
    current platform.system(), else the file's fallback_profile.
    """
    yaml = _import_yaml()
    if not substitutions_path.is_file():
        raise RuntimeError(
            f"Path-substitutions config not found: {substitutions_path}. "
            "It selects the build-time filesystem-path rewrites (e.g. the "
            "Windows data-root swap and backslash normalization)."
        )
    with substitutions_path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}

    profiles = data.get("profiles")
    if not isinstance(profiles, dict) or not profiles:
        raise RuntimeError(f"'{substitutions_path}' must contain a non-empty 'profiles' mapping.")

    fallback = str(data.get("fallback_profile") or "none")
    if profile_override:
        profile_name = profile_override
    else:
        os_defaults = data.get("os_default_profile") or {}
        profile_name = str(os_defaults.get(platform.system(), fallback))

    if profile_name not in profiles:
        raise RuntimeError(
            f"Path-substitution profile {profile_name!r} is not defined in "
            f"{substitutions_path} (available: {sorted(profiles)})."
        )

    profile = profiles[profile_name] or {}
    root_path_substitution = str(profile.get("root_path_substitution") or "")
    new_root_path = str(profile.get("new_root_path") or "")
    convert_forward_slash_to_backslash = bool(profile.get("convert_forward_slash_to_backslash", False))

    apply_to = frozenset(str(item) for item in (data.get("apply_to") or []))
    return PathSubstitution(
        profile_name,
        root_path_substitution,
        new_root_path,
        convert_forward_slash_to_backslash,
        apply_to,
    )


VALID_WORKSHEET_RESOLUTIONS = {"named", "first_sheet"}
VALID_HEADER_DETECTIONS = {"literal_row", "scan_required_headers"}

# Supported eib_population_type values. Anything outside this whitelist fails
# immediately at load time — typos like "Supplier_Invoices" or future not-yet-
# implemented archetypes like "Customer Setup" must be explicitly added here
# (plus one entry in eib_population_defaults.yaml eib_populations block, plus
# one archetype branch in _validate_transform_constants_schema). This is the
# single place to extend when Business later brings Setup EIBs on.
SUPPORTED_EIB_POPULATION_TYPES: frozenset[str] = frozenset(
    {"Supplier Invoice", "Customer Invoice"}
)


def _load_eib_defaults(defaults_path: Path) -> dict[str, dict[str, object]]:
    """Load per-EIB-population-type shared conventions.

    Single source of truth for:
      - target_sheet_name  (e.g. "Submit Supplier Invoice")
      - target_data_start_row  (e.g. 6)
      - clear_published_rows_end_col  (e.g. "EX")

    Fail the build immediately if the file is missing, a profile declares a
    population type outside SUPPORTED_EIB_POPULATION_TYPES, or a type isn't
    listed here — that's config drift and must be caught before any module
    import or workbook write step.
    """
    yaml = _import_yaml()
    if not defaults_path.exists():
        raise RuntimeError(
            f"Missing required EIB-population defaults file: {defaults_path}. "
            "Create it with one entry per eib_population_type."
        )
    with defaults_path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}
    pops = data.get("eib_populations")
    if not isinstance(pops, dict):
        raise RuntimeError(
            f"'{defaults_path}' must contain an 'eib_populations' mapping keyed "
            "by eib_population_type string."
        )
    # Explicit whitelist: any key in the YAML file must ALSO be in the code
    # whitelist; and every code-whitelisted type must have a YAML entry. Both
    # directions so a missing YAML entry for a supported type fails fast.
    yaml_types = set(pops.keys())
    missing_from_yaml = sorted(SUPPORTED_EIB_POPULATION_TYPES - yaml_types)
    extra_in_yaml = sorted(yaml_types - SUPPORTED_EIB_POPULATION_TYPES)
    if missing_from_yaml or extra_in_yaml:
        raise RuntimeError(
            "eib_population_defaults.yaml eib_populations keys don't match the "
            "SUPPORTED_EIB_POPULATION_TYPES whitelist in build_runner.py.\n"
            f"  Supported (code): {sorted(SUPPORTED_EIB_POPULATION_TYPES)}\n"
            f"  Defined   (YAML): {sorted(yaml_types)}\n"
            f"  Missing in YAML:  {missing_from_yaml}\n"
            f"  Extra   in YAML:  {extra_in_yaml}\n"
            "To add a new archetype (e.g. Supplier Setup / Customer Setup), add "
            "its name to SUPPORTED_EIB_POPULATION_TYPES + add a YAML block with "
            "the 8 shared keys (target_sheet_name, target_data_start_row, etc.) "
            "+ add one archetype branch in _validate_transform_constants_schema."
        )
    normalized: dict[str, dict[str, object]] = {}
    required_keys_base = (
        "target_sheet_name",
        "target_data_start_row",
        "clear_published_rows_end_col",
    )
    required_keys_by_pop = {
        "Supplier Invoice": required_keys_base + (
            "tax_point_date_type_default",
            "tax_option_default",
            "submit_flag_default",
            "default_currency_default",
            "sequence_start_default",
            "safe_shared_columns",
        ),
        "Customer Invoice": required_keys_base + (
            "auto_complete_default",
            "submit_flag_default",
            "payment_status_default",
            "document_status_default",
            "default_currency_default",
            "safe_shared_columns",
        ),
    }
    for pop_name, entry in pops.items():
        if not isinstance(entry, dict):
            raise RuntimeError(
                f"eib_populations entry '{pop_name}' must be a mapping with keys "
                f"{required_keys_base} (+ archetype extras): {defaults_path}"
            )
        required_keys = required_keys_by_pop.get(str(pop_name), required_keys_base)
        missing = [k for k in required_keys if k not in entry]
        if missing:
            raise RuntimeError(
                f"eib_populations entry '{pop_name}' missing required keys {missing}: {defaults_path}"
            )
        normalized[str(pop_name)] = entry
    return normalized


def _is_valid_required_header_entry(entry: object) -> bool:
    """True when a required_headers list entry is well-formed.

    Accepted forms:
      (a) plain non-empty string with no pipe characters: "Company"
      (b) pipe-delimited alias list: "Decscription|Description"
          - every alternative is a non-empty stripped string
          - at least two distinct alternatives (otherwise just use form (a))

    Pipe-delimited alias lists are the canonical mechanism for tolerating
    source-files whose header name varies (a known typo, a correct spelling,
    a legacy name kept for a prior extract, etc.). Any one alternative matches.
    """
    if not isinstance(entry, str):
        return False
    if not entry:
        return False
    if "|" not in entry:
        return True
    alternatives = [part.strip() for part in entry.split("|")]
    if len(alternatives) < 2:
        return False
    return all(bool(a) for a in alternatives)


def _validate_source_layout_schema(profile_path: Path, source_layout: dict) -> None:
    """Cross-field schema check for source_layout:

    - worksheet_resolution is one of {named, first_sheet}. If 'named' → non-empty
      worksheet_name required. If 'first_sheet' → worksheet_name MUST be omitted.
    - header_detection is one of {literal_row, scan_required_headers}. Either way,
      header_row + data_start_row must be positive integers (scan_required_headers
      uses them as offsets relative to the result of FindHeaderRow).
    - required_headers must be a non-empty list. Each entry is either:
        * a plain non-empty string "HeaderName" (exact-normalized match), OR
        * a pipe-delimited alias list "Alias1|Alias2|..." where any alternative
          counts as a match (for known-typo / renamed-header tolerance).
    """
    if not isinstance(source_layout, dict):
        raise ValueError(f"Missing 'source_layout' mapping in profile: {profile_path}")
    resolution = source_layout.get("worksheet_resolution")
    detection = source_layout.get("header_detection")
    if resolution not in VALID_WORKSHEET_RESOLUTIONS:
        raise ValueError(
            f"Profile '{profile_path}' source_layout.worksheet_resolution must be one of "
            f"{sorted(VALID_WORKSHEET_RESOLUTIONS)}, got: {resolution!r}"
        )
    if detection not in VALID_HEADER_DETECTIONS:
        raise ValueError(
            f"Profile '{profile_path}' source_layout.header_detection must be one of "
            f"{sorted(VALID_HEADER_DETECTIONS)}, got: {detection!r}"
        )
    if resolution == "named":
        ws_name = source_layout.get("worksheet_name")
        if not ws_name or not isinstance(ws_name, str):
            raise ValueError(
                f"Profile '{profile_path}' worksheet_resolution='named' requires a non-empty "
                "source_layout.worksheet_name string."
            )
    elif resolution == "first_sheet":
        if "worksheet_name" in source_layout and source_layout.get("worksheet_name") not in (
            "",
            None,
        ):
            raise ValueError(
                f"Profile '{profile_path}' worksheet_resolution='first_sheet' must NOT set a "
                "source_layout.worksheet_name — leave the key omitted."
            )
    for int_key in ("header_row", "data_start_row"):
        val = source_layout.get(int_key)
        if not isinstance(val, int) or val <= 0:
            raise ValueError(
                f"Profile '{profile_path}' header_detection='{detection}' requires positive "
                f"integer source_layout.{int_key}; got: {val!r}"
            )
    required_headers = source_layout.get("required_headers")
    if (
        not isinstance(required_headers, list)
        or not required_headers
        or any(not _is_valid_required_header_entry(h) for h in required_headers)
    ):
        raise ValueError(
            f"Profile '{profile_path}' must set a non-empty list at "
            "source_layout.required_headers. Each entry may be a plain non-empty "
            "string 'HeaderName' OR a pipe-delimited alias list "
            "'Alias1|Alias2|...' (e.g. 'Decscription|Description' accepts either "
            "the known typo or the correct spelling)."
        )


REQUIRED_WORKTAG_BUSINESS_OBJECTS = (
    "cost_centre",
    "spend_category",
    "event",
    "camp",
    "fa_learning",
    "commercial_partner",
    "grant",
    "programme",
    "project",
    "project_capex",
)
# Canonical key aliases.  `eib_population_defaults.yaml` (88) and all modern
# profiles use British spelling `programme` (since Workday EIB Row5 for GE is
# "Programme").  Historical profiles and earlier .bas files sometimes used the
# American spelling `program`.  The build guardrails accept BOTH but normalize
# to the canonical `programme` for comparisons so the config contract is
# stable: 2026-08-20 v46.1 Customer Invoice EIB formalized the UK spelling.
CANONICAL_WORKTAG_ALIASES = frozenset({("programme", "program")})
ALLOWED_SUPPLIER_WORKTAG_KEYS = frozenset(REQUIRED_WORKTAG_BUSINESS_OBJECTS) | frozenset(
    alias for (canonical, alias) in CANONICAL_WORKTAG_ALIASES
)
"""Supplier Invoice archetype: ALLOWED keys for a Supplier-Invoice
worktag_business_objects block.  IMPORTANT — these are not mandatory.  A
Supplier-Invoice profile may only declare a subset (e.g. AMEX profiles whose
spend_category is a fixed literal in the .bas, resolved via the companion
fixed_literal_worktags block rather than the Mapping Master hierarchy).

Level-1 validator (purely config-level, runs before any .bas is read):
  (a) forbids unknown keys outside this tuple (including the Customer-specific
      "revenue_category");
  (b) forbids a key simultaneously appearing in both
      worktag_business_objects and fixed_literal_worktags;
  (c) respects the cross-archetype blocker (Supplier archetype = NEVER
      contains revenue_category, Customer archetype = NEVER contains
      spend_category).

Level-2 (config ↔ VBA bi-directional) then enforces that every Private Const
WBO_* declared in the .bas has a matching key in the YAML
worktag_business_objects (and vice versa).  CANONICAL RULE: the CODING AGENT
drives .bas implementation FROM the YAML config only — never ad-hoc strings
inside the .bas.  The VBA runtime does NOT parse this file; the guardrails
enforce the authoring contract at build time so the agent cannot stray."""

ALLOWED_SUPPLIER_WORKTAG_KEYS = frozenset(REQUIRED_WORKTAG_BUSINESS_OBJECTS)

# Customer Invoice archetype: uses revenue_category instead of spend_category.
# Any Customer Invoice profile that activates the account-string mapping
# pipeline (by declaring transform_constants.worktag_business_objects or
# fixed_literal_worktags) must use only keys from this allowed-set; partial
# declarations are permitted as long as the omitted keys are covered by
# fixed_literal_worktags.  CANONICAL RULE — the CODING AGENT drives the .bas
# implementation FROM the YAML: every Private Const WBO_*, every
# ResolveReportingTypeFromFinalCode() call, every ColumnIndex("XX") write
# target must be derivable from the YAML.  The VBA runtime never reads the
# YAML; the guardrails enforce the authoring contract at build time so the
# agent cannot author a .bas that contradicts the config declarations.
REQUIRED_CUSTOMER_INVOICE_WORKTAG_BOS = (
    "revenue_category",
    "cost_centre",
    "event",
    "camp",
    "fa_learning",
    "commercial_partner",
    "grant",
    "programme",
    "project",
    "project_capex",
    "reporting_type",
    "sales_item",
)

ALLOWED_CUSTOMER_WORKTAG_KEYS = frozenset(REQUIRED_CUSTOMER_INVOICE_WORKTAG_BOS) | frozenset(
    alias for (canonical, alias) in CANONICAL_WORKTAG_ALIASES
)

# Lookup: alias/alternate spelling -> canonical worktag key.  Used by the
# Level-1 guardrail to normalize profile keys before checking membership in
# the allowed-set, so both `program` and `programme` resolve to canonical
# `programme` and profiles written in either spelling continue to validate.
_CANONICAL_WORKTAG_ALIAS_TO_CANONICAL = {
    alias: canonical for (canonical, alias) in CANONICAL_WORKTAG_ALIASES
}

def _canonicalize_worktag_key(key: str) -> str:
    return _CANONICAL_WORKTAG_ALIAS_TO_CANONICAL.get(key, key)



# Archetype-safe shared columns that every transform writes to via
# ColumnIndex() in the .bas. These are NOT declared in
# worktag_business_objects nor in supplier_eib_columns /
# customer_invoice_eib_columns, because they are generic layout (not driven
# by the account-string mapping pipeline). Used by the
# config-vs-VBA ColumnIndex() coverage guard.
# NOTE: the archetype-safe shared column sets moved OUT of hard-coded frozensets here into
# config/eib_population_defaults.yaml (`safe_shared_columns:` per population) and are loaded via
# _load_safe_shared_columns() inside _validate_vba_source_hygiene. This gives config-purity parity
# with customer_invoice_eib_columns and lets a cold-start agent see the whole EIB column contract
# in YAML. They were discovered by scanning every .bas for ColumnIndex("XX") writes and removing
# the archetype-specific worktag columns (DP..EX for Supplier; KE/GB/MN/MV..NC for Customer).
# ColumnIndex() read-call sites that don't correspond to a write target
# (e.g. reading existing published row values for clear-end-column purposes).
COLUMN_INDEX_READ_ALLOWLIST = frozenset({"MV"})
# Columns already declared in the source_layout/start_row clearing (passed
# to ClearPublishedRows calls) are not required to also appear in the
# eib_columns map; declare them explicitly here so the guard doesn't flag
# them.
EIB_CLEAR_END_COL_ALLOWLIST = frozenset({
    "EX",   # v46.1 guard column (legacy, defunct bas modules still use)
    "HH",   # v46.2 guard column (current: semantic Project write location)
})

VALID_UNMAPPED_HANDLINGS = frozenset({
    "reject_mandatory_worktags",  # AP CFA/Referee/AMEX style: reject rows if
                                  # mandatory worktags (cost_centre, spend_category)
                                  # do not resolve; track unmapped counters.
    "silent_blank",               # AR Bulk/ML style: leave cell blank if a
                                  # derived/WBO does not resolve. No row rejection.
})

# Archetype-level "Reporting Type" derivation is required for ANY profile
# that uses the shared account-string pipeline (both Supplier Invoice and
# Customer Invoice).
REQUIRED_DERIVED_WORKTAG_KEYS = frozenset({"reporting_type"})
REPORTING_TYPE_REQUIRED_KEYS = ("derivation",)


def _validate_transform_constants_schema(
    profile_path: Path,
    eib_population_type: str,
    transform_consts: dict,
    eib_defaults: dict[str, dict[str, object]],
) -> None:
    """Archetype-aware schema validation for the new transform_constants block.

    Supplier Invoice transforms MUST declare:
      - mapping_sheet (non-empty, typically "MAPPING MASTER")
      - company_fallback (non-empty entity name; used when ref_company JSON has no match)
      - ONE of three tax branches:
          (a) uniform_tax (triple applicability/code/recoverability)
          (b) per_row_tax (rate_0 / rate_20 sub-triples for VAT-split transforms)
          (c) tax_mode: conditional_on_tax_amount with tax_zero_rated +
              tax_standard_rated sub-triples (branch on TaxAmount > 0)
      - If worktag_business_objects is present → all 9 standard names

    Customer Invoice transforms MUST declare:
      - fixed_invoice_type (non-empty: "Standard" or "DSC" or type-specific)
      - fixed_currency (typically "GBP" for current tenant)
      - per_row_tax optional (only ar_invoice_run has it for FINES driver split)
    """
    if not isinstance(transform_consts, dict):
        raise ValueError(
            f"Profile '{profile_path}' must contain a 'transform_constants' mapping "
            "(single source of truth for literals previously inferred from .bas source code."
        )

    eib_type = str(eib_population_type)

    if eib_type == "Supplier Invoice":
        mapping_sheet = transform_consts.get("mapping_sheet")
        if not mapping_sheet or not isinstance(mapping_sheet, str) or not mapping_sheet.strip():
            raise ValueError(
                f"Profile '{profile_path}' transform_constants.mapping_sheet must be a "
                "non-empty string (Supplier Invoice transforms resolve "
                "hierarchy strings through Mapping MASTER)."
            )
        company_fallback = transform_consts.get("company_fallback")
        if not isinstance(company_fallback, str):
            raise ValueError(
                f"Profile '{profile_path}' transform_constants.company_fallback "
                "must be a string (use empty string '' explicitly when per-row "
                "ref_company JSON resolution has no override fallback)."
            )
        uniform_tax = transform_consts.get("uniform_tax")
        per_row_tax = transform_consts.get("per_row_tax")
        tax_mode = transform_consts.get("tax_mode")
        conditional_tax = (
            tax_mode is not None and str(tax_mode) == "conditional_on_tax_amount"
        )
        if uniform_tax is None and per_row_tax is None and not conditional_tax:
            raise ValueError(
                f"Profile '{profile_path}' transform_constants must declare either "
                "'uniform_tax' (triple applicability/code/recoverability) or "
                "'per_row_tax' (rate splits) or 'tax_mode: conditional_on_tax_amount' "
                "with tax_zero_rated + tax_standard_rated sub-triples."
            )
        if uniform_tax is not None:
            if not isinstance(uniform_tax, dict):
                raise ValueError(
                    f"Profile '{profile_path}' transform_constants.uniform_tax "
                    "must be a mapping."
                )
            tax_missing = [
                k for k in ("applicability", "code", "recoverability")
                if k not in uniform_tax or not str(uniform_tax[k]).strip()
            ]
            if tax_missing:
                raise ValueError(
                    f"Profile '{profile_path}' transform_constants.uniform_tax "
                    f"missing non-empty sub-keys {tax_missing}."
                )
        if per_row_tax is not None:
            if not isinstance(per_row_tax, dict):
                raise ValueError(
                    f"Profile '{profile_path}' transform_constants.per_row_tax "
                    "must be a mapping."
                )
            for rate_key in ("rate_0", "rate_20"):
                rate_block = per_row_tax.get(rate_key)
                if not isinstance(rate_block, dict):
                    continue
                rate_missing = [
                        k for k in ("applicability", "code", "recoverability")
                        if k not in rate_block or not str(rate_block[k]).strip()
                    ]
                if rate_missing:
                    raise ValueError(
                        f"Profile '{profile_path}' transform_constants.per_row_tax."
                        f"{rate_key} missing non-empty sub-keys {rate_missing}."
                    )
        if conditional_tax:
            zero_block = transform_consts.get("tax_zero_rated")
            std_block = transform_consts.get("tax_standard_rated")
            for block_label, block in (
                ("tax_zero_rated", zero_block),
                ("tax_standard_rated", std_block),
            ):
                if not isinstance(block, dict):
                    raise ValueError(
                        f"Profile '{profile_path}' transform_constants.{block_label} "
                        "must be a mapping when tax_mode=conditional_on_tax_amount."
                    )
                missing = [
                    k for k in ("applicability", "code", "recoverability")
                    if k not in block or not str(block[k]).strip()
                ]
                if missing:
                    raise ValueError(
                        f"Profile '{profile_path}' transform_constants.{block_label} "
                        f"missing non-empty sub-keys {missing}."
                    )
        worktag_bos = transform_consts.get("worktag_business_objects")
        fixed_literal_bos = transform_consts.get("fixed_literal_worktags")
        if worktag_bos is not None or fixed_literal_bos is not None:
            # Pipeline is active if EITHER worktag_bos or fixed_literal_bos exists.
            if worktag_bos is None:
                worktag_bos = {}
            if fixed_literal_bos is None:
                fixed_literal_bos = {}
            if not isinstance(worktag_bos, dict):
                raise ValueError(
                    f"Profile '{profile_path}' transform_constants."
                    "worktag_business_objects must be a mapping when present."
                )
            if not isinstance(fixed_literal_bos, dict):
                raise ValueError(
                    f"Profile '{profile_path}' transform_constants."
                    "fixed_literal_worktags must be a mapping when present."
                )
            # Level 1a: only Supplier-allowed keys are permitted; Customer's
            # "revenue_category" is strictly forbidden in Supplier archetypes.
            raw_keys = set(worktag_bos.keys()) | set(fixed_literal_bos.keys())
            canonical_keys = {_canonicalize_worktag_key(k) for k in raw_keys}
            canonical_allowed = {_canonicalize_worktag_key(k) for k in ALLOWED_SUPPLIER_WORKTAG_KEYS}
            unknown_keys = sorted(canonical_keys - canonical_allowed)
            if unknown_keys:
                raise ValueError(
                    f"Profile '{profile_path}' Supplier archetype pipeline "
                    f"contains disallowed keys {unknown_keys}. Allowed set = "
                    f"{sorted(ALLOWED_SUPPLIER_WORKTAG_KEYS)}. CANONICAL RULE: the "
                    "CODING AGENT drives .bas implementation FROM the YAML config "
                    "only. Every worktag key the agent writes a Private Const or "
                    "fixed literal for must be a member of the allowed key set so "
                    "the Level-2 bi-directional guard can validate it at build time."
                )
            if "revenue_category" in raw_keys:
                raise ValueError(
                    f"Profile '{profile_path}' is Supplier Invoice archetype but "
                    "its pipeline declares 'revenue_category' — Supplier Invoice "
                    "archetypes use 'spend_category' (or fixed_literal_worktags."
                    "spend_category) instead."
                )
            overlap = sorted(set(worktag_bos.keys()) & set(fixed_literal_bos.keys()))
            if overlap:
                raise ValueError(
                    f"Profile '{profile_path}' has keys {overlap} declared in "
                    "BOTH worktag_business_objects AND fixed_literal_worktags. "
                    "A given accounting-string key must be resolved using EXACTLY "
                    "ONE authoring path: either the CODING AGENT emits a MAPPING "
                    "MASTER hierarchy BuildBusinessObjectHierarchyLookup() call and "
                    "a Private Const WBO_* in the .bas (worktag_business_objects "
                    "declaration), OR the agent emits a Private Const OUTPUT_* "
                    "literal (fixed_literal_worktags declaration). Never both. "
                    "Appearing in both would confuse the agent when it authors the "
                    ".bas module — the config must be unambiguous."
                )
            for wbo_key, wbo_value in worktag_bos.items():
                if not isinstance(wbo_value, str) or not wbo_value.strip():
                    raise ValueError(
                        f"Profile '{profile_path}' worktag_business_objects."
                        f"{wbo_key} must be a non-empty string (the Mapping-"
                        "Master Business Object display name)."
                    )
            for fl_key, fl_spec in fixed_literal_bos.items():
                if not isinstance(fl_spec, dict):
                    raise ValueError(
                        f"Profile '{profile_path}' fixed_literal_worktags."
                        f"{fl_key} must be a mapping with 'value' and "
                        "'target_eib_column' sub-fields."
                    )
                missing_fields = [
                    fn for fn in ("value", "target_eib_column")
                    if fn not in fl_spec or not str(fl_spec[fn]).strip()
                ]
                if missing_fields:
                    raise ValueError(
                        f"Profile '{profile_path}' fixed_literal_worktags."
                        f"{fl_key} missing required sub-fields {missing_fields}."
                    )
            # Every declared value in worktag_bos must be a non-empty string.
            _validate_derived_worktags_block(profile_path, transform_consts)
            _validate_unmapped_handling(profile_path, transform_consts)
            _validate_supplier_eib_columns(profile_path, transform_consts)
            _validate_fixed_dates_block(profile_path, transform_consts)
            _validate_simple_string_constants(profile_path, transform_consts)
            _validate_final_code_derivation_block(profile_path, transform_consts)
            _validate_memo_derivation_block(profile_path, transform_consts)
            _validate_fill_down_context_columns(profile_path, transform_consts)

    elif eib_type == "Customer Invoice":
        fixed_invoice_type = transform_consts.get("fixed_invoice_type")
        if not fixed_invoice_type or not isinstance(fixed_invoice_type, str) or not fixed_invoice_type.strip():
            raise ValueError(
                f"Profile '{profile_path}' transform_constants.fixed_invoice_type "
                "must be a non-empty string."
            )
        fixed_currency = transform_consts.get("fixed_currency")
        if not fixed_currency or not isinstance(fixed_currency, str) or not fixed_currency.strip():
            raise ValueError(
                f"Profile '{profile_path}' transform_constants.fixed_currency "
                "must be a non-empty string.")
        per_row_tax = transform_consts.get("per_row_tax")
        if per_row_tax is not None:
            if not isinstance(per_row_tax, dict):
                raise ValueError(
                    f"Profile '{profile_path}' transform_constants.per_row_tax "
                    "must be a mapping when present."
                )
            prt_required = ("driver_invoice_type_value_fines",
                                 "code_zero_rated", "code_standard")
            prt_missing = [
                k for k in prt_required if k not in per_row_tax or not str(per_row_tax[k]).strip()
            ]
            if prt_missing:
                raise ValueError(
                    f"Profile '{profile_path}' transform_constants.per_row_tax "
                    f"missing {prt_missing}."
                )
        customer_wbos = transform_consts.get("worktag_business_objects")
        customer_fixed_literal = transform_consts.get("fixed_literal_worktags")
        if customer_wbos is not None or customer_fixed_literal is not None:
            # Pipeline ACTIVATION: if either block exists.
            if customer_wbos is None:
                customer_wbos = {}
            if customer_fixed_literal is None:
                customer_fixed_literal = {}
            # ========= ACCOUNTING-STRING MAPPING PIPELINE ACTIVATION =========
            # Guardrails (canonical rule): If a Customer Invoice profile declares ANY
            # worktag_business_objects or fixed_literal_worktags the pipeline is ACTIVATED.
            # When activated, ALL keys must be in ALLOWED_CUSTOMER_WORKTAG_KEYS, and the
            # Supplier-specific key "spend_category" is strictly forbidden.
            if not isinstance(customer_wbos, dict):
                raise ValueError(
                    f"Profile '{profile_path}' transform_constants."
                    "worktag_business_objects must be a mapping when present. "
                    "CANONICAL RULE: the CODING AGENT drives .bas implementation "
                    "FROM the YAML only; use ONLY keys from ALLOWED_CUSTOMER_WORKTAG_KEYS."
                )
            if not isinstance(customer_fixed_literal, dict):
                raise ValueError(
                    f"Profile '{profile_path}' transform_constants."
                    "fixed_literal_worktags must be a mapping when present."
                )
            all_cust_keys_raw = set(customer_wbos.keys()) | set(customer_fixed_literal.keys())
            canonical_cust = {_canonicalize_worktag_key(k) for k in all_cust_keys_raw}
            canonical_cust_allowed = {
                _canonicalize_worktag_key(k) for k in ALLOWED_CUSTOMER_WORKTAG_KEYS
            }
            unknown_cust = sorted(canonical_cust - canonical_cust_allowed)
            if unknown_cust:
                raise ValueError(
                    f"Profile '{profile_path}' Customer archetype pipeline "
                    f"contains disallowed keys {unknown_cust}. Allowed set = "
                    f"{sorted(ALLOWED_CUSTOMER_WORKTAG_KEYS)}. CANONICAL RULE: the "
                    "CODING AGENT drives .bas implementation FROM the YAML config "
                    "only. Every worktag key the agent writes a Private Const or "
                    "fixed literal for must be a member of the allowed key set so "
                    "the Step 6b bi-directional guard can validate it against the "
                    ".bas at build time."
                )
            cust_overlap = sorted(set(customer_wbos.keys()) & set(customer_fixed_literal.keys()))
            if cust_overlap:
                raise ValueError(
                    f"Profile '{profile_path}' Customer Invoice pipeline has "
                    f"keys {cust_overlap} in BOTH worktag_business_objects AND "
                    "fixed_literal_worktags. A given accounting-string key must "
                    "use EXACTLY ONE authoring path: either the CODING AGENT "
                    "emits a MAPPING MASTER BuildBusinessObjectHierarchyLookup() "
                    "plus Private Const WBO_* in the .bas (declared here as "
                    "worktag_business_objects), OR the agent emits a Private "
                    "Const OUTPUT_* literal (declared as fixed_literal_worktags). "
                    "Never both. Appearing in both would confuse the agent when "
                    "it authors the .bas module — the config must be unambiguous."
                )
            if "spend_category" in all_cust_keys_raw:
                raise ValueError(
                    f"Profile '{profile_path}' is Customer Invoice archetype but "
                    "its pipeline declares 'spend_category' which is for Supplier "
                    "Invoice only. Customer Invoice archetypes must use "
                    "'revenue_category' (or fixed_literal_worktags.revenue_category) "
                    "instead."
                )
            for wbo_key, wbo_value in customer_wbos.items():
                if not isinstance(wbo_value, str) or not wbo_value.strip():
                    raise ValueError(
                        f"Profile '{profile_path}' Customer Invoice "
                        f"worktag_business_objects.{wbo_key} must be a non-empty "
                        "string (Mapping-Master Business Object display name)."
                    )
            for fl_key, fl_spec in customer_fixed_literal.items():
                if not isinstance(fl_spec, dict):
                    raise ValueError(
                        f"Profile '{profile_path}' fixed_literal_worktags."
                        f"{fl_key} must be a mapping with 'value' and "
                        "'target_eib_column' sub-fields."
                    )
                missing_fields = [
                    fn for fn in ("value", "target_eib_column")
                    if fn not in fl_spec or not str(fl_spec[fn]).strip()
                ]
                if missing_fields:
                    raise ValueError(
                        f"Profile '{profile_path}' Customer Invoice "
                        f"fixed_literal_worktags.{fl_key} missing sub-fields "
                        f"{missing_fields}."
                    )
            _validate_derived_worktags_block(profile_path, transform_consts)
            _validate_unmapped_handling(profile_path, transform_consts)
            _validate_customer_invoice_eib_columns(profile_path, transform_consts)
    else:
        raise ValueError(
            f"Profile '{profile_path}' has unknown eib_population_type "
            f"{eib_type!r}; no transform_constants archetype rules defined."
        )


def _validate_derived_worktags_block(profile_path: Path, transform_consts: dict) -> None:
    """Validate the derived_worktags block.

    When the worktag pipeline is active, derived_worktags MUST exist and
    MUST declare ALL entries in REQUIRED_DERIVED_WORKTAG_KEYS (currently
    only reporting_type, with a derivation subfield).

    CANONICAL RULE: the CODING AGENT reads the YAML derived_worktags
    declarations and authors the .bas module to call the matching
    ResolveReportingTypeFromFinalCode() helper. The VBA runtime never reads
    the YAML; the Level-2 Step 6c guard later enforces that every YAML
    derived_worktags key has its matching .bas call and vice versa."""
    derived_block = transform_consts.get("derived_worktags")
    if derived_block is None:
        raise ValueError(
            f"Profile '{profile_path}' activates the account-string mapping pipeline via "
            "worktag_business_objects or fixed_literal_worktags but is missing the "
            "'derived_worktags' block. CANONICAL RULE: the CODING AGENT drives .bas "
            "implementation FROM the YAML config only; all derived worktags (e.g. "
            "reporting_type) must be DECLARED in the YAML config FIRST so the agent "
            "can author the matching ResolveReportingTypeFromFinalCode() call in the "
            ".bas — never inferred inside the .bas module."
        )
    if not isinstance(derived_block, dict):
        raise ValueError(
            f"Profile '{profile_path}' transform_constants.derived_worktags must be a mapping."
        )
    missing_required = [
        k for k in REQUIRED_DERIVED_WORKTAG_KEYS
        if k not in derived_block or not isinstance(derived_block[k], dict)
    ]
    if missing_required:
        raise ValueError(
            f"Profile '{profile_path}' derived_worktags missing/non-mapping required "
            f"keys {missing_required}. (All derived worktag key specs must be mappings, "
            "not scalar strings — see REQUIRED_DERIVED_WORKTAG_KEYS.)"
        )
    for derived_key in REQUIRED_DERIVED_WORKTAG_KEYS:
        spec = derived_block[derived_key]
        missing_fields = [
            field_name for field_name in REPORTING_TYPE_REQUIRED_KEYS
            if field_name not in spec or not str(spec[field_name]).strip()
        ]
        if missing_fields:
            raise ValueError(
                f"Profile '{profile_path}' derived_worktags.{derived_key} missing "
                f"required non-empty subfields {missing_fields}. Required tuple "
                f"= {REPORTING_TYPE_REQUIRED_KEYS}. This is the "
                "contract for derived worktag calculation. CANONICAL RULE: the "
                "YAML config is the SINGLE SOURCE OF TRUTH the CODING AGENT reads "
                "when it authors the .bas module — so the contract fields must "
                "be present and unambiguous."
            )


def _validate_unmapped_handling(profile_path: Path, transform_consts: dict) -> None:
    """Validate the unmapped_handling enum field.

    Every worktag-pipeline-active profile must declare an explicit
    unmapped_handling value so the CODING AGENT knows which authoring
    pattern to use when writing the .bas: reject_mandatory_worktags (AP
    CFA/Referee/AMEX style) or silent_blank (AR Bulk/ML style). The VBA
    runtime never reads this value — instead the agent authors if/else
    branches based on it, and the Level-2 guard enforces consistency.
    Rejects typos outside VALID_UNMAPPED_HANDLINGS."""
    unmapped = transform_consts.get("unmapped_handling")
    if unmapped is None or not str(unmapped).strip():
        raise ValueError(
            f"Profile '{profile_path}' is missing explicit "
            "transform_constants.unmapped_handling. CANONICAL RULE: the "
            "CODING AGENT drives .bas implementation FROM the YAML config "
            "only. Declare how unresolved worktag values should be handled "
            "BEFORE the agent authors the .bas — because a default would "
            "let the agent guess and silently cause data-quality bugs. "
            f"Allowed values: {sorted(VALID_UNMAPPED_HANDLINGS)}."
        )
    if str(unmapped) not in VALID_UNMAPPED_HANDLINGS:
        raise ValueError(
            f"Profile '{profile_path}' transform_constants.unmapped_handling={unmapped!r} "
            f"is not in the canonical enum VALID_UNMAPPED_HANDLINGS= "
            f"{sorted(VALID_UNMAPPED_HANDLINGS)}. Fix: use one of these exact strings "
            "so the CODING AGENT emits the correct .bas branches — no guessing, no "
            "silent disagreements between the documented contract and the .bas code."
        )


def _validate_supplier_eib_columns(profile_path: Path, transform_consts: dict) -> None:
    """Validate supplier_eib_columns block (Supplier Invoice archetype).

    For every declared-pipeline key (worktag_business_objects +
    fixed_literal_worktags + derived_worktags), the YAML must pin the exact
    Supplier Invoice EIB column letter. This is the contract the CODING
    AGENT reads when it authors the .bas ColumnIndex("XX") writes — the
    agent never decides on its own which EIB column to use for a worktag.

    The VBA runtime NEVER parses this map; the Level-2 Step 6d guard later
    enforces that every ColumnIndex("XX") write in the .bas either matches
    an entry here, or is already declared in the archetype
    SAFE_SHARED_COLUMNS allowlist (common plumbing like F=Submit, G=Inv#,
    etc). Silent-blank bugs occur when the agent writes a worktag to a
    column that wasn't pinned in config — hence the guardrails."""
    eib_cols = transform_consts.get("supplier_eib_columns")
    if eib_cols is None:
        raise ValueError(
            f"Profile '{profile_path}' (Supplier Invoice) activates the worktag pipeline "
            "but has no transform_constants.supplier_eib_columns block. CANONICAL RULE: "
            "the CODING AGENT drives .bas implementation FROM the YAML only. The agent "
            "needs to know EXACTLY which EIB column letter receives each resolved "
            "worktag — the agent must never guess. Declare supplier_eib_columns in the "
            "YAML FIRST."
        )
    if not isinstance(eib_cols, dict):
        raise ValueError(
            f"Profile '{profile_path}' supplier_eib_columns must be a mapping."
        )
    pipeline_declared_keys: set[str] = set()
    wbos = transform_consts.get("worktag_business_objects") or {}
    fixed = transform_consts.get("fixed_literal_worktags") or {}
    derived = transform_consts.get("derived_worktags") or {}
    pipeline_declared_keys.update(wbos.keys())
    pipeline_declared_keys.update(fixed.keys())
    pipeline_declared_keys.update(derived.keys())
    missing_cols = sorted(k for k in pipeline_declared_keys
                          if k not in eib_cols or not str(eib_cols.get(k, "")).strip())
    if missing_cols:
        raise ValueError(
            f"Profile '{profile_path}' supplier_eib_columns does NOT declare a target "
            f"EIB column for the following pipeline-declared keys: {missing_cols}. "
            "CANONICAL RULE: the CODING AGENT reads supplier_eib_columns to know which "
            "ColumnIndex(\"XX\") writes to emit in the .bas. Every worktag key declared "
            "in worktag_business_objects / fixed_literal_worktags / derived_worktags "
            "MUST have a matching supplier_eib_columns entry FIRST so the agent knows "
            "where to write it."
        )
    # Stray detection: any key in supplier_eib_cols whose column letter has
    # no counterpart in the pipeline. This catches stale / drift entries.
    eib_pipeline_keys = {k for k, v in eib_cols.items()
                         if isinstance(v, str) and v.strip()}
    stray_cols = sorted(eib_pipeline_keys - pipeline_declared_keys)
    if stray_cols:
        raise ValueError(
            f"Profile '{profile_path}' supplier_eib_columns declares keys "
            f"{stray_cols} that have no counterpart in worktag_business_objects + "
            "fixed_literal_worktags + derived_worktags. CANONICAL RULE: the CODING "
            "AGENT drives .bas implementation FROM the YAML — declare entries only "
            "for actual pipeline keys; leftover entries confuse the agent and cause "
            "drift. Remove the stale entries."
        )
    # Cross-check: fixed_literal_worktags.{key}.target_eib_column MUST agree
    # with supplier_eib_columns.{key}. The agent cannot author the .bas to
    # write to two different letters for the same key!
    for fl_key, fl_spec in fixed.items():
        declared = str(fl_spec.get("target_eib_column", "")).strip()
        eib_declared = str(eib_cols.get(fl_key, "")).strip()
        if declared != eib_declared:
            raise ValueError(
                f"Profile '{profile_path}' fixed_literal_worktags.{fl_key}."
                f"target_eib_column={declared!r} but supplier_eib_columns."
                f"{fl_key}={eib_declared!r}. CANONICAL RULE: the CODING AGENT "
                "reads both blocks when authoring the .bas, so they MUST agree "
                "exactly — otherwise the agent would write the literal to the "
                "wrong EIB column."
            )


# --- New structured-field validators for Supplier Invoice transform_constants ---
# These close the schema-validation gaps for the complex nested blocks that the
# CODING AGENT reads when authoring .bas modules. Without these, typos like
# `recipt_date`, `padded_lenght`, or a missing `segment_lengths` array would pass
# the build gate cleanly and only manifest as a runtime failure inside Excel VBA.

def _validate_fixed_dates_block(profile_path: Path, transform_consts: dict) -> None:
    """Validate the optional fixed_dates block.

    Two valid date-handling strategies exist for Supplier Invoice transforms:
      (a) Month-constant dates across all rows: declare transform_constants.fixed_dates
          with invoice_date + receipt_date as parseable ISO-ish strings (e.g. ap_cfa).
      (b) Per-row dates from source columns: do NOT declare fixed_dates; the profile
          should list the date-bearing columns in source_layout.required_headers.

    When fixed_dates IS present, both subfields invoice_date and receipt_date MUST
    exist as non-empty strings. A typo like `reciept_date` or a missing subfield is
    caught here rather than later at VBA runtime."""
    fixed_dates = transform_consts.get("fixed_dates")
    if fixed_dates is None:
        # Strategy (b) per-row dates: no fixed_dates block is fine — the source_layout
        # required_headers list carries InvoiceDate / ReceiptDate column names for the
        # CODING AGENT to read. No enforcement here because required_headers lives in
        # the source_layout scope, not transform_constants.
        return
    if not isinstance(fixed_dates, dict):
        raise ValueError(
            f"Profile '{profile_path}' transform_constants.fixed_dates must be a "
            "mapping with 'invoice_date' and 'receipt_date' sub-fields."
        )
    missing = [
        k for k in ("invoice_date", "receipt_date")
        if k not in fixed_dates or not str(fixed_dates[k]).strip()
    ]
    if missing:
        raise ValueError(
            f"Profile '{profile_path}' transform_constants.fixed_dates is missing "
            f"required non-empty sub-fields {missing}. Both 'invoice_date' and "
            "'receipt_date' must be declared when using month-constant dates "
            "(strategy a). If this transform uses per-row source-column dates "
            "(strategy b), remove the entire fixed_dates block instead."
        )


def _validate_simple_string_constants(profile_path: Path, transform_consts: dict) -> None:
    """Validate simple scalar string constants that the CODING AGENT reads.

    Each of these, when present in the YAML, must be a non-empty string so the
    agent doesn't author a .bas module with a blank Private Const value. Any of
    them may be legitimately absent in certain profiles (e.g. AMEX V2 resolves
    company per-row, so fixed_company_code is a documented sentinel; non-AMEX
    profiles don't declare fixed_supplier_reference_id or fixed_spend_category).
    """
    STRING_CONST_KEYS = (
        "batch_id_token",
        "fixed_company_code",
        "fixed_supplier_reference_id",
        "fixed_spend_category",
    )
    for key in STRING_CONST_KEYS:
        value = transform_consts.get(key)
        if value is None:
            continue
        if not isinstance(value, str) or not value.strip():
            raise ValueError(
                f"Profile '{profile_path}' transform_constants.{key} — when "
                "declared — must be a non-empty string. The CODING AGENT reads "
                "this key to author a corresponding Private Const in the .bas; "
                "a blank or non-string value would produce a broken module."
            )


def _validate_final_code_derivation_block(profile_path: Path, transform_consts: dict) -> None:
    """Validate the optional final_code_derivation block (AMEX Hotel / Trains V2).

    This block drives the VBA DeriveFinalCode() helper — a typo in the nested
    structure means the CODING AGENT authors a .bas that either fails to compile
    (bad field reference) or silently produces wrong Final Codes (wrong segment
    lengths / wrong seg2_char_map). Fail fast at build time so the author fixes
    the YAML, not the Excel VBA editor.

    Required top-level subkeys (all):
      padded_length: int
      pad_char: 1-char string
      strip_chars: list of 1-char strings
      segment_lengths: list[int] of exactly 4 elements (Seg1..Seg4 byte counts)
      seg2_first_char_map: dict[str->6-digit accounting code string]
      on_unmapped_seg2_char: non-empty string enum (ErrorStringAndRejectUnlessOverridden)
      error_token: non-empty string (e.g. "Error")
      output_joiner: 1-char string (typically "-")
      optional_v1_helper_headers: mapping with seg1/seg2/seg3/seg4/seg2_format
      optional_override_headers: mapping with seg1/seg2/seg3_aliases/seg4/seg2_format
    """
    block = transform_consts.get("final_code_derivation")
    if block is None:
        return
    if not isinstance(block, dict):
        raise ValueError(
            f"Profile '{profile_path}' transform_constants.final_code_derivation "
            "must be a mapping when present. See schema in docs/CONFIGURATION.md."
        )
    required_top = (
        "padded_length", "pad_char", "strip_chars", "segment_lengths",
        "seg2_first_char_map", "on_unmapped_seg2_char", "error_token",
        "output_joiner", "optional_v1_helper_headers", "optional_override_headers",
    )
    missing_top = [k for k in required_top if k not in block]
    if missing_top:
        raise ValueError(
            f"Profile '{profile_path}' final_code_derivation missing required "
            f"top-level sub-fields {missing_top}. Required set = {required_top}."
        )
    # Type/shapes
    if not isinstance(block["padded_length"], int) or block["padded_length"] <= 0:
        raise ValueError(
            f"Profile '{profile_path}' final_code_derivation.padded_length must "
            "be a positive integer (typically 22 for the Mercell accounting-string "
            "convention)."
        )
    if not isinstance(block["pad_char"], str) or len(block["pad_char"]) != 1:
        raise ValueError(
            f"Profile '{profile_path}' final_code_derivation.pad_char must be "
            "exactly 1 character long (e.g. '0')."
        )
    if not isinstance(block["strip_chars"], list) or not block["strip_chars"]:
        raise ValueError(
            f"Profile '{profile_path}' final_code_derivation.strip_chars must "
            "be a non-empty list of single-character strings to strip from the "
            "raw code before padding (e.g. ['-', ' ', '.'])."
        )
    for idx, sc in enumerate(block["strip_chars"]):
        if not isinstance(sc, str) or len(sc) != 1:
            raise ValueError(
                f"Profile '{profile_path}' final_code_derivation.strip_chars[{idx}] "
                f"={sc!r} must be a single-character string."
            )
    seglens = block["segment_lengths"]
    if not isinstance(seglens, list) or len(seglens) != 4:
        raise ValueError(
            f"Profile '{profile_path}' final_code_derivation.segment_lengths must "
            "be a list of exactly 4 positive integers: [Seg1, Seg2, Seg3, Seg4] "
            "byte counts (e.g. [4, 6, 8, 4] for the standard Mercell convention)."
        )
    for idx, sl in enumerate(seglens):
        if not isinstance(sl, int) or sl <= 0:
            raise ValueError(
                f"Profile '{profile_path}' final_code_derivation.segment_lengths[{idx}] "
                f"={sl!r} must be a positive integer."
            )
    smap = block["seg2_first_char_map"]
    if not isinstance(smap, dict) or not smap:
        raise ValueError(
            f"Profile '{profile_path}' final_code_derivation.seg2_first_char_map "
            "must be a non-empty mapping of single-character Seg2-prefix keys to "
            "6-character final accounting-code Seg2 values."
        )
    for k, v in smap.items():
        if not isinstance(k, str) or len(k) != 1:
            raise ValueError(
                f"Profile '{profile_path}' final_code_derivation.seg2_first_char_map "
                f"key {k!r} must be a single-character string (the Seg2 first-char "
                "prefix from the raw code)."
            )
        if not isinstance(v, str) or len(v) != 6:
            raise ValueError(
                f"Profile '{profile_path}' final_code_derivation.seg2_first_char_map"
                f"[{k!r}] = {v!r} must be exactly 6 characters (the final Mercell "
                "Seg2 accounting code, e.g. '302101')."
            )
    enum_key = "on_unmapped_seg2_char"
    if not isinstance(block[enum_key], str) or not block[enum_key].strip():
        raise ValueError(
            f"Profile '{profile_path}' final_code_derivation.{enum_key} must be a "
            "non-empty enum string (e.g. 'ErrorStringAndRejectUnlessOverridden')."
        )
    for tk in ("error_token", "output_joiner"):
        if not isinstance(block[tk], str) or not block[tk].strip():
            raise ValueError(
                f"Profile '{profile_path}' final_code_derivation.{tk} must be a "
                "non-empty string."
            )
    # optional_v1_helper_headers
    v1 = block["optional_v1_helper_headers"]
    if not isinstance(v1, dict):
        raise ValueError(
            f"Profile '{profile_path}' final_code_derivation.optional_v1_helper_headers "
            "must be a mapping with seg1/seg2/seg3/seg4/seg2_format sub-fields."
        )
    v1_req = ("seg1", "seg2", "seg3", "seg4", "seg2_format")
    v1_miss = [k for k in v1_req if k not in v1 or not str(v1[k]).strip()]
    if v1_miss:
        raise ValueError(
            f"Profile '{profile_path}' final_code_derivation.optional_v1_helper_headers "
            f"missing non-empty sub-fields {v1_miss}. Required = {v1_req}."
        )
    # optional_override_headers
    ov = block["optional_override_headers"]
    if not isinstance(ov, dict):
        raise ValueError(
            f"Profile '{profile_path}' final_code_derivation.optional_override_headers "
            "must be a mapping with seg1/seg2/seg3_aliases/seg4/seg2_format sub-fields."
        )
    ov_req_string = ("seg1", "seg2", "seg4", "seg2_format")
    ov_miss_str = [k for k in ov_req_string if k not in ov or not str(ov[k]).strip()]
    if ov_miss_str:
        raise ValueError(
            f"Profile '{profile_path}' final_code_derivation.optional_override_headers "
            f"missing non-empty string sub-fields {ov_miss_str}. Also check that "
            "seg3_aliases is a non-empty list."
        )
    if "seg3_aliases" not in ov or not isinstance(ov["seg3_aliases"], list) or not ov["seg3_aliases"]:
        raise ValueError(
            f"Profile '{profile_path}' final_code_derivation.optional_override_headers."
            "seg3_aliases must be a non-empty list of header-name strings (the CODING "
            "AGENT authors a VBA fallback chain using each alias in order)."
        )
    for idx, alias in enumerate(ov["seg3_aliases"]):
        if not isinstance(alias, str) or not alias.strip():
            raise ValueError(
                f"Profile '{profile_path}' final_code_derivation.optional_override_headers."
                f"seg3_aliases[{idx}]={alias!r} must be a non-empty header name."
            )


def _validate_memo_derivation_block(profile_path: Path, transform_consts: dict) -> None:
    """Validate the optional memo_derivation block (AMEX Hotel — col EO).

    This block drives how the VBA Memo (column EO) field is composed from source
    columns. A malformed block causes the CODING AGENT to author a .bas that either
    crashes (references a non-existent source step) or produces a blank memo. Fail
    fast here with a clear message rather than in Excel.

    Required subkeys:
      strategy: non-empty string enum (concat_space_join)
      source_steps: mapping of step-name -> step-spec (each step declares how to
                    derive a piece of the memo from source data)
      concat_order: list of step-names (order of concatenation) referencing names
                    actually declared in source_steps
      target_eib_column: 2-letter uppercase EIB column string
    Optional:
      fallback_optional_column: string (when declared, used if entire concat is blank)
    """
    block = transform_consts.get("memo_derivation")
    if block is None:
        return
    if not isinstance(block, dict):
        raise ValueError(
            f"Profile '{profile_path}' transform_constants.memo_derivation must be "
            "a mapping when present. See docs/CONFIGURATION.md for schema."
        )
    required = ("strategy", "source_steps", "concat_order", "target_eib_column")
    missing = [k for k in required if k not in block]
    if missing:
        raise ValueError(
            f"Profile '{profile_path}' memo_derivation missing required top-level "
            f"sub-fields {missing}. Required = {required}."
        )
    if not isinstance(block["strategy"], str) or not block["strategy"].strip():
        raise ValueError(
            f"Profile '{profile_path}' memo_derivation.strategy must be a non-empty "
            "string enum (e.g. 'concat_space_join')."
        )
    steps = block["source_steps"]
    if not isinstance(steps, dict) or not steps:
        raise ValueError(
            f"Profile '{profile_path}' memo_derivation.source_steps must be a "
            "non-empty mapping of step-name -> step-spec. Each step declares how "
            "to derive one piece of the memo field."
        )
    # Every step must declare how to resolve its value (source_column, or
    # blank_header_column_letter_AH-style resolution, or derivation with columns).
    # Don't deep-validate the per-step schema here — the CODING AGENT has more
    # context and the impl_detail_audit catches drift — but ensure each step is
    # at minimum a mapping with at least one resolution key.
    for step_name, spec in steps.items():
        if not isinstance(spec, dict):
            raise ValueError(
                f"Profile '{profile_path}' memo_derivation.source_steps.{step_name} "
                "must be a mapping declaring how to derive that piece of the memo."
            )
        has_resolution = any(
            k in spec for k in ("source_column", "source_resolution", "derivation")
        )
        if not has_resolution:
            raise ValueError(
                f"Profile '{profile_path}' memo_derivation.source_steps.{step_name} "
                "must declare at least one resolution key: 'source_column', "
                "'source_resolution', or 'derivation'."
            )
    concat = block["concat_order"]
    if not isinstance(concat, list) or not concat:
        raise ValueError(
            f"Profile '{profile_path}' memo_derivation.concat_order must be a "
            "non-empty ordered list of source_steps key names — the order in which "
            "pieces are joined together with the strategy separator."
        )
    # Every concat_order name must exist as a key in source_steps
    bad_refs = [name for name in concat if name not in steps]
    if bad_refs:
        # Also accept the convention where a sanitized sub-token (e.g.
        # "traveller_name_sanitized") references a base step "traveller_name" and
        # the strategy suffix is applied by the engine. Allow exact matches or
        # _sanitized suffix on a known step name.
        really_bad = [
            n for n in bad_refs
            if not (isinstance(n, str) and n.endswith("_sanitized") and n[:-len("_sanitized")] in steps)
        ]
        if really_bad:
            raise ValueError(
                f"Profile '{profile_path}' memo_derivation.concat_order references "
                f"step names {really_bad} that are NOT declared in "
                "memo_derivation.source_steps. concat_order must reference names "
                "declared in source_steps (optionally with the engine-applied "
                "'_sanitized' suffix for traveller_name-style steps)."
            )
    eib_col = block["target_eib_column"]
    if (not isinstance(eib_col, str) or not eib_col.isalpha()
            or not eib_col.isupper() or not (1 <= len(eib_col) <= 2)):
        raise ValueError(
            f"Profile '{profile_path}' memo_derivation.target_eib_column={eib_col!r} "
            "must be a 1- or 2-letter uppercase Workday EIB column letter (e.g. "
            "'EO' for the Supplier Invoice Memo column)."
        )
    fb = block.get("fallback_optional_column")
    if fb is not None and (not isinstance(fb, str) or not fb.strip()):
        raise ValueError(
            f"Profile '{profile_path}' memo_derivation.fallback_optional_column — "
            "when declared — must be a non-empty header-name string."
        )


def _validate_fill_down_context_columns(profile_path: Path, transform_consts: dict) -> None:
    """Validate the optional fill_down_context_columns list.

    When a transform declares that certain columns (Departure date, Destination
    airport, etc.) should be filled down from the first non-empty row (AMEX
    Trains V2 pattern), the list must be a non-empty list of non-empty header
    name strings. The canonical location for this key is transform_constants —
    the duplicate under source_layout was removed as a drift-prevention measure.
    """
    fcols = transform_consts.get("fill_down_context_columns")
    if fcols is None:
        return
    if not isinstance(fcols, list) or not fcols:
        raise ValueError(
            f"Profile '{profile_path}' transform_constants.fill_down_context_columns "
            "— when declared — must be a non-empty list of source header name strings "
            "that the VBA engine fills down from the first non-empty row value before "
            "row iteration."
        )
    for idx, col in enumerate(fcols):
        if not isinstance(col, str) or not col.strip():
            raise ValueError(
                f"Profile '{profile_path}' transform_constants.fill_down_context_columns"
                f"[{idx}]={col!r} must be a non-empty header-name string."
            )


def _validate_customer_invoice_eib_columns(profile_path: Path, transform_consts: dict) -> None:
    """Validate customer_invoice_eib_columns block (Customer Invoice archetype).

    Compared with Supplier Invoice (which has one dedicated column per
    worktag), Customer Invoice uses:
      - 1 single-column Revenue Category slot.
      - 2 single-column Billable Project slots (header GB + line-level MN).
      - 1 N-slot Worktags+ repeatable block declared as an ordered mapping
        `worktags_plus_order` listing every remaining worktag.

    This is the contract the CODING AGENT reads when it authors the .bas
    ColumnIndex("XX") writes. The VBA runtime NEVER parses the YAML; the
    Level-2 Step 6d guard later enforces that every ColumnIndex("XX")
    write in the .bas is covered either by this block or by the archetype
    SAFE_SHARED_COLUMNS allowlist (common plumbing like B=Seq, E=etc).

    Guardrails: fail fast if ANY pipeline-declared key
    (worktag_business_objects + fixed_literal_worktags + derived_worktags)
    is not covered by one of the three areas above, to prevent the class
    of silent-blank bugs where the agent writes a worktag to a column
    that isn't pinned in config. Also rejects stray entries so the agent
    isn't confused by stale / drift declarations."""
    eib_cols = transform_consts.get("customer_invoice_eib_columns")
    if eib_cols is None:
        raise ValueError(
            f"Profile '{profile_path}' (Customer Invoice) activates the account-string "
            "mapping pipeline but has no transform_constants.customer_invoice_eib_columns "
            "block. CANONICAL RULE: the CODING AGENT drives .bas implementation FROM "
            "the YAML config only. The agent needs this block to know exactly which "
            "ColumnIndex(\"XX\") writes to emit — it must never guess. Declare it FIRST."
        )
    if not isinstance(eib_cols, dict):
        raise ValueError(
            f"Profile '{profile_path}' customer_invoice_eib_columns must be a mapping."
        )

    required_top_level = ("revenue_category", "billable_project_header",
                          "billable_project_line", "worktags_plus_order")
    missing_top = [k for k in required_top_level
                   if k not in eib_cols or (k == "worktags_plus_order"
                                             and not isinstance(eib_cols[k], dict))
                   or (k != "worktags_plus_order" and not str(eib_cols.get(k, "")).strip())]
    if missing_top:
        raise ValueError(
            f"Profile '{profile_path}' customer_invoice_eib_columns must contain all "
            f"of {required_top_level}; missing/invalid: {missing_top}. "
            "CANONICAL RULE: the CODING AGENT reads the shape of this block to know "
            "which EIB column letter to use for each slot; it needs the 4 top-level "
            "entries (revenue_category, 2 billable_project slots, worktags_plus_order "
            "as an ordered mapping)."
        )

    worktags_plus_order = eib_cols["worktags_plus_order"]

    # Coverage set = pipeline-declared keys (worktag_bos + fixed_literal +
    # derived_worktags) minus the keys with a dedicated single-column slot
    # (revenue_category always single-column; project is single-column via
    # the two billable_project_* slots).
    wbos = transform_consts.get("worktag_business_objects") or {}
    fixed = transform_consts.get("fixed_literal_worktags") or {}
    derived = transform_consts.get("derived_worktags") or {}
    pipeline_keys: set[str] = set(wbos.keys()) | set(fixed.keys()) | set(derived.keys())
    dedicated_single_columns = {"revenue_category", "project"}
    all_wbos_to_cover = pipeline_keys - dedicated_single_columns

    worktags_keys_set = set(worktags_plus_order.keys())

    missing_in_worktags = sorted(all_wbos_to_cover - worktags_keys_set)
    extra_in_worktags = sorted(worktags_keys_set - all_wbos_to_cover)
    if missing_in_worktags:
        raise ValueError(
            f"Profile '{profile_path}' customer_invoice_eib_columns.worktags_plus_order "
            f"does not cover all pipeline-declared Worktags+ keys; missing: "
            f"{missing_in_worktags}. Expected coverage of: "
            f"{sorted(all_wbos_to_cover)}. CANONICAL RULE: the CODING AGENT reads "
            "worktags_plus_order to know which ColumnIndex(\"XX\") slot to emit in the "
            ".bas for every remaining worktag (those without dedicated single columns). "
            "Dedicated single columns ({revenue_category, project}) are correctly "
            "excluded from the sub-block."
        )
    if extra_in_worktags:
        raise ValueError(
            f"Profile '{profile_path}' customer_invoice_eib_columns.worktags_plus_order "
            f"contains names {extra_in_worktags} that are not in the pipeline-declared "
            "set of keys (worktag_business_objects + fixed_literal_worktags + "
            "derived_worktags). CANONICAL RULE: the CODING AGENT drives .bas "
            "implementation FROM the YAML — declare entries only for actual pipeline "
            "keys. Stale entries confuse the agent into writing to wrong slots; remove "
            "them."
        )
    bad_values = [
        (k, v) for k, v in worktags_plus_order.items()
        if not isinstance(v, str) or not v.strip()
    ]
    if bad_values:
        raise ValueError(
            f"Profile '{profile_path}' customer_invoice_eib_columns.worktags_plus_order "
            f"has entries with missing/invalid column letters: {bad_values}. "
            "CANONICAL RULE: the CODING AGENT needs exact EIB column letters (MV / MW / "
            "MX / MY / MZ / NA / NB / NC) to author the .bas ColumnIndex() writes."
        )
    # Cross-check: fixed_literal_worktags.{key}.target_eib_column must agree
    # with the equivalent slot in the eib_columns block. The agent cannot
    # write one literal to two different letters for the same key.
    for fl_key, fl_spec in fixed.items():
        declared = str(fl_spec.get("target_eib_column", "")).strip()
        if fl_key == "revenue_category":
            eib_declared = str(eib_cols.get("revenue_category", "")).strip()
        elif fl_key == "project":
            header_col = str(eib_cols.get("billable_project_header", "")).strip()
            line_col = str(eib_cols.get("billable_project_line", "")).strip()
            if declared not in (header_col, line_col):
                raise ValueError(
                    f"Profile '{profile_path}' fixed_literal_worktags.project."
                    f"target_eib_column={declared!r} but customer_invoice_eib_columns "
                    f"maps project to billable_project_header={header_col!r} and "
                    f"billable_project_line={line_col!r}. CANONICAL RULE: the CODING "
                    "AGENT reads both blocks when authoring the .bas. If project is a "
                    "fixed literal, its column must match one of the two single-column "
                    "slots."
                )
            continue
        else:
            eib_declared = str(worktags_plus_order.get(fl_key, "")).strip()
        if declared != eib_declared:
            raise ValueError(
                f"Profile '{profile_path}' fixed_literal_worktags.{fl_key}."
                f"target_eib_column={declared!r} but customer_invoice_eib_columns "
                f"maps {fl_key} to {eib_declared!r}. CANONICAL RULE: the CODING AGENT "
                "reads both blocks when authoring the .bas; they MUST agree exactly."
            )

def _load_build_manifest(manifest_path: Path) -> dict:
    yaml = _import_yaml()
    if not manifest_path.exists():
        raise FileNotFoundError(
            f"Build manifest not found: {manifest_path}. Expected at elt_transform_excel/config/build_manifest.yaml. "
            "This file is the single source of truth for which transform profiles are active/defunct and the default profile."
        )
    with manifest_path.open("r", encoding="utf-8") as handle:
        manifest = yaml.safe_load(handle) or {}
    for required_key in ("active_profiles", "defunct_profiles", "default_profile_key"):
        if required_key not in manifest:
            raise RuntimeError(
                f"Build manifest {manifest_path} missing required key '{required_key}'. "
                "Expected keys: version, active_profiles (list of key strings), defunct_profiles (list), default_profile_key (string)."
            )
    if not isinstance(manifest["active_profiles"], list):
        raise RuntimeError(f"Build manifest: active_profiles must be a list of profile keys.")
    if not isinstance(manifest["defunct_profiles"], list):
        raise RuntimeError(f"Build manifest: defunct_profiles must be a list of profile keys.")
    if not isinstance(manifest["default_profile_key"], str):
        raise RuntimeError(f"Build manifest: default_profile_key must be a string (matches one entry in active_profiles).")

    active_set: set[str] = set()
    active_keys_ordered: list[str] = []
    for entry in manifest["active_profiles"]:
        if not isinstance(entry, str):
            raise RuntimeError(f"Build manifest: non-string entry in active_profiles: {entry!r}")
        active_set.add(entry)
        active_keys_ordered.append(entry)
    defunct_set: set[str] = set()
    for entry in manifest["defunct_profiles"]:
        if not isinstance(entry, str):
            raise RuntimeError(f"Build manifest: non-string entry in defunct_profiles: {entry!r}")
        defunct_set.add(entry)

    overlap = sorted(active_set & defunct_set)
    if overlap:
        raise RuntimeError(
            f"Build manifest: keys cannot appear in BOTH active_profiles and defunct_profiles: {overlap}"
        )
    if manifest["default_profile_key"] not in active_set:
        raise RuntimeError(
            f"Build manifest: default_profile_key={manifest['default_profile_key']!r} not in active_profiles. "
            f"Active profiles: {sorted(active_set)}"
        )
    manifest["_active_set"] = active_set
    manifest["_active_keys_ordered"] = active_keys_ordered
    manifest["_defunct_set"] = defunct_set
    return manifest


def _load_transform_bindings(profile_dir: Path, manifest_path: Path) -> list[TransformBinding]:
    yaml = _import_yaml()
    manifest = _load_build_manifest(manifest_path)
    active_set: set[str] = manifest["_active_set"]
    active_keys_ordered: list[str] = manifest["_active_keys_ordered"]
    defunct_set: set[str] = manifest["_defunct_set"]

    # Defunct profiles are expected to still exist on disk (preserved as artifacts);
    # warn if any are missing entirely (typ? or forgotten?).
    for defunct_key in sorted(defunct_set):
        profile_path = profile_dir / f"{defunct_key}.yaml"
        if not profile_path.exists():
            print(
                f"[manifest WARN] defunct_profiles entry '{defunct_key}' has NO matching "
                f"{profile_path}. If it was intentionally deleted, remove it from the manifest "
                "defunct_profiles list too."
            )

    # Build profile_paths from active_profiles (ordered exactly as declared, NOT set hash order)
    profile_paths: list[Path] = []
    for active_key in active_keys_ordered:
        profile_path = profile_dir / f"{active_key}.yaml"
        if not profile_path.exists():
            raise FileNotFoundError(
                f"Build manifest declares active profile '{active_key}' but no YAML found at {profile_path}. "
                f"Either create the profile or move '{active_key}' from active_profiles to defunct_profiles."
            )
        profile_paths.append(profile_path)
    if not profile_paths:
        raise RuntimeError(
            f"No active transform profiles declared in manifest {manifest_path}; active_profiles is empty."
        )

    reference_paths = _load_reference_paths(profile_dir.parent / "reference_paths.yaml")
    eib_defaults = _load_eib_defaults(profile_dir.parent / "eib_population_defaults.yaml")

    bindings: list[TransformBinding] = []
    seen_keys: set[str] = set()
    seen_labels: set[str] = set()
    for profile_path in profile_paths:
        with profile_path.open("r", encoding="utf-8") as handle:
            profile = yaml.safe_load(handle) or {}
        base_binding = _parse_transform_binding(
            profile_path, profile, reference_paths, eib_defaults
        )
        expanded = _expand_shell_variants(profile_path, profile, base_binding)
        for binding in expanded:
            normalized_key = binding.profile_key.lower()
            # shell-variant expansions still inherit the manifest's active declaration.
            if normalized_key not in active_set and normalized_key.rstrip("_0123456789") not in active_set:
                # Allow shell variants of base keys; enforce only that base (non-variant) is in manifest.
                pass
            if normalized_key in seen_keys:
                raise RuntimeError(
                    f"Duplicate transform.key '{binding.profile_key}' found in transform profiles."
                )
            seen_keys.add(normalized_key)
            normalized_label = binding.profile_label.lower()
            if normalized_label in seen_labels:
                raise RuntimeError(
                    f"Duplicate transform.label '{binding.profile_label}' found in transform profiles."
                )
            seen_labels.add(normalized_label)
            bindings.append(binding)
    # Present bindings in ALPHABETICAL order by human-readable label (matching the
    # Runner dropdown's alphabetical order) regardless of manifest declaration order.
    bindings.sort(key=lambda b: (b.profile_label.lower(), b.profile_key.lower()))
    return bindings


SHELL_VARIANT_FIELD_MAP: dict[str, str] = {
    "profile_key": "profile_key",
    "profile_label": "profile_label",
    "source_data": "source_data",
    "target_eib": "target_eib",
    "output_file_path": "output_file_path",
    "mapping_spreadsheet": "mapping_spreadsheet",
    "mapping_sheet": "mapping_sheet",
    "customer_spreadsheet": "customer_spreadsheet",
    "supplier_spreadsheet": "supplier_spreadsheet",
    "workday_reference_ids": "workday_reference_ids",
    "source_worksheet_name": "worksheet_name",
    "source_worksheet_resolution": "worksheet_resolution",
    "source_header_row": "header_row",
    "source_data_start_row": "data_start_row",
}

SHELL_VARIANT_REQUIRED_FIELDS = ("profile_key", "profile_label", "source_data", "output_file_path")


def _expand_shell_variants(
    profile_path: Path,
    profile: dict,
    base_binding: TransformBinding,
) -> list[TransformBinding]:
    import dataclasses

    shell_defaults = profile.get("shell_defaults")
    if not isinstance(shell_defaults, dict):
        return [base_binding]
    variants_raw = shell_defaults.get("variants")
    if variants_raw is None:
        return [base_binding]
    if not isinstance(variants_raw, list) or not variants_raw:
        raise ValueError(
            f"Profile '{profile_path}' shell_defaults.variants must be a non-empty list "
            f"when declared. If you want the single base binding, remove the 'variants' key."
        )

    variant_bindings: list[TransformBinding] = []
    for idx, variant in enumerate(variants_raw):
        if not isinstance(variant, dict):
            raise ValueError(
                f"Profile '{profile_path}' shell_defaults.variants[{idx}] must be a mapping. "
                f"Got: {type(variant).__name__}"
            )
        unknown_keys = [k for k in variant.keys() if k not in SHELL_VARIANT_FIELD_MAP]
        if unknown_keys:
            raise ValueError(
                f"Profile '{profile_path}' shell_defaults.variants[{idx}] contains unsupported "
                f"keys: {unknown_keys}. Allowed keys: {sorted(SHELL_VARIANT_FIELD_MAP.keys())}"
            )
        for req_field in SHELL_VARIANT_REQUIRED_FIELDS:
            if req_field not in variant or variant[req_field] in ("", None):
                raise ValueError(
                    f"Profile '{profile_path}' shell_defaults.variants[{idx}] is missing "
                    f"required field '{req_field}'. All variants MUST explicitly provide: "
                    f"{SHELL_VARIANT_REQUIRED_FIELDS}"
                )
        overrides: dict[str, object] = {}
        for variant_key, binding_field in SHELL_VARIANT_FIELD_MAP.items():
            if variant_key in variant:
                overrides[binding_field] = variant[variant_key]
        variant_binding = dataclasses.replace(base_binding, **overrides)
        variant_bindings.append(variant_binding)
    return variant_bindings


def _parse_transform_binding(
    profile_path: Path,
    profile: dict,
    reference_paths: dict | None = None,
    eib_defaults: dict[str, dict[str, object]] | None = None,
) -> TransformBinding:
    reference_paths = reference_paths or {}
    eib_defaults = eib_defaults or {}
    if not isinstance(profile, dict):
        raise ValueError(f"Transform profile must parse to a mapping: {profile_path}")

    transform = profile.get("transform")
    runtime = profile.get("runtime")
    shell_defaults = profile.get("shell_defaults")
    source_layout = profile.get("source_layout")
    if not isinstance(transform, dict):
        raise ValueError(f"Transform profile is missing a 'transform' mapping: {profile_path}")
    if not isinstance(runtime, dict):
        raise ValueError(f"Transform profile is missing a 'runtime' mapping: {profile_path}")
    if not isinstance(shell_defaults, dict):
        raise ValueError(f"Transform profile is missing a 'shell_defaults' mapping: {profile_path}")
    has_variants = isinstance(shell_defaults.get("variants"), list) and bool(shell_defaults["variants"])

    _validate_source_layout_schema(profile_path, source_layout or {})

    eib_population_type_raw = transform.get("eib_population_type")
    if isinstance(eib_population_type_raw, str):
        if eib_population_type_raw not in SUPPORTED_EIB_POPULATION_TYPES:
            raise ValueError(
                f"Profile '{profile_path}' transform.eib_population_type = "
                f"{eib_population_type_raw!r} is not in the archetype whitelist "
                f"SUPPORTED_EIB_POPULATION_TYPES = "
                f"{sorted(SUPPORTED_EIB_POPULATION_TYPES)}.\n"
                "Typos like 'Supplier_Invoices' will be caught here. To add a new archetype "
                "(e.g. Supplier Setup / Customer Setup), add its name to build_runner.py "
                "SUPPORTED_EIB_POPULATION_TYPES + add an eib_population_defaults.yaml entry "
                "+ add one archetype branch in _validate_transform_constants_schema."
            )

    transform_constants = profile.get("transform_constants")
    _validate_transform_constants_schema(
        profile_path,
        eib_population_type_raw,
        transform_constants or {},
        eib_defaults,
    )

    profile_key = _require_profile_value(profile_path, transform, "transform", "key")
    profile_label = _require_profile_value(profile_path, transform, "transform", "label")
    source_type = _require_profile_value(profile_path, transform, "transform", "source_type")
    eib_population_type = _require_profile_value(
        profile_path,
        transform,
        "transform",
        "eib_population_type",
    )
    if eib_population_type not in SUPPORTED_EIB_POPULATION_TYPES:
        raise ValueError(
            f"Profile '{profile_path}' transform.eib_population_type = "
            f"{eib_population_type!r} is not in the archetype whitelist "
            f"{sorted(SUPPORTED_EIB_POPULATION_TYPES)}."
        )
    if eib_population_type not in eib_defaults:
        raise ValueError(
            f"Profile '{profile_path}' transform.eib_population_type = "
            f"{eib_population_type!r} is not listed in config/eib_population_defaults.yaml "
            f"(available: {sorted(eib_defaults)}). Add a new eib_populations entry first."
        )
    step_list_macro = _require_profile_value(
        profile_path,
        runtime,
        "runtime",
        "step_list_macro",
    )

    def _source_layout_value_int(name: str) -> int | None:
        val = source_layout.get(name) if isinstance(source_layout, dict) else None
        if val in ("", None):
            return None
        try:
            int_val = int(val)
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"Profile '{profile_path}' source_layout.{name} must be int, got: {val!r}"
            ) from exc
        if int_val <= 0:
            raise ValueError(
                f"Profile '{profile_path}' source_layout.{name} must be >= 1, got: {int_val!r}"
            )
        return int_val

    def _shell_or_source_str(shell_key: str, source_key: str) -> str | None:
        """Resolve a shell-defaults string with fallback to the source_layout schema block.

        For consistency with every other Runner-populated default (source_data, target_eib,
        mapping_sheet, ...), the worksheet-name / resolution / header-row / data-start-row
        values are preferentially read from shell_defaults when explicitly provided, and
        fall back to source_layout (the schema declaration) otherwise. This mirrors the
        shell_defaults + reference_paths layering pattern used for mapping_spreadsheet etc.
        """
        shell_val = _optional_profile_value(shell_defaults, shell_key)
        if shell_val is not None:
            return shell_val
        if isinstance(source_layout, dict):
            return _optional_profile_value(source_layout, source_key)
        return None

    def _shell_or_source_int(shell_key: str, source_key: str) -> int | None:
        shell_raw = shell_defaults.get(shell_key)
        if shell_raw not in ("", None):
            try:
                shell_int = int(shell_raw)
            except (TypeError, ValueError) as exc:
                raise ValueError(
                    f"Profile '{profile_path}' shell_defaults.{shell_key} must be int, got: {shell_raw!r}"
                ) from exc
            if shell_int <= 0:
                raise ValueError(
                    f"Profile '{profile_path}' shell_defaults.{shell_key} must be >= 1, got: {shell_int!r}"
                )
            return shell_int
        return _source_layout_value_int(source_key)

    required_headers_raw = source_layout.get("required_headers") or []
    required_headers: list[str] = [str(h) for h in required_headers_raw]

    resolved_worksheet_name = _shell_or_source_str("source_worksheet_name", "worksheet_name")
    resolved_worksheet_resolution = (
        _optional_profile_value(shell_defaults, "source_worksheet_resolution")
        if _optional_profile_value(shell_defaults, "source_worksheet_resolution") is not None
        else (
            source_layout.get("worksheet_resolution")
            if isinstance(source_layout, dict)
            else None
        )
    )
    resolved_header_row = _shell_or_source_int("source_header_row", "header_row")
    resolved_data_start_row = _shell_or_source_int("source_data_start_row", "data_start_row")

    if has_variants:
        source_data_val = _optional_profile_value(shell_defaults, "source_data") or ""
        output_val = _optional_profile_value(shell_defaults, "output_file_path") or ""
    else:
        source_data_val = _require_profile_value(
            profile_path, shell_defaults, "shell_defaults", "source_data"
        )
        output_val = _require_profile_value(
            profile_path, shell_defaults, "shell_defaults", "output_file_path"
        )
    return TransformBinding(
        profile_path=profile_path,
        profile_key=profile_key,
        profile_label=profile_label,
        source_type=source_type,
        source_data=source_data_val,
        target_eib=_require_profile_value(
            profile_path, shell_defaults, "shell_defaults", "target_eib"
        ),
        output_file_path=output_val,
        mapping_spreadsheet=_optional_profile_value(shell_defaults, "mapping_spreadsheet")
        or reference_paths.get("mapping_spreadsheet", ""),
        mapping_sheet=_optional_profile_value(shell_defaults, "mapping_sheet")
        or reference_paths.get("mapping_sheet", ""),
        customer_spreadsheet=_optional_profile_value(shell_defaults, "customer_spreadsheet")
        or reference_paths.get("customer_spreadsheet", ""),
        supplier_spreadsheet=_optional_profile_value(shell_defaults, "supplier_spreadsheet")
        or reference_paths.get("supplier_spreadsheet", ""),
        workday_reference_ids=_optional_profile_value(shell_defaults, "workday_reference_ids")
        or reference_paths.get("workday_reference_ids", ""),
        eib_population_type=eib_population_type,
        step_list_macro=step_list_macro,
        start_row=_optional_profile_int(profile_path, transform, "transform", "start_row"),
        worksheet_name=resolved_worksheet_name,
        worksheet_resolution=resolved_worksheet_resolution,
        header_row=resolved_header_row,
        data_start_row=resolved_data_start_row,
        header_detection=source_layout.get("header_detection")
        if isinstance(source_layout, dict)
        else None,
        required_headers=required_headers,
        transform_constants_dict=transform_constants if isinstance(transform_constants, dict) else None,
    )


def _optional_profile_value(section_data: dict, key: str) -> str | None:
    value = section_data.get(key)
    if value in ("", None):
        return None
    return str(value)


def _require_profile_value(
    profile_path: Path,
    section_data: dict,
    section_name: str,
    key: str,
) -> str:
    value = section_data.get(key, "")
    if value in ("", None):
        raise ValueError(f"Profile '{section_name}.{key}' must be populated: {profile_path}")
    return str(value)


def _optional_profile_int(
    profile_path: Path,
    section_data: dict,
    section_name: str,
    key: str,
) -> int | None:
    value = section_data.get(key)
    if value in ("", None):
        return None
    try:
        parsed_value = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            f"Profile '{section_name}.{key}' must be an integer when populated: {profile_path}"
        ) from exc
    if parsed_value < 1:
        raise ValueError(
            f"Profile '{section_name}.{key}' must be >= 1 when populated: {profile_path}"
        )
    return parsed_value


_ALWAYS_PRESERVED_MODULE_NAMES = {"bootstrapimporter", "thisworkbook"}
_DOCUMENT_MODULE_NAME_PATTERN = re.compile(r"^sheet\d+$")


def _preserve_existing_module(module: object) -> bool:
    # Document modules (ThisWorkbook, and every Sheet<N> codename — one per
    # worksheet, including preserved hand-authored sheets like arinvoice/mapping,
    # not just the generated Runner sheet) and BootstrapImporter (which never
    # re-imports itself) are excluded by name, not by sniffing the saved source:
    # Excel adds "Attribute VB_Base = ..." to every class module it re-serializes,
    # not just document modules, so that text alone does not distinguish them.
    # `module.kind` only separates "standard" from "other" (class/document/designer
    # combined), which is too coarse to use here without also hiding a future real
    # class module.
    name = str(getattr(module, "name", "")).lower()
    return name in _ALWAYS_PRESERVED_MODULE_NAMES or bool(
        _DOCUMENT_MODULE_NAME_PATTERN.match(name)
    )


def _mapping_rows_from_lookups() -> list[tuple[str, str, str]]:
    """Compose the Runner 'mapping' sheet rows from the ingest_lookups ref tables.

    Reuses create_ingest_template.load_lookup so mapping data is loaded and validated
    exactly like every other reference dataset — the same mechanism, not a parallel one."""
    from create_ingest_template import load_lookup

    rows: list[tuple[str, str, str]] = []
    for entity, lookup_name, source_key, target_key in MAPPING_LOOKUP_SOURCES:
        for row in load_lookup(lookup_name)["rows"]:
            rows.append((entity, str(row[source_key]), str(row[target_key])))
    return rows


def _write_mapping_sheet(worksheet, mapping_rows: list[tuple[str, str, str]]) -> None:
    max_row = max(worksheet.max_row, len(mapping_rows) + 5)
    for row_index in range(1, max_row + 1):
        for column_index in range(1, len(MAPPING_HEADER) + 1):
            worksheet.cell(row=row_index, column=column_index).value = None

    for column_index, header in enumerate(MAPPING_HEADER, start=1):
        worksheet.cell(row=1, column=column_index).value = header
    for row_offset, mapping_row in enumerate(mapping_rows, start=2):
        for column_index, value in enumerate(mapping_row, start=1):
            worksheet.cell(row=row_offset, column=column_index).value = value


def _write_mapping_sheet_block(workbook_path: Path) -> None:
    """Regenerate the Runner's 'mapping' sheet from the ref tables. The sheet content is a
    build output now, not hand-authored data preserved inside the shell workbook."""
    load_workbook = _import_openpyxl()

    mapping_rows = _mapping_rows_from_lookups()
    workbook = load_workbook(workbook_path, keep_vba=True)
    try:
        if MAPPING_SHEET_NAME in workbook.sheetnames:
            worksheet = workbook[MAPPING_SHEET_NAME]
        else:
            worksheet = workbook.create_sheet(MAPPING_SHEET_NAME)
        _write_mapping_sheet(worksheet, mapping_rows)
        workbook.save(workbook_path)
    finally:
        archive = getattr(workbook, "vba_archive", None)
        if archive is not None:
            archive.close()
        workbook.close()


def _clear_runner_tab_technical_columns(workbook_path: Path) -> None:
    """Remove any residual Transform Registry / tech rows from the Runner tab's
    columns G onwards so Finance only ever sees the simple A-F user panel.

    Historically the Transform Registry lived on the Runner tab starting at column
    G (row 2 title, row 3 header, rows 4+ data).  The registry now lives on its own
    very-hidden 'registry' sheet, so any stale cells here are removed on every
    build.  Bootstrap status labels (Bootstrap Import Source/Count/At, Run Button,
    Bootstrap Import Current, Bootstrap Source Folder) stay on the Runner tab in
    columns A/B below the user panel -- they are user-visible audit/setup metadata.
    """
    load_workbook = _import_openpyxl()

    workbook = load_workbook(workbook_path, keep_vba=True)
    try:
        if RUNNER_SHEET_NAME not in workbook.sheetnames:
            raise RuntimeError(
                f"Workbook {workbook_path} is missing worksheet named "
                f"{RUNNER_SHEET_NAME!r} (sheetnames={workbook.sheetnames}). "
                "Recreate the shell with scripts/create_runner_shell.py and rebuild."
            )
        worksheet = workbook[RUNNER_SHEET_NAME]
        clear_start_col = RUNNER_MAX_USER_COLUMN + 1
        max_row = max(worksheet.max_row, REGISTRY_TITLE_ROW + 50)
        max_col = max(worksheet.max_column, REGISTRY_END_COLUMN + 20)
        for row_index in range(1, max_row + 1):
            for column_index in range(clear_start_col, max_col + 1):
                worksheet.cell(row=row_index, column=column_index).value = None
        workbook.save(workbook_path)
    finally:
        archive = getattr(workbook, "vba_archive", None)
        if archive is not None:
            archive.close()
        workbook.close()


def _ensure_worksheet(workbook, sheet_name: str):
    """Return workbook[sheet_name], creating it first if the sheet is missing."""
    if sheet_name in workbook.sheetnames:
        return workbook[sheet_name]
    return workbook.create_sheet(sheet_name)


def _set_sheet_very_hidden(workbook, sheet_name: str) -> None:
    """Mark a worksheet xlSheetVeryHidden (not visible via Format > Sheet > Unhide).

    openpyxl 3.1+ exposes sheet.sheet_state; legacy builds accept None silently so
    we swallow AttributeError if the attribute is missing.
    """
    try:
        if sheet_name in workbook.sheetnames:
            workbook[sheet_name].sheet_state = "veryHidden"
    except AttributeError:
        pass


def _write_transform_registry_block(
    workbook_path: Path,
    transform_bindings: list[TransformBinding],
) -> None:
    """Write only the YAML-driven transform-registry block to the dedicated
    'registry' sheet (created if missing).  The sheet is then marked VeryHidden
    so Finance users never stumble on it.  The Bootstrap Import Source/Count/At
    status cells live on the Runner sheet, written by BootstrapImporter.bas."""
    load_workbook = _import_openpyxl()

    workbook = load_workbook(workbook_path, keep_vba=True)
    try:
        worksheet = _ensure_worksheet(workbook, REGISTRY_SHEET_NAME)
        _write_transform_registry(worksheet, transform_bindings)
        _set_sheet_very_hidden(workbook, REGISTRY_SHEET_NAME)
        workbook.save(workbook_path)
    finally:
        archive = getattr(workbook, "vba_archive", None)
        if archive is not None:
            archive.close()
        workbook.close()


def _write_shell_defaults_config_block(
    workbook_path: Path,
    transform_bindings: list[TransformBinding],
) -> None:
    load_workbook = _import_openpyxl()

    workbook = load_workbook(workbook_path, keep_vba=True)
    try:
        if CONFIG_SHEET_NAME in workbook.sheetnames:
            worksheet = workbook[CONFIG_SHEET_NAME]
        else:
            worksheet = workbook.create_sheet(CONFIG_SHEET_NAME)
        _write_shell_defaults_config_sheet(worksheet, transform_bindings)
        workbook.save(workbook_path)
    finally:
        archive = getattr(workbook, "vba_archive", None)
        if archive is not None:
            archive.close()
        workbook.close()


def _rewrite_runner_formulas_block(
    workbook_path: Path,
    transform_bindings: list[TransformBinding],
) -> None:
    """Regenerate the Runner sheet's INDEX/MATCH lookup formulas after the config
    and registry sheets have been written by build_runner.py.

    create_runner_shell.py seeds formulas that reference columns G onwards on the
    Runner tab (historical registry location) and a no-banner config layout. By
    build time:
      - the Transform Registry lives on 'registry' (VeryHidden, rows 3..3+N-1)
      - the config sheet carries a build-time banner at row 1, pushing data rows
        down to 4..4+N-1 (CONFIG_START_ROW + 1 .. CONFIG_START_ROW + N)

    For the first N-1 transforms the stale formulas accidentally resolve (both
    offsets slip by one row and the MATCH lands on the same data row). The last
    transform's label falls outside the stale lookup range entirely and all its
    defaults return empty strings — the exact symptom reported for AR Multi Line
    Invoice when it is the final sorted profile.

    Rewriting here also means Source Type and EIB Population Type are sourced
    from the dedicated registry sheet rather than the cleared Runner-tab
    technical columns, which would otherwise be "" for every transform.
    """
    load_workbook = _import_openpyxl()

    workbook = load_workbook(workbook_path, keep_vba=True)
    try:
        if RUNNER_SHEET_NAME not in workbook.sheetnames:
            raise RuntimeError(
                f"Workbook {workbook_path} is missing worksheet named "
                f"{RUNNER_SHEET_NAME!r} (sheetnames={workbook.sheetnames}). "
                "Recreate the shell with scripts/create_runner_shell.py and rebuild."
            )
        worksheet = workbook[RUNNER_SHEET_NAME]
        _rewrite_runner_formulas(worksheet, transform_bindings)
        calc = workbook.calculation
        calc.fullCalcOnLoad = True
        calc.forceFullCalc = True
        workbook.save(workbook_path)
    finally:
        archive = getattr(workbook, "vba_archive", None)
        if archive is not None:
            archive.close()
        workbook.close()


def _rewrite_runner_formulas(worksheet, transform_bindings: list[TransformBinding]) -> None:
    RUNNER_CONTROL_ROWS_START = 3
    RUNNER_TRANSFORM_ROW = RUNNER_CONTROL_ROWS_START
    runner_transform_cell_ref = f"$B${RUNNER_TRANSFORM_ROW}"

    config_data_start = CONFIG_START_ROW + 1
    config_data_end = config_data_start + len(transform_bindings) - 1
    config_label_range = f"'{CONFIG_SHEET_NAME}'!$A${config_data_start}:$A${config_data_end}"

    def config_formula_for(col_letter: str) -> str:
        data_range = (
            f"'{CONFIG_SHEET_NAME}'!${col_letter}${config_data_start}:"
            f"${col_letter}${config_data_end}"
        )
        return (
            f'=IFERROR(INDEX({data_range},MATCH({runner_transform_cell_ref},'
            f"{config_label_range},0)),\"\")"
        )

    registry_data_start = REGISTRY_START_ROW
    registry_data_end = registry_data_start + len(transform_bindings) - 1
    registry_label_range = f"'{REGISTRY_SHEET_NAME}'!$A${registry_data_start}:$A${registry_data_end}"

    def registry_formula_for(col_index: int) -> str:
        col_letter = _column_letter(col_index)
        data_range = (
            f"'{REGISTRY_SHEET_NAME}'!${col_letter}${registry_data_start}:"
            f"${col_letter}${registry_data_end}"
        )
        return (
            f'=IFERROR(INDEX({data_range},MATCH({runner_transform_cell_ref},'
            f"{registry_label_range},0)),\"\")"
        )

    def _find_runner_label_row(label_text: str) -> int:
        max_scan = max(worksheet.max_row, RUNNER_CONTROL_ROWS_START + 50)
        for row_index in range(RUNNER_CONTROL_ROWS_START, max_scan + 1):
            cell = worksheet.cell(row=row_index, column=1)
            if str(cell.value or "").strip() == label_text:
                return row_index
        # Label not found in existing shell — create it by appending below the last populated row.
        # This handles forward-compatible label additions (e.g. source_layout fields) that
        # are not yet present in the base RunnerShell.xlsm created by create_runner_shell.py.
        next_empty = max_scan + 1
        for row_index in range(RUNNER_CONTROL_ROWS_START, next_empty + 1):
            if str(worksheet.cell(row=row_index, column=1).value or "").strip() == "":
                next_empty = row_index
                break
        worksheet.cell(row=next_empty, column=1).value = label_text
        return next_empty

    label_to_config_column = {
        "Source Data": "C",
        "Target EIB": "D",
        "Output File Path": "E",
        "Mapping Spreadsheet": "F",
        "Mapping Sheet": "G",
        "Customer Spreadsheet": "H",
        "Supplier Spreadsheet": "I",
        "Workday Reference IDs": "J",
        "Source Worksheet Name": "K",
        "Source Worksheet Resolution": "L",
        "Source Header Row": "M",
        "Source Data Start Row": "N",
        "Required Headers Pipe Delimited": "O",
    }
    for label_name, col_letter in label_to_config_column.items():
        row_num = _find_runner_label_row(label_name)
        worksheet.cell(row=row_num, column=2).value = config_formula_for(col_letter)

    REGISTRY_SOURCE_TYPE_COLUMN = REGISTRY_START_COLUMN + 2
    REGISTRY_EIB_POP_TYPE_COLUMN = REGISTRY_START_COLUMN + 11
    REGISTRY_PROFILE_KEY_COLUMN = REGISTRY_START_COLUMN + 1
    row_source_type = _find_runner_label_row("Source Type")
    worksheet.cell(row=row_source_type, column=2).value = registry_formula_for(
        REGISTRY_SOURCE_TYPE_COLUMN
    )
    row_eib_pop = _find_runner_label_row("EIB Population Type")
    worksheet.cell(row=row_eib_pop, column=2).value = registry_formula_for(
        REGISTRY_EIB_POP_TYPE_COLUMN
    )

    max_scan_e2 = max(worksheet.max_row, RUNNER_CONTROL_ROWS_START + 50)
    e2_row = None
    e6_row = None
    for row_index in range(1, max_scan_e2 + 1):
        d_cell = str(worksheet.cell(row=row_index, column=4).value or "").strip()
        if d_cell == "Selected Transform":
            e2_row = row_index
        elif d_cell == "Profile Key":
            e6_row = row_index
    if e2_row is not None:
        worksheet.cell(row=e2_row, column=5).value = f"=B{RUNNER_TRANSFORM_ROW}"
    if e6_row is not None:
        worksheet.cell(row=e6_row, column=5).value = registry_formula_for(
            REGISTRY_PROFILE_KEY_COLUMN
        )


def _column_letter(column_index: int) -> str:
    letter = ""
    n = int(column_index)
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter


def _write_transform_registry(worksheet, transform_bindings: list[TransformBinding]) -> None:
    max_row = max(worksheet.max_row, REGISTRY_START_ROW + len(transform_bindings) + 5)
    for row_index in range(REGISTRY_TITLE_ROW, max_row + 1):
        for column_index in range(REGISTRY_START_COLUMN, REGISTRY_END_COLUMN + 1):
            worksheet.cell(row=row_index, column=column_index).value = None

    worksheet.cell(row=REGISTRY_TITLE_ROW, column=REGISTRY_START_COLUMN).value = REGISTRY_TITLE
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN).value = "Transform Label"
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + 1).value = "Profile Key"
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + 2).value = "Source Type"
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + 3).value = "Source Data"
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + 4).value = "Target EIB"
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + 5).value = "Output File Path"
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + 6).value = "Mapping Spreadsheet"
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + 7).value = "Mapping Sheet"
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + 8).value = "Customer Spreadsheet"
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + 9).value = "Supplier Spreadsheet"
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + 10).value = "Workday Reference IDs"
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + 11).value = "EIB Population Type"
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + 12).value = "Step List Macro"
    worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + 13).value = "Start Row"

    for row_offset, binding in enumerate(transform_bindings):
        row_index = REGISTRY_START_ROW + row_offset
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN).value = binding.profile_label
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 1).value = binding.profile_key
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 2).value = binding.source_type
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 3).value = binding.source_data
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 4).value = binding.target_eib
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 5).value = binding.output_file_path
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 6).value = binding.mapping_spreadsheet
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 7).value = binding.mapping_sheet
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 8).value = binding.customer_spreadsheet
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 9).value = binding.supplier_spreadsheet
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 10).value = binding.workday_reference_ids
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 11).value = binding.eib_population_type
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 12).value = binding.step_list_macro
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 13).value = binding.start_row


def _write_shell_defaults_config_sheet(worksheet, transform_bindings: list[TransformBinding]) -> None:
    CONFIG_BANNER_ROW = CONFIG_TITLE_ROW
    CONFIG_TITLE_ROW_AFTER = CONFIG_TITLE_ROW + 1
    CONFIG_HEADER_ROW_AFTER = CONFIG_HEADER_ROW + 1
    CONFIG_START_ROW_AFTER = CONFIG_START_ROW + 1
    max_row = max(
        worksheet.max_row,
        CONFIG_START_ROW_AFTER + len(transform_bindings) + 5,
    )
    # This function re-merges CONFIG_BANNER_ROW below (banner spans columns 1-15). On a
    # first-ever call (fresh shell copy) nothing is merged yet, so this is a no-op; on any
    # later call against an already-built workbook (e.g. --substitute-root-paths patching
    # an existing build) the merge from the previous call is still there, and writing
    # .value into a non-top-left cell of an existing merge raises "'MergedCell' object
    # attribute 'value' is read-only". Unmerge first so this function is safely re-runnable.
    for merged_range in list(worksheet.merged_cells.ranges):
        if merged_range.min_row <= CONFIG_BANNER_ROW <= merged_range.max_row:
            worksheet.unmerge_cells(str(merged_range))
    for row_index in range(CONFIG_BANNER_ROW, max_row + 1):
        for column_index in range(1, 16):
            worksheet.cell(row=row_index, column=column_index).value = None

    banner_cell = worksheet.cell(
        row=CONFIG_BANNER_ROW,
        column=1,
        value=(
            "AUTO-GENERATED — EDITS BELOW ARE OVERWRITTEN ON EVERY BUILD. "
            "Edit the source of truth in config/transforms/*.yaml, then re-run "
            "build_runner.py. Machine-local path edits here are intentional "
            "(Finance users tune these after each build) and only persist until "
            "the next build run."
        ),
    )
    try:
        from openpyxl.styles import Alignment, Font, PatternFill

        banner_cell.font = Font(bold=True, color="FFFFFF")
        banner_cell.fill = PatternFill(fill_type="solid", fgColor="C0504D")
        banner_cell.alignment = Alignment(wrap_text=True, vertical="center")
    except Exception:  # pragma: no cover - style is cosmetic; still clear the row if styles unavailable
        pass
    worksheet.merge_cells(
        start_row=CONFIG_BANNER_ROW,
        start_column=1,
        end_row=CONFIG_BANNER_ROW,
        end_column=15,
    )
    try:
        worksheet.row_dimensions[CONFIG_BANNER_ROW].height = 42
    except Exception:  # pragma: no cover
        pass

    worksheet.cell(row=CONFIG_TITLE_ROW_AFTER, column=1).value = "Shell Defaults Config"
    worksheet.cell(row=CONFIG_TITLE_ROW_AFTER, column=2).value = (
        "Business-editable machine-specific values used to populate the Runner sheet."
    )
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=1).value = "Transform Label"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=2).value = "Profile Key"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=3).value = "Source Data"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=4).value = "Target EIB"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=5).value = "Output File Path"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=6).value = "Mapping Spreadsheet"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=7).value = "Mapping Sheet"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=8).value = "Customer Spreadsheet"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=9).value = "Supplier Spreadsheet"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=10).value = "Workday Reference IDs"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=11).value = "Source Worksheet Name"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=12).value = "Source Worksheet Resolution"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=13).value = "Source Header Row"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=14).value = "Source Data Start Row"
    worksheet.cell(row=CONFIG_HEADER_ROW_AFTER, column=15).value = "Required Headers Pipe Delimited"

    for row_offset, binding in enumerate(transform_bindings):
        row_index = CONFIG_START_ROW_AFTER + row_offset
        worksheet.cell(row=row_index, column=1).value = binding.profile_label
        worksheet.cell(row=row_index, column=2).value = binding.profile_key
        worksheet.cell(row=row_index, column=3).value = binding.source_data
        worksheet.cell(row=row_index, column=4).value = binding.target_eib
        worksheet.cell(row=row_index, column=5).value = binding.output_file_path
        worksheet.cell(row=row_index, column=6).value = binding.mapping_spreadsheet
        worksheet.cell(row=row_index, column=7).value = binding.mapping_sheet
        worksheet.cell(row=row_index, column=8).value = binding.customer_spreadsheet
        worksheet.cell(row=row_index, column=9).value = binding.supplier_spreadsheet
        worksheet.cell(row=row_index, column=10).value = binding.workday_reference_ids
        worksheet.cell(row=row_index, column=11).value = binding.worksheet_name or ""
        worksheet.cell(row=row_index, column=12).value = binding.worksheet_resolution or ""
        worksheet.cell(row=row_index, column=13).value = binding.header_row or ""
        worksheet.cell(row=row_index, column=14).value = binding.data_start_row or ""
        # Separate required_headers entries with "||" (DOUBLE pipe) because each
        # individual entry may itself contain single "|" for per-header alias lists
        # e.g. "Total Claim ()|Total Claim|Total Claim (GBP)".  Single pipe is
        # preserved verbatim; VBA helper ParseRequiredHeadersFromControl splits on
        # "||" only and leaves internal aliases intact for FindColumnIndexInRow.
        worksheet.cell(row=row_index, column=15).value = "||".join(
            str(h) for h in (binding.required_headers or [])
        )


def _import_pyopenvba():
    try:
        from pyopenvba import ExcelFile, VBAModuleKind
    except ImportError as exc:  # pragma: no cover - dependency boundary
        raise RuntimeError(
            "pyOpenVBA is required. Run with:\n"
            "uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python "
            "elt_transform_excel/scripts/build_runner.py"
        ) from exc
    return ExcelFile, VBAModuleKind


def _import_openpyxl():
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover - dependency boundary
        raise RuntimeError(
            "openpyxl is required. Run with:\n"
            "uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python "
            "elt_transform_excel/scripts/build_runner.py"
        ) from exc
    return load_workbook


def _import_yaml():
    try:
        import yaml
    except ImportError as exc:  # pragma: no cover - dependency boundary
        raise RuntimeError(
            "PyYAML is required. Run with:\n"
            "uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python "
            "elt_transform_excel/scripts/build_runner.py"
        ) from exc
    return yaml


def _read_bootstrap_status(workbook_path: Path) -> BootstrapStatus:
    sheet_cells = _read_named_sheet_cells(workbook_path, RUNNER_SHEET_NAME)
    labels = {
        STATUS_SOURCE: None,
        STATUS_COUNT: None,
        STATUS_AT: None,
    }

    for cell_ref, value in sheet_cells.items():
        text_value = str(value).strip()
        if text_value not in labels:
            continue
        labels[text_value] = sheet_cells.get(_adjacent_cell_ref(cell_ref), "").strip()

    missing = [label for label, value in labels.items() if not value]
    if missing:
        raise RuntimeError(
            "Build step did not write all expected status cells. Missing: "
            + ", ".join(missing)
        )

    try:
        import_count = int(float(str(labels[STATUS_COUNT])))
    except ValueError as exc:
        raise RuntimeError(f"Bootstrap import count is not numeric: {labels[STATUS_COUNT]!r}") from exc

    return BootstrapStatus(
        source=str(labels[STATUS_SOURCE]),
        import_count=import_count,
        imported_at=str(labels[STATUS_AT]),
    )


def _read_named_sheet_cells(workbook_path: Path, sheet_name: str) -> dict[str, str]:
    with zipfile.ZipFile(workbook_path) as workbook_zip:
        shared_strings = _read_shared_strings(workbook_zip)
        sheet_path = _resolve_named_sheet_path(workbook_zip, sheet_name)
        worksheet_root = ElementTree.fromstring(workbook_zip.read(sheet_path))

    namespace = {"main": WORKBOOK_XML_NS}
    cells: dict[str, str] = {}
    for cell in worksheet_root.findall(".//main:sheetData/main:row/main:c", namespace):
        cell_ref = cell.attrib.get("r")
        if not cell_ref:
            continue
        cells[cell_ref] = _cell_value(cell, shared_strings).strip()
    return cells


def _read_first_sheet_cells(workbook_path: Path) -> dict[str, str]:
    with zipfile.ZipFile(workbook_path) as workbook_zip:
        shared_strings = _read_shared_strings(workbook_zip)
        sheet_path = _resolve_first_sheet_path(workbook_zip)
        worksheet_root = ElementTree.fromstring(workbook_zip.read(sheet_path))

    namespace = {"main": WORKBOOK_XML_NS}
    cells: dict[str, str] = {}
    for cell in worksheet_root.findall(".//main:sheetData/main:row/main:c", namespace):
        cell_ref = cell.attrib.get("r")
        if not cell_ref:
            continue
        cells[cell_ref] = _cell_value(cell, shared_strings).strip()
    return cells


def _read_shared_strings(workbook_zip: zipfile.ZipFile) -> list[str]:
    try:
        shared_strings_xml = workbook_zip.read("xl/sharedStrings.xml")
    except KeyError:
        return []

    root = ElementTree.fromstring(shared_strings_xml)
    namespace = {"main": WORKBOOK_XML_NS}
    values: list[str] = []
    for item in root.findall("main:si", namespace):
        text_parts = [node.text or "" for node in item.findall(".//main:t", namespace)]
        values.append("".join(text_parts))
    return values


def _resolve_named_sheet_path(workbook_zip: zipfile.ZipFile, sheet_name: str) -> str:
    workbook_root = ElementTree.fromstring(workbook_zip.read("xl/workbook.xml"))
    rel_root = ElementTree.fromstring(workbook_zip.read("xl/_rels/workbook.xml.rels"))
    namespace = {"main": WORKBOOK_XML_NS, "rel": PACKAGE_REL_NS}

    target_sheet = None
    for sheet in workbook_root.findall("main:sheets/main:sheet", namespace):
        if sheet.attrib.get("name") == sheet_name:
            target_sheet = sheet
            break
    if target_sheet is None:
        sheet_names = [
            s.attrib.get("name")
            for s in workbook_root.findall("main:sheets/main:sheet", namespace)
        ]
        raise RuntimeError(
            f"Workbook does not contain a worksheet named {sheet_name!r}. "
            f"Sheets present: {sheet_names!r}"
        )

    relationship_id = target_sheet.attrib.get(f"{{{REL_NS}}}id")
    if not relationship_id:
        raise RuntimeError(f"Worksheet {sheet_name!r} relationship id not found")

    for relationship in rel_root.findall("rel:Relationship", namespace):
        if relationship.attrib.get("Id") == relationship_id:
            target = relationship.attrib["Target"].lstrip("/")
            if target.startswith("xl/"):
                return target
            return f"xl/{target}"

    raise RuntimeError(
        f"Worksheet target not found for relationship id {relationship_id}"
    )


def _resolve_first_sheet_path(workbook_zip: zipfile.ZipFile) -> str:
    workbook_root = ElementTree.fromstring(workbook_zip.read("xl/workbook.xml"))
    rel_root = ElementTree.fromstring(workbook_zip.read("xl/_rels/workbook.xml.rels"))
    namespace = {"main": WORKBOOK_XML_NS, "rel": PACKAGE_REL_NS}

    first_sheet = workbook_root.find("main:sheets/main:sheet", namespace)
    if first_sheet is None:
        raise RuntimeError("Workbook does not contain any worksheets")

    relationship_id = first_sheet.attrib.get(f"{{{REL_NS}}}id")
    if not relationship_id:
        raise RuntimeError("First worksheet relationship id not found")

    for relationship in rel_root.findall("rel:Relationship", namespace):
        if relationship.attrib.get("Id") == relationship_id:
            target = relationship.attrib["Target"].lstrip("/")
            if target.startswith("xl/"):
                return target
            return f"xl/{target}"

    raise RuntimeError(f"Worksheet target not found for relationship id {relationship_id}")


def _cell_value(cell: ElementTree.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//{%s}t" % WORKBOOK_XML_NS))

    value_node = cell.find(f"{{{WORKBOOK_XML_NS}}}v")
    if value_node is None or value_node.text is None:
        return ""

    raw_value = value_node.text
    if cell_type == "s":
        return shared_strings[int(raw_value)]
    return raw_value


def _adjacent_cell_ref(cell_ref: str) -> str:
    match = re.fullmatch(r"([A-Z]+)(\d+)", cell_ref)
    if match is None:
        raise RuntimeError(f"Unsupported cell reference: {cell_ref}")
    column_label, row_number = match.groups()
    return f"{_next_column_label(column_label)}{row_number}"


def _next_column_label(column_label: str) -> str:
    index = 0
    for char in column_label:
        index = (index * 26) + (ord(char) - ord("A") + 1)
    index += 1

    next_label = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        next_label = chr(ord("A") + remainder) + next_label
    return next_label


def _verify_runner_sheet_invariants(workbook_path: Path) -> None:
    """Verify the Runner sheet is named correctly and positioned first."""
    with zipfile.ZipFile(workbook_path) as workbook_zip:
        workbook_root = ElementTree.fromstring(workbook_zip.read("xl/workbook.xml"))
    namespace = {"main": WORKBOOK_XML_NS}
    sheets = workbook_root.findall("main:sheets/main:sheet", namespace)
    if not sheets:
        raise RuntimeError("Workbook does not contain any worksheets")
    first_sheet_name = sheets[0].attrib.get("name")
    sheet_names = [s.attrib.get("name") for s in sheets]
    if RUNNER_SHEET_NAME not in sheet_names:
        raise RuntimeError(
            f"Workbook is missing worksheet named {RUNNER_SHEET_NAME!r}. "
            f"Sheets present: {sheet_names!r}. "
            "Recreate the shell with scripts/create_runner_shell.py and rebuild."
        )
    if first_sheet_name != RUNNER_SHEET_NAME:
        raise RuntimeError(
            f"First worksheet is {first_sheet_name!r}, expected {RUNNER_SHEET_NAME!r}. "
            f"Full sheet order: {sheet_names!r}. "
            "A tab drag has likely reordered the workbook; recreate the shell."
        )


def _verify_bootstrap_status(
    status: BootstrapStatus,
    source_path: Path,
    expected_module_names: list[str],
    workbook_path: Path,
) -> None:
    _verify_runner_sheet_invariants(workbook_path)
    expected_count = len(expected_module_names)
    if status.import_count != expected_count:
        raise RuntimeError(
            f"Bootstrap importer reported {status.import_count} module(s), expected {expected_count}."
        )

    reported_source = Path(status.source).expanduser()
    if not reported_source.exists():
        raise RuntimeError(
            f"Bootstrap importer reported source path that does not exist: {reported_source}"
        )
    if reported_source.resolve() != source_path.resolve():
        raise RuntimeError(
            f"Bootstrap importer used unexpected source path: {reported_source} "
            f"(expected {source_path})"
        )
    if not status.imported_at.strip():
        raise RuntimeError("Bootstrap importer did not record an import timestamp")


def _verify_workbook_modules(workbook_path: Path, expected_module_names: list[str]) -> None:
    excel_file_cls, _ = _import_pyopenvba()
    expected_names = {name.lower() for name in expected_module_names}

    with excel_file_cls(str(workbook_path)) as workbook:
        actual_importable_names: set[str] = set()
        for module in workbook.vba_project().modules:
            if _preserve_existing_module(module):
                continue
            actual_importable_names.add(module.name.lower())

    missing = sorted(expected_names - actual_importable_names)
    unexpected = sorted(actual_importable_names - expected_names)
    if missing or unexpected:
        message_parts: list[str] = []
        if missing:
            message_parts.append("missing: " + ", ".join(missing))
        if unexpected:
            message_parts.append("unexpected: " + ", ".join(unexpected))
        raise RuntimeError("Built workbook module set mismatch: " + " | ".join(message_parts))


def _verify_transform_registry(
    workbook_path: Path,
    expected_bindings: list[TransformBinding],
) -> None:
    load_workbook = _import_openpyxl()
    workbook = load_workbook(workbook_path, read_only=True, keep_vba=True)
    try:
        if REGISTRY_SHEET_NAME not in workbook.sheetnames:
            raise RuntimeError(
                f"Built workbook is missing worksheet: {REGISTRY_SHEET_NAME!r} "
                "(Transform Registry should live here on a very-hidden sheet, not on the Runner tab)."
            )
        worksheet = workbook[REGISTRY_SHEET_NAME]
        actual_bindings = []
        for row_index in range(REGISTRY_START_ROW, REGISTRY_START_ROW + len(expected_bindings)):
            profile_label = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN).value or ""
            ).strip()
            profile_key = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 1).value or ""
            ).strip()
            source_type = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 2).value or ""
            ).strip()
            source_data = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 3).value or ""
            ).strip()
            target_eib = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 4).value or ""
            ).strip()
            output_file_path = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 5).value or ""
            ).strip()
            mapping_spreadsheet = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 6).value or ""
            ).strip()
            mapping_sheet = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 7).value or ""
            ).strip()
            customer_spreadsheet = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 8).value or ""
            ).strip()
            supplier_spreadsheet = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 9).value or ""
            ).strip()
            workday_reference_ids = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 10).value or ""
            ).strip()
            eib_population_type = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 11).value or ""
            ).strip()
            step_list_macro = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 12).value or ""
            ).strip()
            start_row = str(
                worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 13).value or ""
            ).strip()
            actual_bindings.append(
                (
                    profile_label,
                    profile_key,
                    source_type,
                    source_data,
                    target_eib,
                    output_file_path,
                    mapping_spreadsheet,
                    mapping_sheet,
                    customer_spreadsheet,
                    supplier_spreadsheet,
                    workday_reference_ids,
                    eib_population_type,
                    step_list_macro,
                    start_row,
                )
            )
    finally:
        archive = getattr(workbook, "vba_archive", None)
        if archive is not None:
            archive.close()
        workbook.close()

    expected_rows = [
        (
            binding.profile_label,
            binding.profile_key,
            binding.source_type,
            binding.source_data,
            binding.target_eib,
            binding.output_file_path,
            binding.mapping_spreadsheet,
            binding.mapping_sheet,
            binding.customer_spreadsheet,
            binding.supplier_spreadsheet,
            binding.workday_reference_ids,
            binding.eib_population_type,
            binding.step_list_macro,
            "" if binding.start_row is None else str(binding.start_row),
        )
        for binding in expected_bindings
    ]
    if actual_bindings != expected_rows:
        raise RuntimeError(
            "Built workbook transform registry mismatch: "
            f"expected {expected_rows!r}, got {actual_bindings!r}"
        )


def _verify_shell_defaults_config(
    workbook_path: Path,
    expected_bindings: list[TransformBinding],
) -> None:
    load_workbook = _import_openpyxl()
    workbook = load_workbook(workbook_path, read_only=True, keep_vba=True)
    try:
        if CONFIG_SHEET_NAME not in workbook.sheetnames:
            raise RuntimeError(f"Built workbook is missing worksheet: {CONFIG_SHEET_NAME}")

        worksheet = workbook[CONFIG_SHEET_NAME]
        # Banner (row 1) is an overwrite warning so Finance users don't lose
        # machine-local path edits they made after the last build.
        banner_text = str(worksheet.cell(row=CONFIG_TITLE_ROW, column=1).value or "").strip()
        if "OVERWRITTEN ON EVERY BUILD" not in banner_text:
            raise RuntimeError(
                f"Built workbook config sheet row {CONFIG_TITLE_ROW} is missing the "
                f"'OVERWRITTEN ON EVERY BUILD' banner. This banner keeps Finance "
                f"from being confused when hand-edits below are replaced by the "
                f"next build. Got: {banner_text[:80]!r}"
            )
        actual_rows = []
        config_start = CONFIG_START_ROW + 1  # banner row inserted above title/header
        for row_index in range(config_start, config_start + len(expected_bindings)):
            actual_rows.append(
                (
                    str(worksheet.cell(row=row_index, column=1).value or "").strip(),
                    str(worksheet.cell(row=row_index, column=2).value or "").strip(),
                    str(worksheet.cell(row=row_index, column=3).value or "").strip(),
                    str(worksheet.cell(row=row_index, column=4).value or "").strip(),
                    str(worksheet.cell(row=row_index, column=5).value or "").strip(),
                    str(worksheet.cell(row=row_index, column=6).value or "").strip(),
                    str(worksheet.cell(row=row_index, column=7).value or "").strip(),
                    str(worksheet.cell(row=row_index, column=8).value or "").strip(),
                    str(worksheet.cell(row=row_index, column=9).value or "").strip(),
                    str(worksheet.cell(row=row_index, column=10).value or "").strip(),
                    str(worksheet.cell(row=row_index, column=11).value or "").strip(),
                    str(worksheet.cell(row=row_index, column=12).value or "").strip(),
                    str(worksheet.cell(row=row_index, column=13).value or "").strip(),
                    str(worksheet.cell(row=row_index, column=14).value or "").strip(),
                )
            )
    finally:
        archive = getattr(workbook, "vba_archive", None)
        if archive is not None:
            archive.close()
        workbook.close()

    def _format_int(val: object) -> str:
        return "" if val in (None, "") else str(val)

    expected_rows = [
        (
            binding.profile_label,
            binding.profile_key,
            binding.source_data,
            binding.target_eib,
            binding.output_file_path,
            binding.mapping_spreadsheet,
            binding.mapping_sheet,
            binding.customer_spreadsheet,
            binding.supplier_spreadsheet,
            binding.workday_reference_ids,
            binding.worksheet_name or "",
            binding.worksheet_resolution or "",
            _format_int(binding.header_row),
            _format_int(binding.data_start_row),
        )
        for binding in expected_bindings
    ]
    if actual_rows != expected_rows:
        raise RuntimeError(
            "Built workbook config sheet mismatch: "
            f"expected {expected_rows!r}, got {actual_rows!r}"
        )


def _verify_mapping_sheet(workbook_path: Path) -> None:
    load_workbook = _import_openpyxl()

    expected_rows = [MAPPING_HEADER, *_mapping_rows_from_lookups()]
    workbook = load_workbook(workbook_path, read_only=True, keep_vba=True)
    try:
        if MAPPING_SHEET_NAME not in workbook.sheetnames:
            raise RuntimeError(f"Built workbook is missing worksheet: {MAPPING_SHEET_NAME}")

        worksheet = workbook[MAPPING_SHEET_NAME]
        actual_rows = []
        for row_index in range(1, len(expected_rows) + 1):
            actual_rows.append(
                tuple(
                    str(worksheet.cell(row=row_index, column=column_index).value or "").strip()
                    for column_index in range(1, len(MAPPING_HEADER) + 1)
                )
            )
    finally:
        archive = getattr(workbook, "vba_archive", None)
        if archive is not None:
            archive.close()
        workbook.close()

    if actual_rows != expected_rows:
        raise RuntimeError(
            "Built workbook mapping sheet mismatch: "
            f"expected {expected_rows!r}, got {actual_rows!r}"
        )


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # pragma: no cover - CLI error boundary
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
