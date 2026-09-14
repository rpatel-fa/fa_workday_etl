#!/usr/bin/env python3
"""Create the minimal Runner shell workbook with bootstrap VBA and control-sheet layout."""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve()
MODULE_ROOT = SCRIPT_PATH.parents[1]
DEFAULT_OUTPUT = MODULE_ROOT / "runner_shell" / "RunnerShell.xlsm"
DEFAULT_PROFILE_DIR = MODULE_ROOT / "config" / "transforms"
DEFAULT_BUILD_MANIFEST = MODULE_ROOT / "config" / "build_manifest.yaml"
BOOTSTRAP_SOURCE = MODULE_ROOT / "vba_src" / "bootstrap" / "BootstrapImporter.bas"

RUNNER_SHEET_NAME = "Runner"
CONFIG_SHEET_NAME = "config"
GENERATED_SHEET_NAMES = {RUNNER_SHEET_NAME, CONFIG_SHEET_NAME}

REGISTRY_START_COLUMN = 7
REGISTRY_END_COLUMN = 20
REGISTRY_TITLE_ROW = 2
REGISTRY_HEADER_ROW = 3
REGISTRY_START_ROW = 4

CONFIG_TITLE_ROW = 1
CONFIG_HEADER_ROW = 2
CONFIG_START_ROW = 3

TITLE = "elt_transform_excel Runner Shell"
INSTRUCTION_TEXT = (
    "Pick a Transform on the left, then click the 'Run Transform' button. "
    "The button is created by the bootstrap import (it cannot exist before the macro "
    "behind it is imported); until then, run modDispatcher.RunSelectedTransform from "
    "Tools > Macro > Macros."
)

CONTROL_LABELS = [
    "Transform",
    "Source Type",
    "Source Data",
    "Target EIB",
    "Output File Path",
    "Mapping Spreadsheet",
    "Mapping Sheet",
    "Customer Spreadsheet",
    "Supplier Spreadsheet",
    "Workday Reference IDs",
    "EIB Population Type",
]
CONTROL_ROWS_START = 3

RESULT_ROWS = [
    ("Run Status", "Not run yet"),
    ("Rows Processed", "0"),
    ("Rows Skipped", "0"),
    ("Issues", ""),
    ("Output Workbook", ""),
    ("Selection Key", ""),
    ("Last Run At", ""),
]


@dataclass(frozen=True)
class TransformProfile:
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
    worksheet_name: str | None
    worksheet_resolution: str | None
    header_row: int | None
    data_start_row: int | None

def create_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Create runner_shell/RunnerShell.xlsm with a single visible control sheet and "
            "the BootstrapImporter VBA module sourced from vba_src/bootstrap/BootstrapImporter.bas."
        )
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT),
        help=f"Path to the Runner shell workbook to create (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace the existing output workbook if it already exists.",
    )
    parser.add_argument(
        "--build-manifest",
        default=str(DEFAULT_BUILD_MANIFEST),
        help=(
            "Path to build_manifest.yaml — single source of truth for active/defunct transform "
            "keys AND the default_profile_key used to pre-populate the Runner control sheet. "
            f"(default: {DEFAULT_BUILD_MANIFEST})"
        ),
    )
    parser.add_argument(
        "--profile-dir",
        default=str(DEFAULT_PROFILE_DIR),
        help=(
            "Folder of all transform profile YAML files (active + defunct live here as files; "
            f"the manifest decides which ones are actually bound into the Runner shell). "
            f"(default: {DEFAULT_PROFILE_DIR})"
        ),
    )
    return parser


def main() -> int:
    args = create_parser().parse_args()

    output_path = Path(args.output).expanduser().resolve()
    _validate_output_path(output_path, overwrite=args.overwrite)
    manifest_path = Path(args.build_manifest).expanduser().resolve()
    manifest = _load_build_manifest(manifest_path)
    default_key: str = manifest["default_profile_key"]
    active_set: set[str] = manifest["_active_set"]
    active_keys_ordered: list[str] = manifest["_active_keys_ordered"]
    profile_dir = Path(args.profile_dir).expanduser().resolve()
    profile_path = profile_dir / f"{default_key}.yaml"
    if not profile_path.exists():
        raise FileNotFoundError(
            f"Manifest default_profile_key={default_key!r} has no matching YAML at {profile_path}. "
            "Ensure the declared default exists in the profile directory."
        )
    active_profile = _load_profile_config(profile_path)
    transform_profiles = _load_transform_profiles(profile_dir, active_keys_ordered)
    control_rows = _build_control_rows(active_profile)
    transform_label_options = [profile.profile_label for profile in transform_profiles]

    bootstrap_source = _read_bootstrap_source(BOOTSTRAP_SOURCE)

    excel_file_cls, module_kind = _import_pyopenvba()
    load_workbook, DataValidation = _import_openpyxl()

    # create_new() below replaces the whole file, including any hand-authored sheets
    # (e.g. live mapping tabs) that this script does not itself generate. Capture them
    # first so --overwrite refreshes the generated Runner sheet/VBA without silently
    # destroying hand-authored data — this is exactly how the "arinvoice"/"mapping"
    # mapping tabs were lost in a previous session.
    preserved_sheets = _read_preserved_sheets(output_path, load_workbook)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with excel_file_cls.create_new(str(output_path)) as workbook:
        project = workbook.vba_project()
        # pyOpenVBA's create_new() seeds a fresh VBA project with a default empty
        # "Module1" placeholder; remove it so the shell holds only BootstrapImporter,
        # per the one-sheet-plus-bootstrap-only invariant (ARCHITECTURE.md).
        existing_names = {module.name for module in project.modules}
        if "Module1" in existing_names:
            project.delete_module("Module1")
        project.add_module(
            "BootstrapImporter",
            bootstrap_source,
            kind=module_kind.standard,
        )
        workbook.save()

    workbook = load_workbook(output_path, keep_vba=True)
    try:
        worksheet = workbook.active
        worksheet.title = RUNNER_SHEET_NAME
        config_sheet = workbook.create_sheet(CONFIG_SHEET_NAME)

        _write_layout(worksheet, control_rows, active_profile)
        _write_transform_registry(worksheet, transform_profiles)
        _write_shell_defaults_config(config_sheet, transform_profiles)
        _apply_profile_formulas(worksheet, active_profile, len(transform_profiles))
        _add_dropdown_validation(
            worksheet,
            DataValidation,
            transform_label_options,
        )
        _write_preserved_sheets(workbook, preserved_sheets)

        workbook.save(output_path)
    finally:
        archive = getattr(workbook, "vba_archive", None)
        if archive is not None:
            archive.close()
        workbook.close()

    print(f"Created Runner shell: {output_path}")
    print(f"Profile config: {profile_path}")
    print(f"Transform profile: {_profile_value(active_profile, 'transform', 'label')}")
    print("Bootstrap module: BootstrapImporter")
    print(f"Control rows: {len(control_rows)}")
    print(f"Result rows: {len(RESULT_ROWS)}")
    return 0


def _read_preserved_sheets(output_path: Path, load_workbook) -> dict[str, list[tuple[int, int, object]]]:
    if not output_path.exists():
        return {}

    workbook = load_workbook(output_path, data_only=False)
    try:
        return {
            name: [
                (cell.row, cell.column, cell.value)
                for row in workbook[name].iter_rows()
                for cell in row
                if cell.value is not None
            ]
            for name in workbook.sheetnames
            if name not in GENERATED_SHEET_NAMES
        }
    finally:
        workbook.close()


def _write_preserved_sheets(
    workbook, preserved_sheets: dict[str, list[tuple[int, int, object]]]
) -> None:
    for name, cells in preserved_sheets.items():
        sheet = workbook.create_sheet(name)
        for row, column, value in cells:
            sheet.cell(row=row, column=column, value=value)


def _validate_output_path(output_path: Path, overwrite: bool) -> None:
    if output_path.suffix.lower() != ".xlsm":
        raise ValueError(f"Output path must be an .xlsm workbook: {output_path}")
    if output_path.exists() and not overwrite:
        raise FileExistsError(
            f"Output already exists: {output_path}\n"
            "Pass --overwrite to replace it."
        )


def _read_bootstrap_source(source_path: Path) -> str:
    if not source_path.exists():
        raise FileNotFoundError(f"Bootstrap VBA source not found: {source_path}")

    source_text = source_path.read_text(encoding="utf-8")
    normalized = source_text.replace("\r\n", "\n").replace("\r", "\n")
    return normalized.replace("\n", "\r\n")


def _load_build_manifest(manifest_path: Path) -> dict:
    yaml = _import_yaml()
    if not manifest_path.exists():
        raise FileNotFoundError(
            f"Build manifest not found: {manifest_path}. Expected at implementations/elt_transform_excel/config/build_manifest.yaml."
        )
    with manifest_path.open("r", encoding="utf-8") as handle:
        manifest = yaml.safe_load(handle) or {}
    for required_key in ("active_profiles", "defunct_profiles", "default_profile_key"):
        if required_key not in manifest:
            raise RuntimeError(
                f"Build manifest {manifest_path} missing required key '{required_key}'."
            )
    if not isinstance(manifest["active_profiles"], list):
        raise RuntimeError("Build manifest: active_profiles must be a list.")
    if not isinstance(manifest["defunct_profiles"], list):
        raise RuntimeError("Build manifest: defunct_profiles must be a list.")
    if not isinstance(manifest["default_profile_key"], str):
        raise RuntimeError("Build manifest: default_profile_key must be a string.")
    active_set: set[str] = {str(x) for x in manifest["active_profiles"]}
    defunct_set: set[str] = {str(x) for x in manifest["defunct_profiles"]}
    # Keep ORDERED manifest declaration order for profile loading:
    active_keys_ordered: list[str] = [str(x) for x in manifest["active_profiles"]]
    overlap = sorted(active_set & defunct_set)
    if overlap:
        raise RuntimeError(f"Build manifest: keys in BOTH active_profiles and defunct_profiles: {overlap}")
    if manifest["default_profile_key"] not in active_set:
        raise RuntimeError(
            f"Build manifest: default_profile_key={manifest['default_profile_key']!r} not in active_profiles "
            f"(active: {sorted(active_set)})"
        )
    manifest["_active_set"] = active_set
    manifest["_defunct_set"] = defunct_set
    manifest["_active_keys_ordered"] = active_keys_ordered
    return manifest


def _load_profile_config(profile_path: Path) -> dict:
    if not profile_path.exists():
        raise FileNotFoundError(f"Profile config not found: {profile_path}")

    yaml = _import_yaml()
    with profile_path.open("r", encoding="utf-8") as handle:
        profile = yaml.safe_load(handle) or {}

    if not isinstance(profile, dict):
        raise ValueError(f"Profile config must parse to a mapping: {profile_path}")

    required_sections = ("transform", "shell_defaults")
    for section in required_sections:
        if section not in profile or not isinstance(profile[section], dict):
            raise ValueError(
                f"Profile config must contain a '{section}' mapping: {profile_path}"
            )

    return profile


def _load_reference_paths(reference_path: Path) -> dict:
    """Load the shared reference-data paths (customer/supplier/mapping) that every transform
    inherits unless its profile overrides the key. Single source of truth so an updated Workday
    extract or finance mapping file is changed once, not in every profile."""
    yaml = _import_yaml()
    if not reference_path.is_file():
        raise ValueError(
            f"Shared reference-paths file not found: {reference_path}. "
            "Transform profiles inherit customer/supplier/mapping paths from it."
        )
    with reference_path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}
    paths = data.get("reference_paths")
    if not isinstance(paths, dict):
        raise ValueError(f"'{reference_path}' must contain a 'reference_paths' mapping.")
    return {key: str(value or "") for key, value in paths.items()}


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
    base_profile: TransformProfile,
) -> list[TransformProfile]:
    import dataclasses

    shell_defaults = profile.get("shell_defaults")
    if not isinstance(shell_defaults, dict):
        return [base_profile]
    variants_raw = shell_defaults.get("variants")
    if variants_raw is None:
        return [base_profile]
    if not isinstance(variants_raw, list) or not variants_raw:
        raise ValueError(
            f"Profile '{profile_path}' shell_defaults.variants must be a non-empty list "
            f"when declared. If you want the single base profile, remove the 'variants' key."
        )

    variant_profiles: list[TransformProfile] = []
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
        variant_profile = dataclasses.replace(base_profile, **overrides)
        variant_profiles.append(variant_profile)
    return variant_profiles


def _load_transform_profiles(profile_dir: Path, active_keys_ordered: list[str]) -> list[TransformProfile]:
    yaml = _import_yaml()
    # Iterate in the exact order declared in manifest.active_profiles (not set/glob order, not sorted by filename).
    profile_paths: list[Path] = []
    for active_key in active_keys_ordered:
        profile_path = profile_dir / f"{active_key}.yaml"
        if not profile_path.exists():
            raise FileNotFoundError(
                f"build_manifest declares active profile '{active_key}' but no YAML at {profile_path}."
            )
        profile_paths.append(profile_path)
    if not profile_paths:
        raise ValueError("active_profiles in build_manifest is empty; nothing to write to the Runner shell.")

    reference_paths = _load_reference_paths(profile_dir.parent / "reference_paths.yaml")

    profiles: list[TransformProfile] = []
    seen_keys: set[str] = set()
    seen_labels: set[str] = set()
    for profile_path in profile_paths:
        with profile_path.open("r", encoding="utf-8") as handle:
            profile_data = yaml.safe_load(handle) or {}
        base_profile = _parse_transform_profile(profile_path, profile_data, reference_paths)
        expanded = _expand_shell_variants(profile_path, profile_data, base_profile)
        for profile in expanded:
            normalized_key = profile.profile_key.lower()
            if normalized_key in seen_keys:
                raise ValueError(f"Duplicate transform.key '{profile.profile_key}': {profile_path}")
            seen_keys.add(normalized_key)

            normalized_label = profile.profile_label.lower()
            if normalized_label in seen_labels:
                raise ValueError(f"Duplicate transform.label '{profile.profile_label}': {profile_path}")
            seen_labels.add(normalized_label)

            profiles.append(profile)
    # Present transforms to the user in ALPHABETICAL order by human-readable label
    # (regardless of manifest declaration order or file discovery order).
    profiles.sort(key=lambda p: (p.profile_label.lower(), p.profile_key.lower()))
    return profiles


def _optional_profile_section_value(
    profile_path: Path,
    section_data: dict,
    section_name: str,
    key: str,
) -> str | None:
    value = section_data.get(key)
    if value in ("", None):
        return None
    return str(value)


def _parse_transform_profile(
    profile_path: Path, profile: dict, reference_paths: dict | None = None
) -> TransformProfile:
    reference_paths = reference_paths or {}
    if not isinstance(profile, dict):
        raise ValueError(f"Profile config must parse to a mapping: {profile_path}")

    transform = profile.get("transform")
    runtime = profile.get("runtime")
    shell_defaults = profile.get("shell_defaults")
    if not isinstance(transform, dict):
        raise ValueError(f"Transform profile is missing a 'transform' mapping: {profile_path}")
    if not isinstance(runtime, dict):
        raise ValueError(f"Transform profile is missing a 'runtime' mapping: {profile_path}")
    if not isinstance(shell_defaults, dict):
        raise ValueError(f"Transform profile is missing a 'shell_defaults' mapping: {profile_path}")

    has_variants = (
        isinstance(shell_defaults.get("variants"), list) and bool(shell_defaults["variants"])
    )
    if has_variants:
        source_data_val = _optional_profile_section_value(
            profile_path, shell_defaults, "shell_defaults", "source_data"
        ) or ""
        output_val = _optional_profile_section_value(
            profile_path, shell_defaults, "shell_defaults", "output_file_path"
        ) or ""
    else:
        source_data_val = _profile_section_value(
            profile_path, shell_defaults, "shell_defaults", "source_data"
        )
        output_val = _profile_section_value(
            profile_path, shell_defaults, "shell_defaults", "output_file_path"
        )

    source_layout = profile.get("source_layout") or {}

    def _shell_or_source_str(shell_key: str, source_key: str) -> str | None:
        shell_val = _optional_profile_section_value(
            profile_path, shell_defaults, "shell_defaults", shell_key
        )
        if shell_val is not None:
            return shell_val
        return _optional_profile_section_value(
            profile_path, source_layout, "source_layout", source_key
        )

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
        src_raw = source_layout.get(source_key)
        if src_raw in ("", None):
            return None
        try:
            src_int = int(src_raw)
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"Profile '{profile_path}' source_layout.{source_key} must be int, got: {src_raw!r}"
            ) from exc
        if src_int <= 0:
            raise ValueError(
                f"Profile '{profile_path}' source_layout.{source_key} must be >= 1, got: {src_int!r}"
            )
        return src_int

    resolved_worksheet_name = _shell_or_source_str("source_worksheet_name", "worksheet_name")
    resolved_worksheet_resolution = _shell_or_source_str(
        "source_worksheet_resolution", "worksheet_resolution"
    )
    resolved_header_row = _shell_or_source_int("source_header_row", "header_row")
    resolved_data_start_row = _shell_or_source_int("source_data_start_row", "data_start_row")

    return TransformProfile(
        profile_key=_profile_section_value(profile_path, transform, "transform", "key"),
        profile_label=_profile_section_value(profile_path, transform, "transform", "label"),
        source_type=_profile_section_value(profile_path, transform, "transform", "source_type"),
        source_data=source_data_val,
        target_eib=_profile_section_value(
            profile_path, shell_defaults, "shell_defaults", "target_eib"
        ),
        output_file_path=output_val,
        mapping_spreadsheet=str(shell_defaults.get("mapping_spreadsheet", "") or "")
        or reference_paths.get("mapping_spreadsheet", ""),
        mapping_sheet=str(shell_defaults.get("mapping_sheet", "") or "")
        or reference_paths.get("mapping_sheet", ""),
        customer_spreadsheet=str(shell_defaults.get("customer_spreadsheet", "") or "")
        or reference_paths.get("customer_spreadsheet", ""),
        supplier_spreadsheet=str(shell_defaults.get("supplier_spreadsheet", "") or "")
        or reference_paths.get("supplier_spreadsheet", ""),
        workday_reference_ids=str(shell_defaults.get("workday_reference_ids", "") or "")
        or reference_paths.get("workday_reference_ids", ""),
        eib_population_type=_profile_section_value(
            profile_path, transform, "transform", "eib_population_type"
        ),
        step_list_macro=_profile_section_value(
            profile_path, runtime, "runtime", "step_list_macro"
        ),
        start_row=_optional_profile_int(profile_path, transform, "transform", "start_row"),
        worksheet_name=resolved_worksheet_name,
        worksheet_resolution=resolved_worksheet_resolution,
        header_row=resolved_header_row,
        data_start_row=resolved_data_start_row,
    )


def _build_control_rows(profile: dict) -> list[tuple[str, str]]:
    values = [
        _profile_value(profile, "transform", "label"),
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
    ]
    return list(zip(CONTROL_LABELS, [str(value) for value in values], strict=True))


def _profile_value(profile: dict, section: str, key: str) -> str:
    section_data = profile.get(section, {})
    value = section_data.get(key, "")
    if value in ("", None):
        raise ValueError(f"Profile '{section}.{key}' must be populated.")
    return str(value)


def _profile_section_value(
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


def _import_pyopenvba():
    try:
        from pyopenvba import ExcelFile, VBAModuleKind
    except ImportError as exc:  # pragma: no cover - dependency boundary
        raise RuntimeError(
            "pyOpenVBA is required. Run with:\n"
            "uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python "
            "elt_transform_excel/scripts/create_runner_shell.py"
        ) from exc
    return ExcelFile, VBAModuleKind


def _import_openpyxl():
    try:
        from openpyxl import load_workbook
        from openpyxl.worksheet.datavalidation import DataValidation
    except ImportError as exc:  # pragma: no cover - dependency boundary
        raise RuntimeError(
            "openpyxl is required. Run with:\n"
            "uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python "
            "elt_transform_excel/scripts/create_runner_shell.py"
        ) from exc
    return load_workbook, DataValidation


def _import_yaml():
    try:
        import yaml
    except ImportError as exc:  # pragma: no cover - dependency boundary
        raise RuntimeError(
            "PyYAML is required. Run with:\n"
            "uv run --with pyOpenVBA --with openpyxl --with pyyaml --no-project python "
            "elt_transform_excel/scripts/create_runner_shell.py"
        ) from exc
    return yaml


def _write_layout(worksheet, control_rows: list[tuple[str, str]], profile: dict) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill

    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    action_fill = PatternFill(fill_type="solid", fgColor="D9EAD3")
    subhead_fill = PatternFill(fill_type="solid", fgColor="F3F3F3")

    worksheet.sheet_view.showGridLines = True
    worksheet.freeze_panes = "A3"
    worksheet.column_dimensions["A"].width = 24
    worksheet.column_dimensions["B"].width = 90
    worksheet.column_dimensions["D"].width = 20
    worksheet.column_dimensions["E"].width = 40

    worksheet["A1"] = TITLE
    worksheet["A1"].font = Font(bold=True, size=14)
    worksheet["A2"] = "Generated by"
    worksheet["B2"] = "scripts/create_runner_shell.py"
    worksheet["D2"] = "Selected Transform"
    worksheet["E2"] = _profile_value(profile, "transform", "label")
    worksheet["D2"].font = Font(bold=True)

    for row_index, (label, value) in enumerate(control_rows, start=3):
        worksheet[f"A{row_index}"] = label
        worksheet[f"B{row_index}"] = value
        worksheet[f"A{row_index}"].font = Font(bold=True)

    worksheet["D3"] = "Run Action"
    worksheet["E3"] = "Click the 'Run Transform' button below"
    worksheet["D4"] = "Macro Name"
    worksheet["E4"] = "modDispatcher.RunSelectedTransform"
    worksheet["D5"] = "Note"
    worksheet["E5"] = INSTRUCTION_TEXT
    worksheet["D6"] = "Profile Key"
    worksheet["E6"] = _profile_value(profile, "transform", "key")
    for cell_ref in ("D3", "D4", "D5", "D6"):
        worksheet[cell_ref].font = Font(bold=True)
    for cell_ref in ("D3", "E3", "D4", "E4", "D5", "E5", "D6", "E6"):
        worksheet[cell_ref].fill = action_fill
        worksheet[cell_ref].alignment = Alignment(vertical="top", wrap_text=True)

    result_heading_row = 3 + len(control_rows) + 1
    worksheet[f"A{result_heading_row}"] = "Results"
    worksheet[f"A{result_heading_row}"].font = Font(bold=True)
    worksheet[f"A{result_heading_row}"].fill = subhead_fill

    for row_index, (label, value) in enumerate(RESULT_ROWS, start=result_heading_row + 1):
        worksheet[f"A{row_index}"] = label
        worksheet[f"B{row_index}"] = value
        worksheet[f"A{row_index}"].font = Font(bold=True)

    for row_index in range(3, 3 + len(control_rows)):
        worksheet[f"A{row_index}"].fill = header_fill
    for row_index in range(result_heading_row + 1, result_heading_row + 1 + len(RESULT_ROWS)):
        worksheet[f"A{row_index}"].fill = subhead_fill

    for row_index in range(2, result_heading_row + len(RESULT_ROWS) + 1):
        worksheet[f"B{row_index}"].alignment = Alignment(vertical="top", wrap_text=True)
        worksheet[f"E{row_index}"].alignment = Alignment(vertical="top", wrap_text=True)


def _write_transform_registry(worksheet, transform_profiles: list[TransformProfile]) -> None:
    max_row = max(worksheet.max_row, REGISTRY_START_ROW + len(transform_profiles) + 5)
    for row_index in range(REGISTRY_TITLE_ROW, max_row + 1):
        for column_index in range(REGISTRY_START_COLUMN, REGISTRY_END_COLUMN + 1):
            worksheet.cell(row=row_index, column=column_index).value = None

    worksheet.cell(row=REGISTRY_TITLE_ROW, column=REGISTRY_START_COLUMN).value = "Transform Registry"
    headers = [
        "Transform Label",
        "Profile Key",
        "Source Type",
        "Source Data",
        "Target EIB",
        "Output File Path",
        "Mapping Spreadsheet",
        "Mapping Sheet",
        "Customer Spreadsheet",
        "Supplier Spreadsheet",
        "Workday Reference IDs",
        "EIB Population Type",
        "Step List Macro",
        "Start Row",
    ]
    for offset, header in enumerate(headers):
        worksheet.cell(row=REGISTRY_HEADER_ROW, column=REGISTRY_START_COLUMN + offset).value = header

    for row_offset, profile in enumerate(transform_profiles):
        row_index = REGISTRY_START_ROW + row_offset
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN).value = profile.profile_label
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 1).value = profile.profile_key
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 2).value = profile.source_type
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 3).value = profile.source_data
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 4).value = profile.target_eib
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 5).value = profile.output_file_path
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 6).value = profile.mapping_spreadsheet
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 7).value = profile.mapping_sheet
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 8).value = profile.customer_spreadsheet
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 9).value = profile.supplier_spreadsheet
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 10).value = profile.workday_reference_ids
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 11).value = profile.eib_population_type
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 12).value = profile.step_list_macro
        worksheet.cell(row=row_index, column=REGISTRY_START_COLUMN + 13).value = profile.start_row


def _write_shell_defaults_config(worksheet, transform_profiles: list[TransformProfile]) -> None:
    from openpyxl.styles import Alignment, Font, PatternFill

    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    subhead_fill = PatternFill(fill_type="solid", fgColor="F3F3F3")

    worksheet.sheet_view.showGridLines = True
    worksheet.freeze_panes = f"A{CONFIG_START_ROW}"
    worksheet.column_dimensions["A"].width = 24
    worksheet.column_dimensions["B"].width = 22
    worksheet.column_dimensions["C"].width = 90
    worksheet.column_dimensions["D"].width = 90
    worksheet.column_dimensions["E"].width = 90
    worksheet.column_dimensions["F"].width = 90
    worksheet.column_dimensions["G"].width = 90
    worksheet.column_dimensions["H"].width = 90
    worksheet.column_dimensions["I"].width = 90
    worksheet.column_dimensions["J"].width = 90

    worksheet[f"A{CONFIG_TITLE_ROW}"] = "Shell Defaults Config"
    worksheet[f"A{CONFIG_TITLE_ROW}"].font = Font(bold=True, size=14)
    worksheet[f"B{CONFIG_TITLE_ROW}"] = (
        "Business-editable machine-specific values used to populate the Runner sheet."
    )
    worksheet[f"B{CONFIG_TITLE_ROW}"].alignment = Alignment(vertical="top", wrap_text=True)

    headers = [
        "Transform Label",
        "Profile Key",
        "Source Data",
        "Target EIB",
        "Output File Path",
        "Mapping Spreadsheet",
        "Mapping Sheet",
        "Customer Spreadsheet",
        "Supplier Spreadsheet",
        "Workday Reference IDs",
    ]
    for column_index, header in enumerate(headers, start=1):
        worksheet.cell(row=CONFIG_HEADER_ROW, column=column_index).value = header
        worksheet.cell(row=CONFIG_HEADER_ROW, column=column_index).font = Font(bold=True)
        worksheet.cell(row=CONFIG_HEADER_ROW, column=column_index).fill = header_fill

    for row_offset, profile in enumerate(transform_profiles):
        row_index = CONFIG_START_ROW + row_offset
        worksheet.cell(row=row_index, column=1).value = profile.profile_label
        worksheet.cell(row=row_index, column=2).value = profile.profile_key
        worksheet.cell(row=row_index, column=3).value = profile.source_data
        worksheet.cell(row=row_index, column=4).value = profile.target_eib
        worksheet.cell(row=row_index, column=5).value = profile.output_file_path
        worksheet.cell(row=row_index, column=6).value = profile.mapping_spreadsheet
        worksheet.cell(row=row_index, column=7).value = profile.mapping_sheet
        worksheet.cell(row=row_index, column=8).value = profile.customer_spreadsheet
        worksheet.cell(row=row_index, column=9).value = profile.supplier_spreadsheet
        worksheet.cell(row=row_index, column=10).value = profile.workday_reference_ids

    for row_index in range(CONFIG_START_ROW, CONFIG_START_ROW + len(transform_profiles)):
        for column_index in range(1, len(headers) + 1):
            worksheet.cell(row=row_index, column=column_index).alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

    worksheet[f"A{CONFIG_HEADER_ROW}"].fill = subhead_fill


def _apply_profile_formulas(
    worksheet,
    active_profile: dict,
    profile_count: int,
) -> None:
    label_row = CONTROL_ROWS_START + CONTROL_LABELS.index("Transform")
    lookup_rows = range(REGISTRY_START_ROW, REGISTRY_START_ROW + profile_count)
    lookup_start = min(lookup_rows)
    lookup_end = max(lookup_rows)

    def formula_for(registry_column_label: str) -> str:
        return (
            '=IFERROR(INDEX($'
            f"{registry_column_label}${lookup_start}:${registry_column_label}${lookup_end}"
            ',MATCH($B$'
            f"{label_row}"
            ',$G$'
            f"{lookup_start}"
            ':$G$'
            f"{lookup_end}"
            ',0)),"")'
        )

    def config_formula_for(config_column_label: str) -> str:
        return (
            '=IFERROR(INDEX('
            f"'{CONFIG_SHEET_NAME}'!${config_column_label}${CONFIG_START_ROW}:${config_column_label}$"
            f"{CONFIG_START_ROW + profile_count - 1}"
            ',MATCH($B$'
            f"{label_row}"
            f",'{CONFIG_SHEET_NAME}'!$A${CONFIG_START_ROW}:$A$"
            f"{CONFIG_START_ROW + profile_count - 1}"
            ',0)),"")'
        )

    row_by_label = {
        label: CONTROL_ROWS_START + index
        for index, label in enumerate(CONTROL_LABELS)
    }
    worksheet[f"B{row_by_label['Source Type']}"] = formula_for("I")
    worksheet[f"B{row_by_label['Source Data']}"] = config_formula_for("C")
    worksheet[f"B{row_by_label['Target EIB']}"] = config_formula_for("D")
    worksheet[f"B{row_by_label['Output File Path']}"] = config_formula_for("E")
    worksheet[f"B{row_by_label['Mapping Spreadsheet']}"] = config_formula_for("F")
    worksheet[f"B{row_by_label['Mapping Sheet']}"] = config_formula_for("G")
    worksheet[f"B{row_by_label['Customer Spreadsheet']}"] = config_formula_for("H")
    worksheet[f"B{row_by_label['Supplier Spreadsheet']}"] = config_formula_for("I")
    worksheet[f"B{row_by_label['Workday Reference IDs']}"] = config_formula_for("J")
    worksheet[f"B{row_by_label['EIB Population Type']}"] = formula_for("R")
    worksheet["E2"] = f"=B{label_row}"
    worksheet["E6"] = formula_for("H")
    worksheet[f"B{label_row}"] = _profile_value(active_profile, "transform", "label")


def _add_dropdown_validation(
    worksheet,
    data_validation_cls,
    transform_label_options: list[str],
) -> None:
    transform_validation = data_validation_cls(
        type="list",
        formula1=_to_excel_list_formula(transform_label_options),
        allow_blank=False,
    )

    worksheet.add_data_validation(transform_validation)
    transform_row = CONTROL_ROWS_START + CONTROL_LABELS.index("Transform")
    transform_validation.add(worksheet[f"B{transform_row}"])


def _to_excel_list_formula(options: list[str]) -> str:
    escaped = [option.replace('"', '""') for option in options]
    return '"' + ",".join(escaped) + '"'


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # pragma: no cover - CLI error boundary
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
