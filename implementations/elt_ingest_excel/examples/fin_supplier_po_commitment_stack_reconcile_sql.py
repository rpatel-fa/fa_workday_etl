"""End-to-end reconcile runner for the Commitment Stack → Workday PO EIB pipeline.

Acknowledges the current architecture: the active Workday PO EIB is now built from
the Commitment Stack Report (Proactis) — not the legacy split PO/receipt tabs.

Prerequisite: the commitment stack pipeline has been run to TRANSFORM so these
tables exist:
  src_fin_supplier_po_commitment_stack_{base,nrm}
  src_fin_supplier_po_commitment_stack (+ ref)
  workday_supplier_po_commitment_stack
  fin_supplier_po_audit_commitment  (for the audit reconcile leg)

Usage:
  # Run reconcile only against the current DuckDB state, then print summary
  uv run python examples/fin_supplier_po_commitment_stack_reconcile_sql.py

  # Full flow: refresh commitment stack pipeline + run reconcile
  uv run python examples/fin_supplier_purchase_order_receipt.py --run-to-phase TRANSFORM
  uv run python examples/fin_supplier_po_commitment_stack_reconcile_sql.py

  # Export the review pack xlsx for finance stakeholders
  uv run python examples/fin_supplier_po_commitment_stack_reconcile_sql.py --export-review-workbook
"""

import json
from pathlib import Path
from datetime import datetime, timezone

import duckdb
import pandas as pd
from base_runner import create_parser
from elt_ingest_excel.transform.sql_validation_dir import run_validation_sql_dir
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

RECONCILE_SQL_DIR = "config/reconcile/finance/supplier_purchase_order_receipt"
REVIEW_EXPORT_CONFIG = (
    "config/publish/finance/publish_supplier_po_commitment_stack_reconcile_review.json"
)


def _resolve_database_path(database_path, runtime_config_name):
    if database_path:
        return Path(database_path).expanduser().resolve()

    runtime_config_file = (
        Path(__file__).parent.parent / "config" / "runtime" / runtime_config_name
    ).resolve()
    runtime_cfg = json.loads(runtime_config_file.read_text())
    if isinstance(runtime_cfg, list):
        runtime_cfg = runtime_cfg[0] if runtime_cfg else {}

    for key in ("databasePath", "database_path", "duckdbPath", "duckdb_path"):
        value = runtime_cfg.get(key)
        if value:
            return Path(value).expanduser().resolve()

    raise ValueError(
        "No DuckDB database path. Pass --database-path or configure in "
        f"{runtime_config_file}"
    )


def _resolve_export_config_path(export_config_path):
    path = Path(export_config_path).expanduser()
    if path.is_absolute() and path.exists():
        return path.resolve()
    if path.exists():
        return path.resolve()
    candidate = (Path(__file__).parent.parent / path).resolve()
    if candidate.exists():
        return candidate
    raise FileNotFoundError(
        f"Review export config not found. Tried {path.resolve()} and {candidate}"
    )


def _export_review_workbook(*, database_path, export_config_path, export_dir):
    export_cfg = json.loads(export_config_path.read_text())
    if not isinstance(export_cfg, list):
        raise ValueError(
            f"Invalid export config (expected list): {export_config_path}"
        )

    exported_paths = []
    created = datetime(2000, 1, 1, tzinfo=timezone.utc)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(
        start_color="4F81BD", end_color="4F81BD", fill_type="solid"
    )

    with duckdb.connect(str(database_path), read_only=True) as conn:
        for workbook_cfg in export_cfg:
            if not isinstance(workbook_cfg, dict):
                raise ValueError(
                    f"Invalid workbook entry in export config: {export_config_path}"
                )
            configured_dir = Path(workbook_cfg["tgtWorkbookPathName"]).expanduser()
            tgt_dir = export_dir or configured_dir
            tgt_name = workbook_cfg["tgtWorkbookFileName"]
            tgt_path = (tgt_dir / f"{tgt_name}.xlsx").expanduser().resolve()

            wb = Workbook()
            wb.properties.created = created
            wb.properties.modified = created
            if "Sheet" in wb.sheetnames:
                del wb["Sheet"]

            for sheet_cfg in workbook_cfg.get("sheets", []):
                sheet_name = sheet_cfg["sheetName"]
                if len(sheet_name) > 31:
                    raise ValueError(
                        f"Sheet name too long ({len(sheet_name)} chars, max 31): {sheet_name!r}"
                    )
                table_name = sheet_cfg["srcTableName"]
                try:
                    df = conn.execute(f'SELECT * FROM "{table_name}"').df()
                except Exception as e:
                    print(f"  [warn] Skipping sheet {sheet_name!r}: {e}")
                    continue
                df = df.astype(object).where(pd.notna(df), None)

                ws = wb.create_sheet(title=sheet_name)
                ws.append(list(df.columns))
                for row in df.itertuples(index=False, name=None):
                    ws.append(row)

                for cell in ws[1]:
                    cell.font = header_font
                    cell.fill = header_fill

                ws.auto_filter.ref = ws.dimensions
                ws.freeze_panes = "A2"

                for column_cells in ws.columns:
                    max_length = 0
                    for cell in column_cells:
                        if cell.value is None:
                            continue
                        max_length = max(max_length, len(str(cell.value)))
                    column_letter = get_column_letter(column_cells[0].column)
                    ws.column_dimensions[column_letter].width = min(
                        max_length + 2, 100
                    )

            tgt_dir.mkdir(parents=True, exist_ok=True)
            wb.save(tgt_path)
            exported_paths.append(tgt_path)

    return exported_paths


def _print_reconcile_summary(database_path):
    with duckdb.connect(str(database_path), read_only=True) as conn:
        board = conn.execute(
            "SELECT sort_order, headline, headline_value "
            "FROM reconcile_cs_board_summary ORDER BY sort_order"
        ).fetchall()
        report = conn.execute(
            "SELECT section, metric, source_value, target_value, difference, status "
            "FROM reconcile_cs_report ORDER BY section, metric"
        ).fetchall()
        issues = conn.execute(
            "SELECT table_name, rule_type, recon_key, detail, message "
            "FROM reconciliation_supplier_po_receipt_issues"
        ).fetchall()

    print()
    print("=================================================================")
    print(" Commitment Stack → Workday PO EIB Reconcile (INTERIM)")
    print("=================================================================")
    for _, headline, value in board:
        if value is not None:
            print(f"  {headline:<38s} {value}")
    print()
    print("--- Source → Target report ---")
    for section, metric, src, tgt, diff, status in report:
        status_suffix = f"  [{status}]" if status and status != "INFORMATIONAL" else ""
        if src is not None and tgt is not None:
            print(
                f"  {section:<20s} {metric:<40s} "
                f"src={src} tgt={tgt} diff={diff}{status_suffix}"
            )
        elif src is not None:
            print(
                f"  {section:<20s} {metric:<40s} src={src} tgt=- diff={diff}{status_suffix}"
            )
        else:
            print(
                f"  {section:<20s} {metric:<40s} src=- tgt={tgt} diff={diff}{status_suffix}"
            )
    if issues:
        print()
        print(f"--- ISSUES ({len(issues)}) ---")
        for table_name, rule_type, recon_key, detail, message in issues:
            print(
                f"  * [{table_name}] {rule_type} {recon_key} "
                f"detail={detail} — {message}"
            )
    else:
        print()
        print("No issues (0 rows in reconciliation_supplier_po_receipt_issues).")
    print()


if __name__ == "__main__":
    parser = create_parser(
        "Run Commitment Stack → Workday PO EIB reconciliation",
        default_run_to_phase="TRANSFORM",
    )
    parser.add_argument(
        "--export-review-workbook",
        action="store_true",
        help="Export the reconcile review tables as a finance-facing .xlsx workbook",
    )
    parser.add_argument(
        "--export-review-workbook-only",
        action="store_true",
        help="Export the review workbook only; skip reconcile SQL rerun",
    )
    parser.add_argument(
        "--export-review-config",
        type=str,
        default=REVIEW_EXPORT_CONFIG,
        help="Override path to the review export JSON config",
    )
    parser.add_argument(
        "--export-review-dir",
        type=str,
        default=None,
        help="Override output directory for the review workbook",
    )
    args = parser.parse_args()

    database_path = _resolve_database_path(
        database_path=args.database_path,
        runtime_config_name=args.runtime_config,
    )

    if args.export_review_workbook_only:
        _print_reconcile_summary(database_path)
        paths = _export_review_workbook(
            database_path=database_path,
            export_config_path=_resolve_export_config_path(
                args.export_review_config
            ),
            export_dir=Path(args.export_review_dir).expanduser().resolve()
            if args.export_review_dir
            else None,
        )
        for p in paths:
            print(f"Exported review workbook: {p}")
        raise SystemExit(0)

    reconcile_sql_dir = (
        Path(__file__).parent.parent / RECONCILE_SQL_DIR
    ).resolve()
    print(f"Running reconcile SQL from {reconcile_sql_dir} against {database_path}")
    run_validation_sql_dir(
        sql_dir=str(reconcile_sql_dir),
        cfg_runtime_name=args.runtime_config,
        database_path=args.database_path,
    )
    _print_reconcile_summary(database_path)

    if args.export_review_workbook:
        paths = _export_review_workbook(
            database_path=database_path,
            export_config_path=_resolve_export_config_path(
                args.export_review_config
            ),
            export_dir=Path(args.export_review_dir).expanduser().resolve()
            if args.export_review_dir
            else None,
        )
        for p in paths:
            print(f"Exported review workbook: {p}")
