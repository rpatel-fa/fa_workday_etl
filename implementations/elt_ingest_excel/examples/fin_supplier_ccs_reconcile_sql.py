"""End-to-end runner for Finance supplier GP -> Workday reconciliation.

Refreshes the Workday supplier extract, loads the go-live manual-update list,
runs the supplier_ccs SQL reconciliation, and (optionally) exports an
analyst-friendly review-pack .xlsx — the supplier analogue of the customer CCS
reconcile pack.
"""

from pathlib import Path

import duckdb
import openpyxl
from base_runner import create_parser, run_pipeline
from elt_ingest_excel import PipelinePhase, SaveMode
from elt_ingest_excel.transform.sql_validation_dir import run_validation_sql_dir

# Reuse the generic review-workbook exporter from the customer CCS runner rather
# than duplicating it (code-reuse gate).
from fin_customer_ccs_reconcile_sql import (
    _export_review_workbook,
    _resolve_database_path,
    _resolve_export_config_path,
)
from supplier_ccs_summary_sheet import build_summary_sheet


def _export_pack_with_summary(*, database_path, export_config_path, export_dir):
    """Export the review workbook, then prepend the executive summary tab."""
    paths = _export_review_workbook(
        database_path=database_path,
        export_config_path=export_config_path,
        export_dir=export_dir,
    )
    for path in paths:
        build_summary_sheet(path, database_path)
    return paths

EXTRACT_INGEST_CONFIG = "ingest/finance/workday_supplier_extract.json"
RECONCILE_SQL_DIR = "config/reconcile/finance/supplier_ccs"
REVIEW_EXPORT_CONFIG = "config/publish/finance/publish_supplier_ccs_reconcile_review.json"
DEFAULT_MANUAL_UPDATE_FILE = "~/Downloads/Suppliers to be manually updated at Go Live.xlsx"

SUMMARY_METRICS = [
    "1a_distinct_gp_vendor_id",
    "1c_fin_workday_extract_supplier_rows",
    "2a_gp_missing_in_workday",
    "2b_extra_workday_reference_ids_not_in_gp",
    "2c_workday_dummy_rows",
    "3a_manual_update_total",
    "3e_manual_update_unmatched",
    "4b_address_gap_unexplained",
    "4d_bank_gap_unexplained",
    "4f_payment_gap_unexplained",
]


def _load_manual_update_table(database_path: Path, manual_update_file: Path) -> int:
    """Parse the name-only, bucket-structured manual-update workbook into
    fin_supplier_manual_update (seq, bucket, supplier_name_raw).

    Section headers ("REMOVED ADDRESS", "REMOVED BANK ...", "REMOVED ALL
    PAYMENT ...") are forward-filled onto the supplier names that follow them.
    """
    wb = openpyxl.load_workbook(manual_update_file, data_only=True)
    ws = wb.active
    rows: list[tuple[int, str, str]] = []
    current_bucket: str | None = None
    seq = 0
    for row in ws.iter_rows(values_only=True):
        value = row[0]
        if value is None or not str(value).strip():
            continue
        text = str(value).strip()
        if text.upper().startswith("REMOVED"):
            current_bucket = text
            continue
        seq += 1
        rows.append((seq, current_bucket, text))
    wb.close()

    with duckdb.connect(str(database_path)) as conn:
        conn.execute("DROP TABLE IF EXISTS fin_supplier_manual_update")
        conn.execute(
            "CREATE TABLE fin_supplier_manual_update "
            "(seq INTEGER, bucket VARCHAR, supplier_name_raw VARCHAR)"
        )
        conn.executemany(
            "INSERT INTO fin_supplier_manual_update VALUES (?, ?, ?)", rows
        )
    return len(rows)


def _print_reconcile_summary(database_path: Path) -> None:
    with duckdb.connect(str(database_path), read_only=True) as conn:
        rows = conn.execute(
            """
            SELECT section, metric, metric_value, status
              FROM reconcile_supplier_ccs_report
             ORDER BY section, metric
            """
        ).fetchall()

    seen = {metric for _, metric, _, _ in rows}
    missing = [m for m in SUMMARY_METRICS if m not in seen]
    if missing:
        raise RuntimeError(
            "Supplier CCS reconcile summary is incomplete. Missing metrics: "
            + ", ".join(missing)
        )

    print()
    print("Supplier CCS Reconcile Summary")
    print("------------------------------")
    for _, metric, metric_value, status in rows:
        status_suffix = f"  [{status}]" if status else ""
        print(f"{metric}: {metric_value}{status_suffix}")
    print()
    print("Review tables:")
    print("  - reconcile_supplier_ccs_board_summary")
    print("  - reconcile_supplier_ccs_report")
    print("  - reconcile_supplier_ccs_closure_summary")
    print("  - reconcile_supplier_ccs_manual_update")
    print("  - reconcile_supplier_ccs_attribute_gap")
    print("  - reconcile_supplier_ccs_workday_dummy_row")


if __name__ == "__main__":
    parser = create_parser(
        "Run Finance supplier GP -> Workday extract refresh plus SQL reconciliation",
        default_run_to_phase="TRANSFORM",
    )
    parser.add_argument(
        "--manual-update-file",
        type=str,
        default=DEFAULT_MANUAL_UPDATE_FILE,
        help="Path to the go-live manual-update supplier list workbook",
    )
    parser.add_argument(
        "--export-review-workbook",
        action="store_true",
        help="Export the supplier CCS reconcile review tables into an .xlsx workbook",
    )
    parser.add_argument(
        "--export-review-workbook-only",
        action="store_true",
        help="Export the review workbook from existing DuckDB state (skips ingest/reconcile)",
    )
    parser.add_argument(
        "--export-review-config",
        type=str,
        default=REVIEW_EXPORT_CONFIG,
        help="Path to the review export config JSON",
    )
    parser.add_argument(
        "--export-review-dir",
        type=str,
        default=None,
        help="Override output directory for the review workbook export",
    )
    args = parser.parse_args()

    run_to_phase = PipelinePhase[args.run_to_phase]
    database_path = _resolve_database_path(
        database_path=args.database_path,
        runtime_config_name=args.runtime_config,
    )

    if args.export_review_workbook_only:
        _print_reconcile_summary(database_path)
        for path in _export_pack_with_summary(
            database_path=database_path,
            export_config_path=_resolve_export_config_path(args.export_review_config),
            export_dir=Path(args.export_review_dir).expanduser().resolve()
            if args.export_review_dir
            else None,
        ):
            print(f"Exported review workbook: {path}")
        raise SystemExit(0)

    # 1. Refresh the Workday supplier extract table from the current extract file.
    run_pipeline(
        cfg_ingest_path="ingest/finance",
        cfg_ingest_name="workday_supplier_extract.json",
        cfg_transform_path="",
        cfg_publish_path="",
        cfg_publish_name="",
        cfg_runtime_name=args.runtime_config,
        database_path=args.database_path,
        run_to_phase=PipelinePhase.INGEST,
        save_mode=SaveMode.RECREATE,
    )

    if run_to_phase == PipelinePhase.INGEST:
        raise SystemExit(0)

    # 2. Load the go-live manual-update list.
    manual_update_file = Path(args.manual_update_file).expanduser().resolve()
    manual_rows = _load_manual_update_table(database_path, manual_update_file)
    print(f"Loaded manual-update list: {manual_rows} supplier names")

    # 3. Run the SQL reconciliation chain.
    run_validation_sql_dir(
        sql_dir=RECONCILE_SQL_DIR,
        cfg_runtime_name=args.runtime_config,
        database_path=args.database_path,
    )
    _print_reconcile_summary(database_path)

    # 4. Optionally export the review pack.
    if args.export_review_workbook:
        for path in _export_pack_with_summary(
            database_path=database_path,
            export_config_path=_resolve_export_config_path(args.export_review_config),
            export_dir=Path(args.export_review_dir).expanduser().resolve()
            if args.export_review_dir
            else None,
        ):
            print(f"Exported review workbook: {path}")
