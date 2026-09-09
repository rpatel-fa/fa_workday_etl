"""End-to-end runner for Finance CCS customer extract plus reconciliation."""

import json
from pathlib import Path
from datetime import datetime, timezone

import duckdb
import pandas as pd
from base_runner import create_parser, run_pipeline
from elt_ingest_excel import PipelinePhase
from elt_ingest_excel.transform.sql_validation_dir import run_validation_sql_dir
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

RECONCILE_SQL_DIR = "elt_ingest_excel/config/reconcile/finance/customer_ccs"
REVIEW_EXPORT_CONFIG = (
    "config/publish/finance/publish_customer_ccs_reconcile_review.json"
)


def _resolve_database_path(database_path: str | None, runtime_config_name: str) -> Path:
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
        "No DuckDB database path provided. "
        f"Pass --database-path or configure it in {runtime_config_file}"
    )


def _resolve_export_config_path(export_config_path: str) -> Path:
    path = Path(export_config_path).expanduser()
    if path.is_absolute() and path.exists():
        return path.resolve()

    if path.exists():
        return path.resolve()

    candidate = (Path(__file__).parent.parent / path).resolve()
    if candidate.exists():
        return candidate

    raise FileNotFoundError(
        "Review export config file not found. Tried: "
        f"{path.resolve()} and {candidate}"
    )


def _export_review_workbook(
    *,
    database_path: Path,
    export_config_path: Path,
    export_dir: Path | None,
) -> list[Path]:
    export_cfg = json.loads(export_config_path.read_text())
    if not isinstance(export_cfg, list):
        raise ValueError(
            "Invalid export config format (expected list): " f"{export_config_path}"
        )

    exported_paths: list[Path] = []

    created = datetime(2000, 1, 1, tzinfo=timezone.utc)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(
        start_color="4F81BD", end_color="4F81BD", fill_type="solid"
    )

    with duckdb.connect(str(database_path), read_only=True) as conn:
        for workbook_cfg in export_cfg:
            if not isinstance(workbook_cfg, dict):
                raise ValueError(
                    "Invalid workbook config entry (expected object): "
                    f"{export_config_path}"
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
                        "Excel sheet names must be <= 31 chars. "
                        f"Got {len(sheet_name)} for {sheet_name!r} "
                        f"in {export_config_path}"
                    )

                table_name = sheet_cfg["srcTableName"]
                df = conn.execute(f'SELECT * FROM "{table_name}"').df()
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
                    ws.column_dimensions[column_letter].width = min(max_length + 2, 100)

            tgt_dir.mkdir(parents=True, exist_ok=True)
            wb.save(tgt_path)
            exported_paths.append(tgt_path)

    return exported_paths



def _print_reconcile_summary(database_path: Path) -> None:
    summary_metrics = [
        "1b_distinct_gp_customer_number",
        "1e_src_fin_workday_extract_ccs_customer_row_count",
        "2b_source_customer_numbers_missing_in_workday",
        "2c_extra_workday_reference_ids_not_in_source",
        "4a_matched_customer_ids",
        "4a_tax_status_country_diff_on_matched_customer_ids",
        "4a_transaction_tax_status_diff_on_matched_customer_ids",
        "5a_provenance_case_rows",
    ]

    with duckdb.connect(str(database_path), read_only=True) as conn:
        rows = conn.execute(
            """
            SELECT section, metric, metric_value, status
              FROM reconcile_customer_ccs_report
             WHERE metric IN (
                   '1b_distinct_gp_customer_number'
                 , '1c_src_fin_customer_row_count'
                 , '1e_src_fin_workday_extract_ccs_customer_row_count'
                 , '2b_source_customer_numbers_missing_in_workday'
                 , '2c_extra_workday_reference_ids_not_in_source'
                 , '4a_matched_customer_ids'
                 , '4a_tax_status_country_diff_on_matched_customer_ids'
                 , '4a_transaction_tax_status_diff_on_matched_customer_ids'
                 , '5a_provenance_case_rows'
                 )
             ORDER BY
                   CASE metric
                     WHEN '1b_distinct_gp_customer_number' THEN 1
                     WHEN '1c_src_fin_customer_row_count' THEN 2
                     WHEN '1e_src_fin_workday_extract_ccs_customer_row_count' THEN 3
                     WHEN '2b_source_customer_numbers_missing_in_workday' THEN 4
                     WHEN '2c_extra_workday_reference_ids_not_in_source' THEN 5
                     WHEN '4a_matched_customer_ids' THEN 6
                     WHEN '4a_tax_status_country_diff_on_matched_customer_ids' THEN 7
                     WHEN '4a_transaction_tax_status_diff_on_matched_customer_ids' THEN 8
                     WHEN '5a_provenance_case_rows' THEN 9
                   END
            """
        ).fetchall()

    seen_metrics = {metric for _, metric, _, _ in rows}
    missing_metrics = [metric for metric in summary_metrics if metric not in seen_metrics]
    if missing_metrics:
        raise RuntimeError(
            "CCS reconcile summary is incomplete. Missing metrics: "
            + ", ".join(missing_metrics)
        )

    print()
    print("CCS Reconcile Summary")
    print("---------------------")
    for _, metric, metric_value, status in rows:
        status_suffix = f"  [{status}]" if status else ""
        print(f"{metric}: {metric_value}{status_suffix}")
    print()
    print("Review tables:")
    print("  - reconcile_customer_ccs_provenance_case_summary")
    print("  - reconcile_customer_ccs_contact_summary")
    print("  - reconcile_customer_ccs_contact_review_case_summary")
    print("  - reconcile_customer_ccs_contact_review_pattern_summary")
    print("  - reconcile_customer_ccs_tax_review_operator_pack")
    print("  - reconcile_customer_ccs_closure_summary")
    print("  - reconcile_customer_ccs_closure_decision_pack")
    print("  - reconcile_customer_ccs_report")

if __name__ == "__main__":
    parser = create_parser(
        "Run Finance CCS customer extract refresh plus SQL reconciliation",
        default_run_to_phase="TRANSFORM",
    )
    parser.add_argument(
        "--export-review-workbook",
        action="store_true",
        help="Export the key CCS reconcile review tables into an analyst-friendly .xlsx workbook",
    )
    parser.add_argument(
        "--export-review-workbook-only",
        action="store_true",
        help="Export the review workbook from an existing DuckDB state (skips ingest/transform/reconcile)",
    )
    parser.add_argument(
        "--export-review-config",
        type=str,
        default=REVIEW_EXPORT_CONFIG,
        help="Path to the review export config JSON (defaults to repo config/publish/...)",
    )
    parser.add_argument(
        "--export-review-dir",
        type=str,
        default=None,
        help="Override output directory for the review workbook export (defaults to config value)",
    )
    args = parser.parse_args()

    run_to_phase = PipelinePhase[args.run_to_phase]

    if args.export_review_workbook_only:
        database_path = _resolve_database_path(
            database_path=args.database_path,
            runtime_config_name=args.runtime_config,
        )
        _print_reconcile_summary(database_path)
        export_paths = _export_review_workbook(
            database_path=database_path,
            export_config_path=_resolve_export_config_path(args.export_review_config),
            export_dir=Path(args.export_review_dir).expanduser().resolve()
            if args.export_review_dir
            else None,
        )
        for path in export_paths:
            print(f"Exported review workbook: {path}")
        raise SystemExit(0)

    run_pipeline(
        cfg_ingest_path="ingest/finance",
        cfg_ingest_name="workday_ccs_customer.json",
        cfg_transform_path="transform/sql/finance/workday_extract_ccs_customer",
        cfg_publish_path="",
        cfg_publish_name="",
        cfg_runtime_name=args.runtime_config,
        database_path=args.database_path,
        run_to_phase=run_to_phase,
    )

    if run_to_phase != PipelinePhase.INGEST:
        reconcile_sql_dir = (
            Path(__file__).parent.parent / "config" / "reconcile" / "finance" / "customer_ccs"
        ).resolve()
        run_validation_sql_dir(
            sql_dir=str(reconcile_sql_dir),
            cfg_runtime_name=args.runtime_config,
            database_path=args.database_path,
        )
        _print_reconcile_summary(
            _resolve_database_path(
                database_path=args.database_path,
                runtime_config_name=args.runtime_config,
            )
        )

        if args.export_review_workbook:
            database_path = _resolve_database_path(
                database_path=args.database_path,
                runtime_config_name=args.runtime_config,
            )
            export_paths = _export_review_workbook(
                database_path=database_path,
                export_config_path=_resolve_export_config_path(args.export_review_config),
                export_dir=Path(args.export_review_dir).expanduser().resolve()
                if args.export_review_dir
                else None,
            )
            for path in export_paths:
                print(f"Exported review workbook: {path}")
