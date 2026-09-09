"""Pipeline runner for Finance Customer Invoice Adjustment data."""

from pathlib import Path

from base_runner import create_parser, run_pipeline
from elt_ingest_excel import PipelinePhase
from elt_ingest_excel.transform.sql_validation_dir import run_validation_sql_dir


if __name__ == "__main__":
    parser = create_parser(
        "Run Finance Customer Invoice Adjustment pipeline",
        default_run_to_phase="INGEST",
    )
    args = parser.parse_args()

    run_to_phase = PipelinePhase[args.run_to_phase]

    run_pipeline(
        cfg_ingest_path="ingest/finance",
        cfg_ingest_name="customer_invoice_adjustment.json",
        cfg_transform_path="transform/sql/finance/customer_invoice_adjustment",
        cfg_publish_path="publish/finance",
        cfg_publish_name="publish_customer_invoice_adjustment.json",
        cfg_runtime_name=args.runtime_config,
        database_path=args.database_path,
        run_to_phase=run_to_phase,
    )

    if run_to_phase != PipelinePhase.INGEST:
        validation_sql_dir = (
            Path(__file__).parent.parent
            / "config" / "validate" / "finance" / "customer_invoice_adjustment"
        ).resolve()
        run_validation_sql_dir(
            sql_dir=str(validation_sql_dir),
            cfg_runtime_name=args.runtime_config,
            database_path=args.database_path,
        )
