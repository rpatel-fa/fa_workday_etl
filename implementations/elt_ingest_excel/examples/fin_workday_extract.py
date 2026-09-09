# examples/fin_workday_extract.py
"""Pipeline runner for Finance Workday Extract data."""

from pathlib import Path

from base_runner import run_pipeline, create_parser
from elt_ingest_excel import PipelinePhase
from elt_ingest_excel.transform.sql_validation_dir import run_validation_sql_dir

if __name__ == "__main__":
    parser = create_parser("Run Finance Workday Extract pipeline")
    args = parser.parse_args()

    run_to_phase = PipelinePhase[args.run_to_phase]

    run_pipeline(
        cfg_ingest_path="ingest/finance",
        cfg_ingest_name="workday_ccs.json",
        cfg_transform_path="transform/sql/finance/workday_extract",
        cfg_publish_path="",
        cfg_publish_name="",
        cfg_runtime_name=args.runtime_config,
        database_path=args.database_path,
        run_to_phase=run_to_phase,
    )

    if run_to_phase != PipelinePhase.INGEST:
        validation_sql_dir = (
            Path(__file__).parent.parent / "config" / "validate" / "finance" / "customer"
        ).resolve()
        run_validation_sql_dir(
            sql_dir=str(validation_sql_dir),
            cfg_runtime_name=args.runtime_config,
            database_path=args.database_path,
        )
