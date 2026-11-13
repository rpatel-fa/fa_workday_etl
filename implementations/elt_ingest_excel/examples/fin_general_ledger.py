"""Pipeline runner for Finance General Ledger data."""

from base_runner import create_parser, run_pipeline
from elt_ingest_excel import PipelinePhase


if __name__ == "__main__":
    parser = create_parser(
        "Run Finance General Ledger pipeline",
        default_run_to_phase="INGEST",
    )
    args = parser.parse_args()

    run_pipeline(
        cfg_ingest_path="ingest/finance",
        cfg_ingest_name="general_ledger.json",
        cfg_transform_path="transform/sql/finance/general_ledger",
        cfg_publish_path="publish/finance",
        cfg_publish_name="publish_general_ledger.json",
        cfg_runtime_name=args.runtime_config,
        database_path=args.database_path,
        run_to_phase=PipelinePhase[args.run_to_phase],
    )
