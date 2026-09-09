"""Pipeline runner for Finance Workday CCS Customer extract data."""

from base_runner import create_parser, run_pipeline
from elt_ingest_excel import PipelinePhase


if __name__ == "__main__":
    parser = create_parser(
        "Run Finance Workday CCS Customer extract pipeline",
        default_run_to_phase="TRANSFORM",
    )
    args = parser.parse_args()

    run_to_phase = PipelinePhase[args.run_to_phase]

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
