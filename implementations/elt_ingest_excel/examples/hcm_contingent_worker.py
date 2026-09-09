# examples/hcm_contingent_worker.py
"""Pipeline runner for HCM Contingent Worker data."""

import os
from pathlib import Path

from base_runner import run_pipeline, create_parser
from elt_ingest_excel import PipelinePhase
from elt_ingest_excel.transform.sql_validation_dir import run_validation_sql_dir

if __name__ == "__main__":
    if os.environ.get("ALLOW_HCM_CONTINGENT_WORKER_RUN") != "1":
        raise SystemExit(
            "Refusing to run hcm_contingent_worker pipeline. "
            "Set ALLOW_HCM_CONTINGENT_WORKER_RUN=1 to explicitly allow execution."
        )

    parser = create_parser("Run HCM Contingent Worker pipeline")
    args = parser.parse_args()

    run_to_phase = PipelinePhase[args.run_to_phase]

    run_pipeline(
        cfg_ingest_path="ingest/hcm",
        cfg_ingest_name="contingent_worker.json",
        cfg_transform_path="transform/sql/hcm/contingent_worker",
        cfg_publish_path="publish/hcm",
        cfg_publish_name="publish_contingent_worker.json",
        cfg_runtime_name=args.runtime_config,
        database_path=args.database_path,
        run_to_phase=run_to_phase,
    )

    if run_to_phase != PipelinePhase.INGEST:
        validation_sql_dir = (
            Path(__file__).parent.parent / "config" / "validate" / "hcm" / "contingent_worker"
        ).resolve()
        run_validation_sql_dir(
            sql_dir=str(validation_sql_dir),
            cfg_runtime_name=args.runtime_config,
            database_path=args.database_path,
        )
