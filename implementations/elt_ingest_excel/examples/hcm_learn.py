# examples/hcm_learn.py
"""Pipeline runner for HCM Learn data.

NOTE: the ingest workbook filename, Workday Learning target mappings, and the
publish template are PLACEHOLDERS — update config/ingest/hcm/learn.json,
config/transform/sql/hcm/learn/, and config/publish/hcm/publish_learn.json
once the source/target specifics are confirmed.
"""

from pathlib import Path

from base_runner import run_pipeline, create_parser
from elt_ingest_excel import PipelinePhase
from elt_ingest_excel.transform.sql_validation_dir import run_validation_sql_dir

if __name__ == "__main__":
    parser = create_parser("Run HCM Learn pipeline")
    args = parser.parse_args()

    run_to_phase = PipelinePhase[args.run_to_phase]

    run_pipeline(
        cfg_ingest_path="ingest/hcm",
        cfg_ingest_name="learn.json",
        cfg_transform_path="transform/sql/hcm/learn",
        cfg_publish_path="publish/hcm",
        cfg_publish_name="publish_learn.json",
        cfg_runtime_name=args.runtime_config,
        database_path=args.database_path,
        run_to_phase=run_to_phase,
    )

    if run_to_phase != PipelinePhase.INGEST:
        validation_sql_dir = (
            Path(__file__).parent.parent / "config" / "validate" / "hcm" / "learn"
        ).resolve()
        run_validation_sql_dir(
            sql_dir=str(validation_sql_dir),
            cfg_runtime_name=args.runtime_config,
            database_path=args.database_path,
        )
