# examples/fin_delta_customer_debtor.py
"""Pipeline runner for Finance Customer/Debtor delta data."""

import os
from pathlib import Path

from base_runner import run_pipeline, create_parser
from elt_ingest_excel import PipelinePhase
from elt_ingest_excel.transform.sql_validation_dir import run_validation_sql_dir

if __name__ == "__main__":
    if os.environ.get("ALLOW_FIN_DELTA_CUSTOMER_DEBTOR_RUN") != "1":
        raise SystemExit(
            "Refusing to run fin_delta_customer_debtor pipeline. "
            "Set ALLOW_FIN_DELTA_CUSTOMER_DEBTOR_RUN=1 to explicitly allow execution."
        )

    parser = create_parser("Run Finance Customer/Debtor delta pipeline")
    args = parser.parse_args()

    run_to_phase = PipelinePhase[args.run_to_phase]

    run_pipeline(
        cfg_ingest_path="ingest/finance",
        cfg_ingest_name="delta_customer_debtor.json",
        cfg_transform_path="transform/sql/finance/delta_customer",
        cfg_publish_path="publish/finance",
        cfg_publish_name="publish_customer.json",
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
