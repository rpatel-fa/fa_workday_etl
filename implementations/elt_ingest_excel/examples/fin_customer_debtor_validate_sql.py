"""SQL validation runner for Finance Customer/Debtor data."""

from base_runner import create_parser

from elt_ingest_excel.transform.sql_validation_dir import run_validation_sql_dir

VALIDATION_SQL_DIR = "elt_ingest_excel/config/validate/finance/customer"

if __name__ == "__main__":
    parser = create_parser("Run Finance Customer/Debtor SQL validations")
    args = parser.parse_args()

    run_validation_sql_dir(
        sql_dir=VALIDATION_SQL_DIR,
        cfg_runtime_name=args.runtime_config,
        database_path=args.database_path,
    )
