# examples/fin_customer_splink_compare.py
"""
Fuzzy name-match comparison for Finance Customer data.

Self-joins src_fin_customer_raw on Jaro-Winkler name similarity and writes all
pairs above the threshold to customer_fuzzy_name_comparison in the pipeline
database for review in DBeaver.

Only pairs with DIFFERENT nrm_customer_name are returned — these are cases the
deterministic normalisation did not group.  The name_similarity column (0–1) is
a direct Jaro-Winkler score, not a model probability.

Output table columns:
  unique_id_l / unique_id_r    — business_unit|customer_number for each side
  nrm_customer_name_l / _r     — normalised names being compared
  name_similarity              — Jaro-Winkler score (0–1); 1.0 = identical

Usage:
    uv run python examples/fin_customer_splink_compare.py
    uv run python examples/fin_customer_splink_compare.py --threshold 0.92
    uv run python examples/fin_customer_splink_compare.py --database-path /path/to/db.duckdb

Default threshold 0.95 eliminates personal-name false positives while preserving genuine
near-duplicates. Values below 0.92 introduce significant noise from different people
sharing the same first name (e.g. "STEPHEN HORN" vs "STEPHEN HARTSHORNE").
"""

from pathlib import Path

import duckdb

from base_runner import _load_runtime_database_path, create_parser

OUTPUT_TABLE = "customer_fuzzy_name_comparison"


def parse_args():
    parser = create_parser("Fuzzy name-match comparison for customers")
    parser.add_argument(
        "--threshold",
        type=float,
        default=0.95,
        help="Jaro-Winkler similarity threshold (default: 0.95)",
    )
    return parser.parse_args()


def resolve_database_path(args) -> str:
    if args.database_path:
        return args.database_path
    config_base = Path(__file__).parent.parent / "config"
    runtime_file = config_base / "runtime" / args.runtime_config
    path = _load_runtime_database_path(runtime_file)
    if not path:
        raise ValueError(
            f"No database path found in {runtime_file}. "
            f"Pass --database-path or configure databasePath in {runtime_file}."
        )
    return path


def run_fuzzy_match(con: duckdb.DuckDBPyConnection, threshold: float) -> int:
    con.execute(f"DROP TABLE IF EXISTS {OUTPUT_TABLE}")
    con.execute(f"""
        CREATE TABLE {OUTPUT_TABLE} AS
        SELECT
            a.nrm_business_unit || '|' || a.nrm_customer_number   AS unique_id_l
          , a.nrm_customer_name                                     AS nrm_customer_name_l
          , b.nrm_business_unit || '|' || b.nrm_customer_number   AS unique_id_r
          , b.nrm_customer_name                                     AS nrm_customer_name_r
          , ROUND(jaro_winkler_similarity(a.nrm_customer_name, b.nrm_customer_name), 4) AS name_similarity
        FROM src_fin_customer_raw  a
        INNER JOIN src_fin_customer_raw  b
            ON a.nrm_customer_name < b.nrm_customer_name
           AND jaro_winkler_similarity(a.nrm_customer_name, b.nrm_customer_name) >= {threshold}
        WHERE a.data_rnk = 1
          AND b.data_rnk = 1
        ORDER BY name_similarity DESC
    """)
    return con.execute(f"SELECT COUNT(*) FROM {OUTPUT_TABLE}").fetchone()[0]


def main() -> None:
    args = parse_args()
    con = duckdb.connect(resolve_database_path(args))
    count = run_fuzzy_match(con, args.threshold)
    con.close()

    print(f"Written {count:,} rows to {OUTPUT_TABLE}  (threshold: {args.threshold})")
    print(f"\nSuggested queries:")
    print(f"  SELECT * FROM {OUTPUT_TABLE} ORDER BY name_similarity DESC;")
    print(f"  SELECT * FROM {OUTPUT_TABLE} WHERE name_similarity >= 0.99 ORDER BY name_similarity DESC;")


if __name__ == "__main__":
    main()
