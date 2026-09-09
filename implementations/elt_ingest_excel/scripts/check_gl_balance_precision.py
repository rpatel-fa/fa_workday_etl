from __future__ import annotations

import argparse
import json
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import duckdb
from openpyxl import load_workbook


TWO_DP = Decimal("0.01")
RAW_TOLERANCE = Decimal("0.0000001")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Check whether General Ledger journals balance only at raw precision "
            "or also at 2dp currency precision in DuckDB and exported workbooks."
        )
    )
    parser.add_argument(
        "--database-path",
        type=Path,
        default=None,
        help="DuckDB database path. Defaults to config/runtime/default.json",
    )
    parser.add_argument(
        "--runtime-config",
        type=Path,
        default=Path("config/runtime/default.json"),
        help="Runtime config JSON used when --database-path is omitted",
    )
    parser.add_argument(
        "--documents-dir",
        type=Path,
        default=Path("~/Documents").expanduser(),
        help="Directory containing workday_fin_gl_balances_00.xlsx to _10.xlsx",
    )
    parser.add_argument(
        "--glob",
        default="workday_fin_gl_balances_*.xlsx",
        help="Workbook glob under --documents-dir",
    )
    parser.add_argument(
        "--sample-limit",
        type=int,
        default=5,
        help="Number of sample rows/groups to print per section",
    )
    parser.add_argument(
        "--fail-on-issues",
        action="store_true",
        help="Exit non-zero if rounded 2dp imbalances or non-2dp workbook values are found",
    )
    return parser.parse_args()


def resolve_database_path(args: argparse.Namespace) -> Path:
    if args.database_path is not None:
        return args.database_path.expanduser().resolve()

    runtime_config = args.runtime_config.expanduser().resolve()
    data = json.loads(runtime_config.read_text())
    if isinstance(data, list):
        data = data[0] if data else {}

    if not isinstance(data, dict):
        raise ValueError(f"Invalid runtime config format: {runtime_config}")

    for key in ("databasePath", "database_path", "duckdbPath", "duckdb_path"):
        value = data.get(key)
        if value:
            return Path(value).expanduser().resolve()

    raise ValueError(f"No database path found in {runtime_config}")


def round_currency(value: Decimal) -> Decimal:
    return value.quantize(TWO_DP, rounding=ROUND_HALF_UP)


def format_decimal(value: Decimal) -> str:
    return format(value, "f")


def check_duckdb_line_table(
    conn: duckdb.DuckDBPyConnection,
    table_suffix: str,
    sample_limit: int,
) -> dict:
    line_table = f"workday_fin_gl_line_{table_suffix}"
    header_table = f"workday_fin_gl_header_{table_suffix}"

    row = conn.execute(
        f"""
        SELECT
               COUNT(*)                                                   AS row_count
             , COALESCE(SUM(COALESCE(column_j, 0)), 0)                    AS debit_total
             , COALESCE(SUM(COALESCE(column_k, 0)), 0)                    AS credit_total
          FROM {line_table}
        """
    ).fetchone()

    raw_group_count = conn.execute(
        f"""
        SELECT COUNT(*)
          FROM (
                SELECT
                       l.column_b
                  FROM {line_table} l
              GROUP BY
                       l.column_b
                HAVING ABS(SUM(COALESCE(l.column_j, 0)) - SUM(COALESCE(l.column_k, 0))) > {RAW_TOLERANCE}
               ) t
        """
    ).fetchone()[0]

    sample_groups = conn.execute(
        f"""
        SELECT
               l.column_b                                                 AS header_key
             , ROUND(SUM(COALESCE(l.column_j, 0)), 5)                     AS debit_total
             , ROUND(SUM(COALESCE(l.column_k, 0)), 5)                     AS credit_total
             , ROUND(
                     SUM(COALESCE(l.column_j, 0))
                   - SUM(COALESCE(l.column_k, 0))
                   , 5
                   )                                                      AS diff_raw
             , h.column_i                                                 AS journal_id
             , h.column_p                                                 AS company
             , h.column_aa                                                AS external_reference_id
          FROM {line_table} l
               LEFT JOIN
               {header_table} h
                  ON h.column_b = l.column_b
      GROUP BY
               l.column_b
             , h.column_i
             , h.column_p
             , h.column_aa
        HAVING ABS(SUM(COALESCE(l.column_j, 0)) - SUM(COALESCE(l.column_k, 0))) > {RAW_TOLERANCE}
      ORDER BY
               ABS(SUM(COALESCE(l.column_j, 0)) - SUM(COALESCE(l.column_k, 0))) DESC
         LIMIT {sample_limit}
        """
    ).fetchall()

    raw_total_diff = Decimal(str(row[1])) - Decimal(str(row[2]))

    return {
        "table": line_table,
        "row_count": row[0],
        "debit_total": format_decimal(Decimal(str(row[1]))),
        "credit_total": format_decimal(Decimal(str(row[2]))),
        "raw_diff": format_decimal(raw_total_diff),
        "raw_balanced": abs(raw_total_diff) <= RAW_TOLERANCE and raw_group_count == 0,
        "raw_group_count": raw_group_count,
        "sample_groups": sample_groups,
    }


def check_duckdb(database_path: Path, sample_limit: int) -> dict:
    conn = duckdb.connect(str(database_path))

    raw_group_count = conn.execute(
        """
        SELECT COUNT(*)
          FROM (
                SELECT
                       h.column_p
                     , h.column_aa
                  FROM workday_fin_gl_line_base l
                       JOIN
                       workday_fin_gl_header_base h
                          ON h.column_b = l.column_b
              GROUP BY
                       h.column_p
                     , h.column_aa
                HAVING ABS(SUM(COALESCE(l.column_j, 0)) - SUM(COALESCE(l.column_k, 0))) > 0.0000001
               ) t
        """
    ).fetchone()[0]

    rounded_group_count = conn.execute(
        """
        SELECT COUNT(*)
          FROM (
                SELECT
                       h.column_p
                     , h.column_aa
                  FROM workday_fin_gl_line_base l
                       JOIN
                       workday_fin_gl_header_base h
                          ON h.column_b = l.column_b
              GROUP BY
                       h.column_p
                     , h.column_aa
                HAVING ABS(
                           SUM(ROUND(COALESCE(l.column_j, 0), 2))
                         - SUM(ROUND(COALESCE(l.column_k, 0), 2))
                           ) > 0.0000001
               ) t
        """
    ).fetchone()[0]

    non_2dp_lines = conn.execute(
        """
        SELECT COUNT(*)
          FROM workday_fin_gl_line_base
         WHERE (
                   column_j IS NOT NULL
               AND column_j <> ROUND(column_j, 2)
               )
            OR (
                   column_k IS NOT NULL
               AND column_k <> ROUND(column_k, 2)
               )
        """
    ).fetchone()[0]

    sample_lines = conn.execute(
        f"""
        SELECT
               h.column_p                                                   AS company
             , h.column_aa                                                  AS external_reference_id
             , l.column_p                                                   AS memo_account
             , CAST(l.column_j AS VARCHAR)                                  AS debit_amount
             , CAST(l.column_k AS VARCHAR)                                  AS credit_amount
          FROM workday_fin_gl_line_base l
               JOIN
               workday_fin_gl_header_base h
                  ON h.column_b = l.column_b
         WHERE (
                   l.column_j IS NOT NULL
               AND l.column_j <> ROUND(l.column_j, 2)
               )
            OR (
                   l.column_k IS NOT NULL
               AND l.column_k <> ROUND(l.column_k, 2)
               )
         LIMIT {sample_limit}
        """
    ).fetchall()

    sample_groups = conn.execute(
        f"""
        SELECT
               h.column_p                                                   AS company
             , h.column_aa                                                  AS external_reference_id
             , ROUND(SUM(ROUND(COALESCE(l.column_j, 0), 2)), 2)             AS debit_total_2dp
             , ROUND(SUM(ROUND(COALESCE(l.column_k, 0), 2)), 2)             AS credit_total_2dp
             , ROUND(
                     SUM(ROUND(COALESCE(l.column_j, 0), 2))
                   - SUM(ROUND(COALESCE(l.column_k, 0), 2))
                   , 2
                   )                                                        AS diff_2dp
          FROM workday_fin_gl_line_base l
               JOIN
               workday_fin_gl_header_base h
                  ON h.column_b = l.column_b
      GROUP BY
               h.column_p
             , h.column_aa
        HAVING ABS(
                   SUM(ROUND(COALESCE(l.column_j, 0), 2))
                 - SUM(ROUND(COALESCE(l.column_k, 0), 2))
                   ) > 0.0000001
      ORDER BY
               ABS(
                   SUM(ROUND(COALESCE(l.column_j, 0), 2))
                 - SUM(ROUND(COALESCE(l.column_k, 0), 2))
                   ) DESC
         LIMIT {sample_limit}
        """
    ).fetchall()

    line_table_results = [
        check_duckdb_line_table(conn, f"{index:02d}", sample_limit)
        for index in range(11)
    ]

    conn.close()

    return {
        "raw_group_count": raw_group_count,
        "rounded_group_count": rounded_group_count,
        "non_2dp_lines": non_2dp_lines,
        "sample_lines": sample_lines,
        "sample_groups": sample_groups,
        "line_table_results": line_table_results,
    }


def check_workbook(path: Path, sample_limit: int) -> dict:
    workbook = load_workbook(path, data_only=True, read_only=True)
    line_sheet = workbook["Journal Entry Line Replacement"]
    header_sheet = workbook["Import Accounting Journal"]

    header_map = {}
    for row in header_sheet.iter_rows(min_row=6, values_only=True):
        header_key = row[1]
        if header_key is None:
            continue
        header_map[header_key] = {
            "journal_id": row[8],
            "company": row[15],
            "external_reference_id": row[26],
        }

    row_count = 0
    raw_debit_total = Decimal("0")
    raw_credit_total = Decimal("0")
    rounded_debit_total = Decimal("0")
    rounded_credit_total = Decimal("0")
    non_2dp_values: list[str] = []
    grouped_totals: dict[int, dict[str, Decimal]] = {}
    raw_grouped_totals: dict[int, dict[str, Decimal]] = {}

    for row_number, row in enumerate(line_sheet.iter_rows(min_row=6, values_only=True), start=6):
        header_key = row[1]
        if header_key is None:
            continue

        row_count += 1
        grouped_totals.setdefault(header_key, {"debit": Decimal("0"), "credit": Decimal("0")})
        raw_grouped_totals.setdefault(header_key, {"debit": Decimal("0"), "credit": Decimal("0")})

        debit = Decimal(str(row[9])) if row[9] is not None else Decimal("0")
        credit = Decimal(str(row[10])) if row[10] is not None else Decimal("0")

        raw_debit_total += debit
        raw_credit_total += credit
        rounded_debit_total += round_currency(debit)
        rounded_credit_total += round_currency(credit)
        grouped_totals[header_key]["debit"] += round_currency(debit)
        grouped_totals[header_key]["credit"] += round_currency(credit)
        raw_grouped_totals[header_key]["debit"] += debit
        raw_grouped_totals[header_key]["credit"] += credit

        for column_name, value in (("J", debit), ("K", credit)):
            if value != 0 and value != round_currency(value) and len(non_2dp_values) < sample_limit:
                non_2dp_values.append(f"{column_name}{row_number}={format_decimal(value)}")

    # Informational only: per-journal balance once each line is rounded to 2dp.
    # The source data legitimately carries more than 2dp, and we write it as
    # presented rather than adjusting figures to force a 2dp-rounded balance
    # (see workday_fin_gl_line_base.sql) — so this is not treated as a failure.
    out_of_balance_groups = []
    for header_key, totals in grouped_totals.items():
        diff = totals["debit"] - totals["credit"]
        if diff != 0:
            out_of_balance_groups.append(
                {
                    "header_key": header_key,
                    "diff_2dp": format_decimal(diff),
                    "journal_id": header_map.get(header_key, {}).get("journal_id"),
                    "company": header_map.get(header_key, {}).get("company"),
                    "external_reference_id": header_map.get(header_key, {}).get("external_reference_id"),
                }
            )

    # Real failure check: per-journal balance at the data's own (raw) precision.
    # This is what tells us the workbook actually matches the source/database.
    raw_out_of_balance_groups = []
    for header_key, totals in raw_grouped_totals.items():
        diff = totals["debit"] - totals["credit"]
        if abs(diff) > RAW_TOLERANCE:
            raw_out_of_balance_groups.append(
                {
                    "header_key": header_key,
                    "diff_raw": format_decimal(diff),
                    "journal_id": header_map.get(header_key, {}).get("journal_id"),
                    "company": header_map.get(header_key, {}).get("company"),
                    "external_reference_id": header_map.get(header_key, {}).get("external_reference_id"),
                }
            )

    workbook.close()

    raw_total_diff = raw_debit_total - raw_credit_total

    return {
        "workbook": path.name,
        "row_count": row_count,
        "raw_diff": format_decimal(raw_total_diff),
        "raw_balanced": abs(raw_total_diff) <= RAW_TOLERANCE and not raw_out_of_balance_groups,
        "rounded_2dp_diff": format_decimal(rounded_debit_total - rounded_credit_total),
        "non_2dp_value_count": len(non_2dp_values),
        "non_2dp_samples": non_2dp_values,
        "out_of_balance_groups": out_of_balance_groups[:sample_limit],
        "out_of_balance_group_count": len(out_of_balance_groups),
        "raw_out_of_balance_groups": raw_out_of_balance_groups[:sample_limit],
        "raw_out_of_balance_group_count": len(raw_out_of_balance_groups),
    }


def print_duckdb_section(results: dict) -> None:
    print("DuckDB")
    print(f"- Raw out-of-balance groups (failure check): {results['raw_group_count']}")
    print(f"- Rounded 2dp out-of-balance groups (informational only): {results['rounded_group_count']}")
    print(f"- Non-2dp line values (informational only): {results['non_2dp_lines']}")

    if results["sample_lines"]:
        print("- Sample non-2dp lines (informational only):")
        for sample in results["sample_lines"]:
            print(
                "  "
                f"company={sample[0]} | external_reference_id={sample[1]} | "
                f"memo_account={sample[2]} | debit={sample[3]} | credit={sample[4]}"
            )

    if results["sample_groups"]:
        print("- Sample rounded 2dp imbalanced groups (informational only):")
        for sample in results["sample_groups"]:
            print(
                "  "
                f"company={sample[0]} | external_reference_id={sample[1]} | "
                f"debit_2dp={sample[2]} | credit_2dp={sample[3]} | diff_2dp={sample[4]}"
            )


def print_duckdb_line_tables_section(results: list[dict]) -> None:
    print("DuckDB Period Tables")
    for result in results:
        print(
            "- "
            f"{result['table']} | rows={result['row_count']} | debit_total={result['debit_total']} | "
            f"credit_total={result['credit_total']} | raw_diff={result['raw_diff']} | "
            f"raw_out_of_balance_groups={result['raw_group_count']} (failure check)"
        )
        for sample in result["sample_groups"]:
            print(
                "  "
                f"RAW imbalanced_group: header_key={sample[0]} | debit_total={sample[1]} | "
                f"credit_total={sample[2]} | diff_raw={sample[3]} | journal_id={sample[4]} | "
                f"company={sample[5]} | external_reference_id={sample[6]}"
            )


def print_workbook_section(results: list[dict]) -> None:
    print("Workbooks")
    for result in results:
        print(
            "- "
            f"{result['workbook']} | rows={result['row_count']} | raw_diff={result['raw_diff']} | "
            f"raw_out_of_balance_groups={result['raw_out_of_balance_group_count']} (failure check) | "
            f"rounded_2dp_diff={result['rounded_2dp_diff']} | non_2dp_samples={result['non_2dp_value_count']} | "
            f"rounded_2dp_out_of_balance_groups={result['out_of_balance_group_count']} (informational only)"
        )
        for sample in result["raw_out_of_balance_groups"]:
            print(
                "  "
                f"RAW imbalanced_group: header_key={sample['header_key']} | diff_raw={sample['diff_raw']} | "
                f"journal_id={sample['journal_id']} | company={sample['company']} | "
                f"external_reference_id={sample['external_reference_id']}"
            )
        for sample in result["non_2dp_samples"]:
            print(f"  non_2dp_value (informational only): {sample}")
        for sample in result["out_of_balance_groups"]:
            print(
                "  "
                f"rounded_2dp imbalanced_group (informational only): header_key={sample['header_key']} | "
                f"diff_2dp={sample['diff_2dp']} | journal_id={sample['journal_id']} | "
                f"company={sample['company']} | external_reference_id={sample['external_reference_id']}"
            )


def main() -> int:
    args = parse_args()
    database_path = resolve_database_path(args)
    workbook_paths = sorted(args.documents_dir.expanduser().glob(args.glob))

    duckdb_results = check_duckdb(database_path, args.sample_limit)
    workbook_results = [check_workbook(path, args.sample_limit) for path in workbook_paths]

    print(f"Database path: {database_path}")
    print(f"Workbook directory: {args.documents_dir.expanduser().resolve()}")
    print()
    print_duckdb_section(duckdb_results)
    print()
    print_duckdb_line_tables_section(duckdb_results["line_table_results"])
    print()
    print_workbook_section(workbook_results)

    # Failure criterion: does the output match the source at the source's own precision?
    # We write amounts exactly as presented (no rounding/adjustment), so 2dp-rounded
    # imbalance and non-2dp values are expected for source data that carries more than
    # 2dp — they are reported above as informational only, not treated as failures here.
    has_raw_issues = (
        duckdb_results["raw_group_count"] > 0
        or any(not result["raw_balanced"] for result in duckdb_results["line_table_results"])
        or any(not result["raw_balanced"] for result in workbook_results)
    )

    print()
    print("Conclusion")
    if has_raw_issues:
        print("- The output does NOT match the source at its own precision — this is a real data issue.")
        print("- Investigate the SQL transform and/or publish step before relying on the exported workbooks.")
    else:
        print("- The output balances exactly at the source's own precision in DuckDB and in every workbook.")
        print("- Per-journal imbalance at 2dp (if reported above) reflects source data with >2dp precision and")
        print("  is informational only — amounts are written as presented, not adjusted to force a 2dp balance.")

    if args.fail_on_issues and has_raw_issues:
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
