"""Export multiple SQL query results into a single Excel workbook.

This utility reads a SQL file containing multiple statements, executes each
statement against DuckDB, and writes one worksheet per result set.

Usage
-----
    uv run --package elt-ingest-excel python \
      elt_ingest_excel/scripts/export_sql_workbook.py \
      --sql /path/to/queries.sql

Optional per-statement sheet names can be declared with a leading comment:

    -- sheet: Workday in Delta
    SELECT * FROM ...;

    -- sheet: Delta in Workday
    SELECT * FROM ...;
"""

from __future__ import annotations

import argparse
import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import duckdb
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DEFAULT_RUNTIME_CONFIG = Path(
    "/Users/Rakesh.Patel/Documents/__code/git/emailrak/elt_lake/"
    "elt_ingest_excel/config/runtime/default.json"
)
SHEET_DIRECTIVE_RE = re.compile(r"^\s*--\s*(?:sheet|name)\s*:\s*(.+?)\s*$", re.IGNORECASE)
INVALID_SHEET_CHARS_RE = re.compile(r"[\[\]\:\*\?\/\\]")
QUERY_KEYWORDS = {"select", "with", "show", "describe", "pragma", "explain"}


@dataclass(frozen=True)
class QueryBlock:
    sql: str
    sheet_name: str | None


def parse_args(argv: Iterable[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--sql",
        required=True,
        help="Path to a SQL file containing one or more statements.",
    )
    parser.add_argument(
        "--out",
        help="Output workbook path. Defaults to <sql file>.xlsx.",
    )
    parser.add_argument(
        "--duckdb",
        help="DuckDB path. Defaults to the path in runtime/default.json.",
    )
    parser.add_argument(
        "--runtime-config",
        default=str(DEFAULT_RUNTIME_CONFIG),
        help="Runtime config used to resolve the default DuckDB path.",
    )
    parser.add_argument(
        "--read-write",
        action="store_true",
        help="Open DuckDB in read-write mode (required for CREATE TABLE statements).",
    )
    return parser.parse_args(argv)


def resolve_duckdb_path(explicit_path: str | None, runtime_config_path: str) -> str:
    if explicit_path:
        return os.path.expanduser(explicit_path)

    with open(os.path.expanduser(runtime_config_path), encoding="utf-8") as fh:
        config = json.load(fh)

    database_path = (
        config.get("databasePath")
        or config.get("duckdb", {}).get("path")
        or config.get("duckdb_path")
    )
    if not database_path:
        raise ValueError(
            f"Could not find a DuckDB path in runtime config: {runtime_config_path}"
        )
    return os.path.expanduser(database_path)


def split_sql_statements(sql_text: str) -> list[str]:
    statements: list[str] = []
    buffer: list[str] = []
    in_single_quote = False
    in_double_quote = False
    in_line_comment = False
    in_block_comment = False
    i = 0

    while i < len(sql_text):
        ch = sql_text[i]
        nxt = sql_text[i + 1] if i + 1 < len(sql_text) else ""

        if in_line_comment:
            buffer.append(ch)
            if ch == "\n":
                in_line_comment = False
            i += 1
            continue

        if in_block_comment:
            buffer.append(ch)
            if ch == "*" and nxt == "/":
                buffer.append(nxt)
                in_block_comment = False
                i += 2
            else:
                i += 1
            continue

        if ch == "-" and nxt == "-" and not in_single_quote and not in_double_quote:
            buffer.append(ch)
            buffer.append(nxt)
            in_line_comment = True
            i += 2
            continue

        if ch == "/" and nxt == "*" and not in_single_quote and not in_double_quote:
            buffer.append(ch)
            buffer.append(nxt)
            in_block_comment = True
            i += 2
            continue

        if ch == "'" and not in_double_quote:
            buffer.append(ch)
            if nxt == "'" and in_single_quote:
                buffer.append(nxt)
                i += 2
                continue
            in_single_quote = not in_single_quote
            i += 1
            continue

        if ch == '"' and not in_single_quote:
            buffer.append(ch)
            in_double_quote = not in_double_quote
            i += 1
            continue

        if ch == ";" and not in_single_quote and not in_double_quote:
            statement = "".join(buffer).strip()
            if statement:
                statements.append(statement)
            buffer = []
            i += 1
            continue

        buffer.append(ch)
        i += 1

    tail = "".join(buffer).strip()
    if tail:
        statements.append(tail)

    return statements


def extract_sheet_name(statement: str) -> QueryBlock:
    lines = statement.splitlines()
    sheet_name: str | None = None
    cleaned_lines: list[str] = []
    seen_sql_content = False

    for line in lines:
        match = SHEET_DIRECTIVE_RE.match(line)
        stripped = line.strip()

        if match and not seen_sql_content and sheet_name is None:
            sheet_name = match.group(1).strip()
            continue

        if stripped and not stripped.startswith("--"):
            seen_sql_content = True

        cleaned_lines.append(line)

    sql = "\n".join(cleaned_lines).strip()
    return QueryBlock(sql=sql, sheet_name=sheet_name)


def sanitize_sheet_name(name: str, used_names: set[str], index: int) -> str:
    base = INVALID_SHEET_CHARS_RE.sub("_", name).strip().strip("'")
    if not base:
        base = f"query_{index:02d}"

    base = base[:31]
    candidate = base
    suffix = 2
    while candidate in used_names:
        suffix_text = f"_{suffix}"
        candidate = f"{base[: 31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    used_names.add(candidate)
    return candidate


def iter_query_blocks(sql_path: Path) -> list[QueryBlock]:
    statements = split_sql_statements(sql_path.read_text(encoding="utf-8"))
    blocks = [extract_sheet_name(statement) for statement in statements]
    blocks = [block for block in blocks if block.sql]
    if not blocks:
        raise ValueError(f"No executable SQL statements found in {sql_path}")
    return blocks


def is_query_statement(sql: str) -> bool:
    for line in sql.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("--"):
            continue
        keyword = stripped.split(None, 1)[0].lower()
        return keyword in QUERY_KEYWORDS
    return False


def render_workbook(
    con: duckdb.DuckDBPyConnection, blocks: list[QueryBlock], out_path: Path
) -> int:
    workbook = Workbook()
    workbook.remove(workbook.active)

    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    header_font = Font(bold=True)
    wrap_alignment = Alignment(wrap_text=True, vertical="top")
    plain_alignment = Alignment(vertical="top")
    used_names: set[str] = set()
    sheet_count = 0

    for index, block in enumerate(blocks, start=1):
        cursor = con.execute(block.sql)
        if not is_query_statement(block.sql) and block.sheet_name is None:
            continue
        headers = [item[0] for item in (cursor.description or [])]
        if not headers:
            continue

        rows = cursor.fetchall()

        raw_name = block.sheet_name or f"query_{index:02d}"
        sheet_name = sanitize_sheet_name(raw_name, used_names, index)
        ws = workbook.create_sheet(title=sheet_name)
        ws.freeze_panes = "A2"
        sheet_count += 1

        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = wrap_alignment

        for row_idx, row in enumerate(rows, start=2):
            max_lines = 1
            for col_idx, value in enumerate(row, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                text_value = "" if value is None else str(value)
                if "\n" in text_value or "\r" in text_value:
                    cell.alignment = wrap_alignment
                    normalized = text_value.replace("\r\n", "\n").replace("\r", "\n")
                    max_lines = max(max_lines, normalized.count("\n") + 1)
                else:
                    cell.alignment = plain_alignment
            ws.row_dimensions[row_idx].height = min(max(15, 15 * max_lines), 120)

        for col_idx, header in enumerate(headers, start=1):
            values = [str(header)]
            for row in rows:
                value = row[col_idx - 1]
                if value is not None:
                    values.append(str(value).replace("\r\n", "\n").replace("\r", "\n"))
            max_len = max(
                (
                    max((len(line) for line in text.split("\n")), default=0)
                    for text in values
                ),
                default=0,
            )
            ws.column_dimensions[get_column_letter(col_idx)].width = min(
                max(max_len + 2, 12),
                60,
            )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    if sheet_count == 0:
        raise ValueError("No result sets found to export (no worksheets created).")
    workbook.save(out_path)
    return sheet_count


def main(argv: Iterable[str] | None = None) -> int:
    args = parse_args(argv)
    sql_path = Path(os.path.expanduser(args.sql)).resolve()
    out_path = (
        Path(os.path.expanduser(args.out)).resolve()
        if args.out
        else sql_path.with_suffix(".xlsx")
    )
    duckdb_path = resolve_duckdb_path(args.duckdb, args.runtime_config)
    blocks = iter_query_blocks(sql_path)

    con = duckdb.connect(duckdb_path, read_only=not args.read_write)
    try:
        sheet_count = render_workbook(con, blocks, out_path)
    finally:
        con.close()

    print(out_path)
    print(f"Exported {sheet_count} worksheet(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
