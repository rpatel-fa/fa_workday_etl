from __future__ import annotations

import json
from pathlib import Path

import duckdb

from .db import open_connection


def _load_runtime_database_path(runtime_config_file: Path) -> str | None:
    if not runtime_config_file.exists():
        return None

    data = json.loads(runtime_config_file.read_text())
    if isinstance(data, list):
        data = data[0] if data else {}
    if not isinstance(data, dict):
        raise ValueError(
            f"Invalid runtime config format (expected object): {runtime_config_file}"
        )

    for key in ("databasePath", "database_path", "duckdbPath", "duckdb_path"):
        value = data.get(key)
        if value:
            return str(value)
    return None


def _resolve_sql_dir(sql_dir: str | Path, *, config_base_path: Path) -> Path:
    raw = Path(sql_dir).expanduser()
    candidates: list[Path] = []

    if raw.is_absolute():
        candidates.append(raw)
    else:
        candidates.append((config_base_path / raw).resolve())
        candidates.append((Path.cwd() / raw).resolve())
        this_file = Path(__file__).resolve()
        candidates.append((this_file.parents[3] / raw).resolve())
        candidates.append((this_file.parents[4] / raw).resolve())

    for p in candidates:
        if p.exists() and p.is_dir():
            return p

    tried = "\n".join(str(p) for p in candidates)
    raise FileNotFoundError(f"SQL directory not found. Tried:\n{tried}")


def _execute_sql_file(conn: duckdb.DuckDBPyConnection, sql_path: Path) -> None:
    sql_content = sql_path.read_text()
    statements = conn.extract_statements(sql_content)
    for i, statement in enumerate(statements, 1):
        try:
            conn.execute(statement)
        except duckdb.Error as e:
            raise RuntimeError(f"{sql_path.name}: statement {i} failed: {e}") from e


def run_validation_sql_dir(
    *,
    sql_dir: str | Path,
    cfg_runtime_name: str = "default.json",
    database_path: str | Path | None = None,
    continue_on_error: bool = False,
    config_base_path: str | Path | None = None,
) -> None:
    resolved_config_base_path = (
        Path(config_base_path).expanduser().resolve()
        if config_base_path is not None
        else (Path(__file__).resolve().parents[3] / "config").resolve()
    )

    runtime_config_file = (resolved_config_base_path / "runtime" / cfg_runtime_name).resolve()
    resolved_database_path = str(database_path) if database_path is not None else None
    resolved_database_path = resolved_database_path or _load_runtime_database_path(
        runtime_config_file
    )
    if not resolved_database_path:
        raise ValueError(
            "No DuckDB database path provided. "
            f"Pass --database-path or configure it in {runtime_config_file}"
        )

    db_path = Path(resolved_database_path).expanduser().resolve()
    resolved_sql_dir = _resolve_sql_dir(sql_dir, config_base_path=resolved_config_base_path)
    init_sql = (resolved_sql_dir / "init.sql").resolve()
    order_txt = (resolved_sql_dir / "order.txt").resolve()
    available_sql = {
        p.name: p
        for p in resolved_sql_dir.glob("*.sql")
        if p.is_file()
    }
    if order_txt.exists():
        ordered_names = [
            line.strip()
            for line in order_txt.read_text().splitlines()
            if line.strip() and not line.strip().startswith("#")
        ]
        sql_files = [
            available_sql[name]
            for name in ordered_names
            if name in available_sql and available_sql[name].resolve() != init_sql.resolve()
        ]
        seen = set()
        for name in available_sql:
            if name == init_sql.name:
                continue
            if name not in ordered_names:
                sql_files.append(available_sql[name])
    else:
        sql_files = sorted(
            p
            for p in available_sql.values()
            if p.resolve() != init_sql.resolve()
        )
    if not sql_files:
        raise FileNotFoundError(f"No .sql files found in: {resolved_sql_dir}")

    with open_connection(db_path) as conn:
        if init_sql.exists():
            _execute_sql_file(conn, init_sql)
            print(f"OK  {init_sql.name}")
        for sql_path in sql_files:
            try:
                _execute_sql_file(conn, sql_path)
                print(f"OK  {sql_path.name}")
            except Exception as e:
                print(f"ERR {sql_path.name}  {e}")
                if not continue_on_error:
                    raise
