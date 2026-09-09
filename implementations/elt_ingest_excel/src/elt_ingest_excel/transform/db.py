from pathlib import Path
from typing import TYPE_CHECKING

import duckdb

from .udf import register_all

if TYPE_CHECKING:
    from duckdb import DuckDBPyConnection


def open_connection(database_path: Path) -> "DuckDBPyConnection":
    conn = duckdb.connect(str(database_path))
    register_all(conn)
    return conn
