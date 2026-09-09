"""DuckDB database writer for ingested data."""

from pathlib import Path
from typing import TYPE_CHECKING, Union

import duckdb
import pandas as pd

from ..models import SaveMode, WriteResult

if TYPE_CHECKING:
    from ..reporting import PipelineReporter


class DuckDBWriter:
    """Writer for saving DataFrames to DuckDB tables.

    Supports multiple save modes:
    - DROP: Drop the table only
    - RECREATE: Drop and recreate with new data
    - OVERWRITE: Delete existing rows and insert new data
    - APPEND: Add new rows to existing table

    Example usage:
        with DuckDBWriter("/path/to/database.duckdb") as writer:
            result = writer.write(df, "my_table", SaveMode.RECREATE)
            print(f"Wrote {result.rows_written} rows")
    """

    def __init__(
        self,
        database_path: Union[str, Path],
        reporter: "PipelineReporter | None" = None,
    ):
        """Initialize the DuckDB writer.

        Args:
            database_path: Path to the DuckDB database file.
            reporter: Optional reporter for output. If None, no output is produced.
        """
        self.database_path = Path(database_path).expanduser()
        self._connection: duckdb.DuckDBPyConnection | None = None
        self.reporter = reporter

    def __enter__(self):
        """Context manager entry - connect to DuckDB."""
        self._connection = duckdb.connect(str(self.database_path))
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - close DuckDB connection."""
        if self._connection:
            self._connection.close()
            self._connection = None

    @property
    def connection(self) -> duckdb.DuckDBPyConnection:
        """Get the database connection.

        Raises:
            RuntimeError: If not used within a context manager.
        """
        if self._connection is None:
            raise RuntimeError("DuckDBWriter must be used as a context manager")
        return self._connection

    def write(
        self,
        df: pd.DataFrame,
        table_name: str,
        save_mode: SaveMode = SaveMode.RECREATE,
    ) -> WriteResult:
        """Write a DataFrame to a DuckDB table.

        Args:
            df: The pandas DataFrame to write.
            table_name: Name of the target table.
            save_mode: How to handle existing table/data.

        Returns:
            WriteResult with details of the write operation.
        """
        rows_written = 0

        if save_mode == SaveMode.DROP:
            self._drop_table(table_name)
            # For DROP mode, we don't write any data
            return WriteResult(
                table_name=table_name,
                rows_written=0,
                row_count=0,
                save_mode=save_mode,
            )

        elif save_mode == SaveMode.RECREATE:
            rows_written = self._recreate_table(df, table_name)

        elif save_mode == SaveMode.OVERWRITE:
            rows_written = self._overwrite_table(df, table_name)

        elif save_mode == SaveMode.APPEND:
            rows_written = self._append_to_table(df, table_name)

        row_count = self._get_table_count(table_name)

        return WriteResult(
            table_name=table_name,
            rows_written=rows_written,
            row_count=row_count,
            save_mode=save_mode,
        )

    def _drop_table(self, table_name: str) -> None:
        """Drop a table if it exists.

        Args:
            table_name: Name of the table to drop.
        """
        self.connection.execute(f'DROP TABLE IF EXISTS "{table_name}"')
        if self.reporter:
            self.reporter.print_table_dropped(table_name)

    def _prepare_df(self, df: pd.DataFrame) -> pd.DataFrame:
        """Prepare DataFrame for DuckDB by converting str dtype to object.

        Python 3.14 uses 'str' dtype which DuckDB doesn't recognize.
        Convert to 'object' dtype for compatibility.

        Args:
            df: DataFrame to prepare.

        Returns:
            DataFrame with compatible dtypes.
        """
        # Convert any 'str' dtype columns to 'object' for DuckDB compatibility
        df_copy = df.copy()
        for col in df_copy.columns:
            if df_copy[col].dtype.name == 'str':
                df_copy[col] = df_copy[col].astype(object)
        return df_copy

    def _recreate_table(self, df: pd.DataFrame, table_name: str) -> int:
        """Drop and recreate table with data.

        Uses DuckDB's CREATE OR REPLACE TABLE for atomic operation.

        Args:
            df: DataFrame to write.
            table_name: Name of the target table.

        Returns:
            Number of rows written.
        """
        if df.empty:
            self._drop_table(table_name)
            return 0

        # Drop existing table and create new one from DataFrame
        self._drop_table(table_name)
        df_prepared = self._prepare_df(df)
        rel = self.connection.from_df(df_prepared)
        rel.create(table_name)

        return len(df)

    def _overwrite_table(self, df: pd.DataFrame, table_name: str) -> int:
        """Delete existing data and insert new data.

        If table doesn't exist, creates it.

        Args:
            df: DataFrame to write.
            table_name: Name of the target table.

        Returns:
            Number of rows written.
        """
        if df.empty:
            # Just truncate if table exists
            if self._table_exists(table_name):
                self.connection.execute(f'DELETE FROM "{table_name}"')
            return 0

        if self._table_exists(table_name):
            # Table exists - delete all rows then insert
            self.connection.execute(f'DELETE FROM "{table_name}"')
            df_prepared = self._prepare_df(df)
            rel = self.connection.from_df(df_prepared)
            rel.insert_into(table_name)
        else:
            # Table doesn't exist - create it
            return self._recreate_table(df, table_name)

        return len(df)

    def _append_to_table(self, df: pd.DataFrame, table_name: str) -> int:
        """Append data to existing table.

        If table doesn't exist, creates it.

        Args:
            df: DataFrame to write.
            table_name: Name of the target table.

        Returns:
            Number of rows written.
        """
        if df.empty:
            return 0

        if self._table_exists(table_name):
            # Table exists - just insert
            df_prepared = self._prepare_df(df)
            rel = self.connection.from_df(df_prepared)
            rel.insert_into(table_name)
        else:
            # Table doesn't exist - create it
            return self._recreate_table(df, table_name)

        return len(df)

    def _table_exists(self, table_name: str) -> bool:
        """Check if a table exists in the database.

        Args:
            table_name: Name of the table to check.

        Returns:
            True if table exists, False otherwise.
        """
        result = self.connection.execute(
            "SELECT COUNT(*) FROM information_schema.tables "
            "WHERE table_name = ?",
            [table_name]
        ).fetchone()
        return result[0] > 0 if result else False

    def _get_table_count(self, table_name: str) -> int:
        """Get the row count from a table.

        Args:
            table_name: Name of the table.

        Returns:
            Number of rows in the table.
        """
        result = self.connection.execute(
            f'SELECT COUNT(*) FROM "{table_name}"'
        ).fetchone()
        return result[0] if result else 0
