"""Result dataclasses for pipeline operations."""

from dataclasses import dataclass

from .save_mode import SaveMode


@dataclass
class WriteResult:
    """Result of writing data to a table.

    Attributes:
        table_name: Name of the target table.
        rows_written: Number of rows written.
        row_count: Verified count from the table after writing.
        save_mode: The save mode used.
    """
    table_name: str
    rows_written: int
    row_count: int
    save_mode: SaveMode


@dataclass
class TransformResult:
    """Result of executing a transform SQL file.

    Attributes:
        sql_file: Name of the SQL file executed.
        success: Whether execution succeeded.
        error: Error message if failed, None otherwise.
    """
    sql_file: str
    success: bool
    error: str | None = None


@dataclass
class PublishResult:
    """Result of publishing data to a sheet.

    Attributes:
        sheet_name: Name of the sheet written to.
        table_name: Source DuckDB table.
        rows_written: Number of rows written.
        success: Whether the publish succeeded.
        error: Error message if failed, None otherwise.
    """
    sheet_name: str
    table_name: str
    rows_written: int
    success: bool
    error: str | None = None
