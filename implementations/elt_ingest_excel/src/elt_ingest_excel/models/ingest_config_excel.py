"""Main configuration model for Excel ingestion."""

from dataclasses import dataclass, field
from pathlib import Path

from .ingest_config import WorkbookConfig


@dataclass
class ExcelIngestConfig:
    """Configuration for Excel workbook ingestion into DuckDB.

    Attributes:
        database_path: Path to the DuckDB database file.
        workbooks: List of workbook configurations to process.
        create_tables: Whether to create tables if they don't exist (default True).
        replace_data: Whether to replace existing data or append (default True).
    """
    database_path: Path
    workbooks: list[WorkbookConfig] = field(default_factory=list)
    create_tables: bool = True
    replace_data: bool = True
