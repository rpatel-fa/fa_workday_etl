"""ELT module for ingesting Excel workbooks into DuckDB.

This module provides functionality to:
- Read Excel workbooks using openpyxl
- Load worksheet data into DuckDB tables
- Configure ingestion via JSON configuration files
- Transform data using SQL
- Publish results to Excel workbooks

Example usage:

    from elt_ingest_excel import FileIngestor, PipelinePhase

    # Create pipeline
    pipeline = FileIngestor(
        config_base_path="~/config",
        cfg_ingest_path="ingest/finance",
        cfg_ingest_name="supplier.json",
        cfg_transform_path="transform/finance",
        data_path="~/data",
        data_file_name="suppliers.xlsx",
        database_path="~/output/data.duckdb",
    )

    # Run full ELT pipeline
    load_results, transform_results, publish_results = pipeline.process()
"""

from .models import (
    ExcelIngestConfig,
    FileType,
    WorkbookConfig,
    SheetConfig,
    PublishConfig,
    PublishWorkbookConfig,
    PublishSheetConfig,
)
from .parsers import JsonConfigParser, PublishConfigParser
from .loaders import ExcelReader, SheetProcessor
from .writers import SaveMode, DuckDBWriter, WriteResult
from .publish import ExcelPublisher, PublishResult
from .transform import SqlExecutor, TransformResult
from .elt_pipeline import FileIngestor, PipelinePhase

__all__ = [
    # Models - Ingest
    "ExcelIngestConfig",
    "FileType",
    "WorkbookConfig",
    "SheetConfig",
    # Models - Publish
    "PublishConfig",
    "PublishWorkbookConfig",
    "PublishSheetConfig",
    # Parsers
    "JsonConfigParser",
    "PublishConfigParser",
    # Loaders
    "ExcelReader",
    "SheetProcessor",
    # Writers
    "SaveMode",
    "DuckDBWriter",
    "WriteResult",
    # Publisher
    "ExcelPublisher",
    "PublishResult",
    # Transform
    "SqlExecutor",
    "TransformResult",
    # ELT Pipeline (main workflow)
    "FileIngestor",
    "PipelinePhase",
]

__version__ = "0.1.0"
