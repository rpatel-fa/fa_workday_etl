# examples/base_runner.py
"""
Base runner for ELTP pipeline (Extract, Load, Transform, Publish).

This module contains the shared logic and documentation for running pipelines.
Individual instance runners should import and call run_pipeline() with their
specific configuration values.
"""

import argparse
import json
from pathlib import Path
from typing import Optional

from elt_ingest_excel import FileIngestor, SaveMode, PipelinePhase


def create_parser(
    description: str,
    *,
    default_run_to_phase: str = "PUBLISH",
) -> argparse.ArgumentParser:
    """Create an argument parser with common pipeline options."""
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "--run-to-phase",
        type=str,
        default=default_run_to_phase,
        choices=["INGEST", "TRANSFORM", "PUBLISH"],
        help=f"Pipeline phase to run up to (default: {default_run_to_phase})",
    )
    parser.add_argument(
        "--runtime-config",
        type=str,
        default="default.json",
        help="Runtime config JSON filename under config/runtime/ (default: default.json)",
    )
    parser.add_argument(
        "--database-path",
        type=str,
        default=None,
        help="DuckDB database path override (highest precedence)",
    )
    return parser


def _load_runtime_database_path(runtime_config_file: Path) -> str | None:
    if not runtime_config_file.exists():
        return None

    data = json.loads(runtime_config_file.read_text())
    if isinstance(data, list):
        data = data[0] if data else {}
    if not isinstance(data, dict):
        raise ValueError(f"Invalid runtime config format (expected object): {runtime_config_file}")

    for key in ("databasePath", "database_path", "duckdbPath", "duckdb_path"):
        value = data.get(key)
        if value:
            return str(value)
    return None


def run_pipeline(
    *,
    # Data source
    data_path: str | None = None,
    # Ingest config
    cfg_ingest_path: str,
    cfg_ingest_name: str,
    # Transform config
    cfg_transform_path: str,
    # Publish config
    cfg_publish_path: str,
    cfg_publish_name: str,
    # Optional overrides (sensible defaults provided)
    data_file_name: str = "*",
    database_path: str | None = None,
    sheet_filter: str = "*",
    save_mode: SaveMode = SaveMode.RECREATE,
    publisher_type: str = "xlwings",
    run_to_phase: PipelinePhase = PipelinePhase.PUBLISH,
    config_base_path: Optional[Path] = None,
    cfg_runtime_path: str = "runtime",
    cfg_runtime_name: str = "default.json",
):
    """
    Run the ELTP pipeline with the given configuration.

    Args:
        data_path: Optional directory containing the source Excel file(s). When None, uses workbookPathName from ingest config.
        cfg_ingest_path: Path to ingest config directory (relative to config_base_path)
        cfg_ingest_name: Name of the ingest config JSON file
        cfg_transform_path: Path to transform SQL directory (relative to config_base_path)
        cfg_publish_path: Path to publish config directory (relative to config_base_path)
        cfg_publish_name: Name of the publish config JSON file
        data_file_name: Name of specific Excel file, or "*" to process all workbooks in config
        database_path: Path to DuckDB database file
        sheet_filter: Sheet name filter ("*" for all sheets, or specific name)
        save_mode: How to handle existing tables
            - SaveMode.DROP: Drop table if exists
            - SaveMode.RECREATE: Drop and recreate table
            - SaveMode.OVERWRITE: Overwrite existing data
            - SaveMode.APPEND: Append to existing data
        publisher_type: Excel publisher backend
            - "openpyxl": No Excel required, but may lose drawing shapes in .xlsm files
            - "xlwings": Requires Excel installed, preserves all shapes/macros/formatting
        run_to_phase: Pipeline phase to run up to (inclusive)
            - PipelinePhase.INGEST: Extract and load only
            - PipelinePhase.TRANSFORM: Ingest + SQL transformations
            - PipelinePhase.PUBLISH: Full pipeline
        config_base_path: Base path for config files (defaults to ../config relative to caller)

    Returns:
        Tuple of (load_results, transform_results, publish_results)
    """
    if config_base_path is None:
        config_base_path = Path(__file__).parent.parent / "config"

    runtime_config_file = config_base_path / cfg_runtime_path / cfg_runtime_name
    resolved_database_path = (
        database_path
        or _load_runtime_database_path(runtime_config_file)
    )
    if not resolved_database_path:
        raise ValueError(
            "No DuckDB database path provided. "
            f"Pass --database-path or configure it in {runtime_config_file}"
        )

    ingestor = FileIngestor(
        config_base_path=config_base_path,
        cfg_ingest_path=cfg_ingest_path,
        cfg_ingest_name=cfg_ingest_name,
        cfg_transform_path=cfg_transform_path,
        cfg_publish_path=cfg_publish_path,
        cfg_publish_name=cfg_publish_name,
        data_path=data_path,
        data_file_name=data_file_name,
        database_path=resolved_database_path,
        sheet_filter=sheet_filter,
        save_mode=save_mode,
        publisher_type=publisher_type,
    )

    return ingestor.process(run_to_phase)
