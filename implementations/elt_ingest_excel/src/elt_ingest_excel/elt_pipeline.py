"""Main workflow module for file ingestion and transformation."""

from enum import Enum
from pathlib import Path
from typing import Union

import duckdb
import pandas as pd

from .loaders import SheetProcessor
from .parsers import JsonConfigParser, PublishConfigParser
from .publish import ExcelPublisherOpenpyxl, ExcelPublisherXlwings, PublishResult
from .reporting import PipelineReporter
from .transform import SqlExecutor, TransformResult
from .transform.validation import validate_counties_against_master
from .writers import SaveMode, DuckDBWriter, WriteResult


class PipelinePhase(Enum):
    """Pipeline phases for controlling execution scope.

    Phases are cumulative - each phase includes all previous phases:
    - INGEST: Extract and load data only
    - TRANSFORM: Ingest + run SQL transformations
    - PUBLISH: Ingest + transform + publish to Excel
    """

    INGEST = "ingest"
    TRANSFORM = "transform"
    PUBLISH = "publish"


class FileIngestor:
    """Main ELT workflow orchestrator.

    This class orchestrates the full ELT process by delegating to
    specialized modules:
    1. Extract/Load - SheetProcessor reads source files and writes to DuckDB
    2. Transform - SqlExecutor runs SQL transformations
    3. Publish - ExcelPublisher exports data to Excel workbooks
    4. Reporting - PipelineReporter handles all console output

    Supports multiple file types (EXCEL, DELIMITED), though
    currently only EXCEL is implemented.
    """

    def __init__(
        self,
        config_base_path: Union[str, Path],
        cfg_ingest_path: str,
        cfg_ingest_name: str,
        cfg_transform_path: str,
        cfg_publish_path: str | None = None,
        cfg_publish_name: str | None = None,
        data_path: Union[str, Path, None] = None,
        data_file_name: str = "*",
        database_path: Union[str, Path] = "",
        sheet_filter: str = "*",
        save_mode: SaveMode = SaveMode.RECREATE,
        publisher_type: str = "xlwings",
        master_workbook_path: Union[str, Path, None] = None,
    ):
        """Initialize the file ingestor.

        Args:
            config_base_path: Base path to the config directory.
            cfg_ingest_path: Relative path to ingest config (e.g., "ingest/finance").
            cfg_ingest_name: Name of the ingest JSON config file (e.g., "supplier.json").
            cfg_transform_path: Relative path to transform config (e.g., "transform/finance").
            cfg_publish_path: Relative path to publish config (e.g., "publish/finance").
            cfg_publish_name: Name of the publish JSON config file.
            data_path: Path to the data files directory.
            data_file_name: Name of the data file to process, or "*" for all workbooks in config.
            database_path: Path to DuckDB database file.
            sheet_filter: Sheet name to filter on, or "*" for all sheets.
            save_mode: How to handle existing tables (DROP, RECREATE, OVERWRITE, APPEND).
            publisher_type: Excel publisher to use - "xlwings" (default) or "openpyxl".
                           "xlwings" preserves drawing shapes (requires Excel installed).
                           "openpyxl" works without Excel but may lose shapes in .xlsm files.
            master_workbook_path: Optional path to a master reference workbook used for
                                  county/region validation during the transform phase.
                                  Validation is skipped when None or path does not exist.
        """
        self.config_base_path = Path(config_base_path).expanduser()
        self.cfg_ingest_path = cfg_ingest_path
        self.cfg_ingest_name = cfg_ingest_name
        self.cfg_transform_path = cfg_transform_path
        self.cfg_publish_path = cfg_publish_path
        self.cfg_publish_name = cfg_publish_name
        self.publisher_type = publisher_type
        self.data_path = Path(data_path).expanduser() if data_path else None
        self.data_file_name = data_file_name
        self.database_path = Path(database_path).expanduser()
        self.sheet_filter = sheet_filter
        self.save_mode = save_mode
        self.master_workbook_path = Path(master_workbook_path).expanduser() if master_workbook_path else None

        # Build full config paths
        self.ingest_config_path = self.config_base_path / cfg_ingest_path / cfg_ingest_name
        self.transform_config_path = self.config_base_path / cfg_transform_path
        self.publish_config_path = (
            self.config_base_path / cfg_publish_path / cfg_publish_name
            if cfg_publish_path
            else None
        )

        # Configure pandas display
        pd.set_option("display.max_columns", None)
        pd.set_option("display.width", None)
        pd.set_option("display.max_colwidth", 50)

        # Load ingest configuration
        self.config = JsonConfigParser.from_json(
            self.ingest_config_path,
            database_path=str(self.database_path),
        )

        # Get workbooks to process (filtered by data_file_name, or all if "*")
        self.workbooks = JsonConfigParser.get_workbooks(
            self.config,
            self.data_file_name,
        )

        # Initialize reporter
        self.reporter = PipelineReporter()

        # Store results
        self.load_results: list[WriteResult] = []
        self.transform_results: list[TransformResult] = []
        self.publish_results: list[PublishResult] = []

    def extract_and_load(self) -> list[WriteResult]:
        """Execute Extract and Load phases.

        Reads data from source files and writes to DuckDB tables.
        Processes all workbooks matching the data_file_name filter.

        Returns:
            List of WriteResult objects for each sheet processed.
        """
        self.load_results = []

        with DuckDBWriter(self.database_path, reporter=self.reporter) as writer:
            for workbook in self.workbooks:
                workbook_data_path = (
                    Path(workbook.workbook_path_name).expanduser()
                    if getattr(workbook, "workbook_path_name", None)
                    else self.data_path
                )
                if workbook_data_path is None:
                    raise ValueError(
                        f"No data path available for workbook '{workbook.workbook_file_name}'. "
                        "Provide data_path to FileIngestor or set workbookPathName in ingest config."
                    )

                # Get sheets for this workbook (filtered)
                sheets = JsonConfigParser.get_sheets(workbook, self.sheet_filter)

                self.reporter.print_extract_load_header(
                    config_path=self.ingest_config_path,
                    data_file=workbook_data_path / workbook.workbook_file_name,
                    file_type=workbook.file_type.value,
                    database_path=self.database_path,
                    save_mode=self.save_mode,
                    sheet_count=len(sheets),
                )

                # Create sheet processor for this workbook
                processor = SheetProcessor(
                    data_path=workbook_data_path,
                    data_file_name=workbook.workbook_file_name,
                    workbook_config=workbook,
                    save_mode=self.save_mode,
                    reporter=self.reporter,
                )

                workbook_results = processor.process_sheets(sheets, writer)
                self.load_results.extend(workbook_results)

        self.reporter.print_load_summary(self.load_results)
        return self.load_results

    def transform(self) -> list[TransformResult]:
        """Execute Transform phase.

        Reads order.txt from transform config path and executes
        SQL files in the specified order.

        Returns:
            List of TransformResult objects for each SQL file executed.
        """
        self.reporter.print_transform_header(self.transform_config_path)

        # Create SQL executor and delegate transformation
        executor = SqlExecutor(
            transform_path=self.transform_config_path,
            database_path=self.database_path,
            reporter=self.reporter,
        )

        sql_count = executor.get_sql_file_count()
        if sql_count == 0:
            self.reporter.print_transform_no_order_file(
                self.transform_config_path / "order.txt"
            )
            return []

        self.reporter.print_transform_sql_count(sql_count)

        self.transform_results = executor.execute()

        self.reporter.print_transform_summary(self.transform_results)

        master_path = self.master_workbook_path
        if master_path is not None:
            ok, missing, _ = validate_counties_against_master(
                database_path=self.database_path,
                master_workbook_path=master_path,
                sheet_name="Country States-Regions",
            )
        if master_path is not None and master_path.exists():
            print("\n" + "=" * self.reporter.SEPARATOR_WIDTH)
            print("COUNTY VALIDATION")
            print("=" * self.reporter.SEPARATOR_WIDTH)
            if ok:
                print("All counties match the master list")
            else:
                print("Counties not in master list:")
                for c in sorted(missing):
                    print(f"- {c}")
                raise RuntimeError("County validation failed")

        # Print reconciliation results if tables exist
        self._print_reconciliation_results()

        return self.transform_results

    def _print_reconciliation_results(self) -> None:
        """Query and print reconciliation results for all matching tables.

        Discovers any tables named validation_%_reconciliation in the database
        and prints their contents to verify data loaded correctly.
        """
        with duckdb.connect(str(self.database_path)) as conn:
            table_rows = conn.execute(
                "SELECT table_name FROM information_schema.tables "
                "WHERE table_name LIKE 'validation_%_reconciliation' "
                "ORDER BY table_name",
            ).fetchall()
            reconciliation_tables = [row[0] for row in table_rows]

            for table_name in reconciliation_tables:
                self.reporter.print_reconciliation_header(table_name)

                data_rows = conn.execute(
                    f"SELECT business_unit, ingested_rows, merged_rows, "
                    f"deduped_rows, status FROM {table_name} ORDER BY business_unit"
                ).fetchall()

                total_ingested = 0
                total_merged = 0
                total_deduped = 0

                for row in data_rows:
                    bu, ingested, merged, deduped, status = row
                    self.reporter.print_reconciliation_row(
                        business_unit=bu,
                        ingested_rows=ingested,
                        merged_rows=merged,
                        deduped_rows=deduped,
                        status=status,
                    )
                    total_ingested += ingested
                    total_merged += merged
                    total_deduped += deduped

                self.reporter.print_reconciliation_totals(
                    total_ingested=total_ingested,
                    total_merged=total_merged,
                    total_deduped=total_deduped,
                )

    def publish(self) -> list[PublishResult]:
        """Execute Publish phase.

        Reads publish config and exports transformed data to Excel workbooks.

        Returns:
            List of PublishResult objects for each sheet published.
        """
        self.reporter.print_publish_header()

        if not self.publish_config_path:
            self.reporter.print_publish_no_config()
            return []

        if not self.publish_config_path.exists():
            self.reporter.print_publish_config_not_found(self.publish_config_path)
            return []

        # Load publish configuration
        publish_config = PublishConfigParser.from_json(self.publish_config_path)

        self.reporter.print_publish_config_info(
            config_path=self.publish_config_path,
            publisher_type=self.publisher_type,
            workbook_count=len(publish_config.workbooks),
        )

        self.publish_results = []

        # Select publisher based on type and delegate publishing
        if self.publisher_type == "xlwings":
            publisher_class = ExcelPublisherXlwings
        else:
            publisher_class = ExcelPublisherOpenpyxl

        with publisher_class(self.database_path, reporter=self.reporter) as publisher:
            self.publish_results = publisher.publish(publish_config)

        self.reporter.print_publish_summary(self.publish_results)
        return self.publish_results

    def process(
        self,
        run_to_phase: PipelinePhase | str = PipelinePhase.PUBLISH,
    ) -> tuple[list[WriteResult], list[TransformResult], list[PublishResult]]:
        """Execute ELT pipeline up to the specified phase.

        Args:
            run_to_phase: Phase to run up to (inclusive). Can be:
                - PipelinePhase.INGEST or "ingest": Extract and load only
                - PipelinePhase.TRANSFORM or "transform": Ingest + transform
                - PipelinePhase.PUBLISH or "publish": Full pipeline (default)

        Returns:
            Tuple of (load_results, transform_results, publish_results).
            Later phases return empty lists if not executed.
        """
        # Normalize to enum
        if isinstance(run_to_phase, str):
            run_to_phase = PipelinePhase(run_to_phase.lower())

        # Always run ingest
        load_results = self.extract_and_load()

        if run_to_phase == PipelinePhase.INGEST:
            return load_results, [], []

        # Run transform
        transform_results = self.transform()

        # Abort on any transform failure
        if any(not r.success for r in transform_results):
            raise RuntimeError("Transform failed; aborting pipeline before publish")

        if run_to_phase == PipelinePhase.TRANSFORM:
            return load_results, transform_results, []

        # Run publish
        publish_results = self.publish()
        return load_results, transform_results, publish_results
