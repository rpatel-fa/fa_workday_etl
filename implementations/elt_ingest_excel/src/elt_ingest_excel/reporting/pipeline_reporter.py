"""Pipeline output reporting for ELT process."""

from pathlib import Path

from ..publish import PublishResult
from ..transform import TransformResult
from ..writers import SaveMode, WriteResult


class PipelineReporter:
    """Handles all output reporting for the ELT pipeline.

    Provides formatted console output for each pipeline phase
    including headers, progress details, and summaries.
    """

    SEPARATOR_WIDTH = 60
    SUMMARY_WIDTH = 40

    def print_extract_load_header(
        self,
        config_path: Path,
        data_file: Path,
        file_type: str,
        database_path: Path,
        save_mode: SaveMode,
        sheet_count: int,
    ) -> None:
        """Print header for Extract & Load phase.

        Args:
            config_path: Path to the ingest configuration file.
            data_file: Path to the source data file.
            file_type: Type of file being processed (e.g., "EXCEL").
            database_path: Path to the DuckDB database.
            save_mode: How existing tables are handled.
            sheet_count: Number of sheets to process.
        """
        print("\n" + "=" * self.SEPARATOR_WIDTH)
        print("EXTRACT & LOAD PHASE")
        print("=" * self.SEPARATOR_WIDTH)
        print(f"Config: {config_path}")
        print(f"Data file: {data_file}")
        print(f"File type: {file_type}")
        print(f"Database: {database_path}")
        print(f"Save mode: {save_mode.value}")
        print(f"Processing {sheet_count} sheet(s)")

    def print_transform_header(self, transform_path: Path) -> None:
        """Print header for Transform phase.

        Args:
            transform_path: Path to the transform configuration directory.
        """
        print("\n" + "=" * self.SEPARATOR_WIDTH)
        print("TRANSFORM PHASE")
        print("=" * self.SEPARATOR_WIDTH)
        print(f"Transform path: {transform_path}")

    def print_transform_no_order_file(self, order_file: Path) -> None:
        """Print message when order.txt is not found.

        Args:
            order_file: Path where order.txt was expected.
        """
        print(f"No order.txt found at {order_file}")

    def print_transform_sql_count(self, count: int) -> None:
        """Print number of SQL files to execute.

        Args:
            count: Number of SQL files.
        """
        print(f"SQL files to execute: {count}")

    def print_sql_file_start(self, sql_file: str) -> None:
        """Print message when starting to execute a SQL file.

        Args:
            sql_file: Name of the SQL file being executed.
        """
        print(f"\n  Executing: {sql_file}")

    def print_sql_file_not_found(self, sql_file: str, path: str) -> None:
        """Print error when SQL file is not found.

        Args:
            sql_file: Name of the SQL file.
            path: Full path where file was expected.
        """
        print(f"    ERROR: SQL file not found: {path}")

    def print_sql_statement_executed(self, statement_num: int) -> None:
        """Print message when a SQL statement is executed.

        Args:
            statement_num: The statement number (1-indexed).
        """
        print(f"    Statement {statement_num} executed")

    def print_sql_file_success(self) -> None:
        """Print success message for SQL file execution."""
        print("    SUCCESS")

    def print_sql_file_error(self, error: str) -> None:
        """Print error message for SQL file execution.

        Args:
            error: The error message.
        """
        print(f"    ERROR: {error}")

    def print_publish_header(self) -> None:
        """Print header for Publish phase."""
        print("\n" + "=" * self.SEPARATOR_WIDTH)
        print("PUBLISH PHASE")
        print("=" * self.SEPARATOR_WIDTH)

    def print_publish_no_config(self) -> None:
        """Print message when no publish config is configured."""
        print("No publish config path configured")

    def print_publish_config_not_found(self, config_path: Path) -> None:
        """Print message when publish config file is not found.

        Args:
            config_path: Path where config was expected.
        """
        print(f"Publish config not found: {config_path}")

    def print_publish_config_info(
        self,
        config_path: Path,
        publisher_type: str,
        workbook_count: int,
    ) -> None:
        """Print publish configuration details.

        Args:
            config_path: Path to the publish configuration file.
            publisher_type: Type of publisher being used.
            workbook_count: Number of workbooks to publish.
        """
        print(f"Publish config: {config_path}")
        print(f"Publisher: {publisher_type}")
        print(f"Workbooks to publish: {workbook_count}")

    def print_sheet_start(
        self,
        sheet_name: str,
        target_table: str,
        header_row: int,
        data_row: int,
    ) -> None:
        """Print message when starting to process a sheet.

        Args:
            sheet_name: Name of the sheet being processed.
            target_table: Target table name.
            header_row: Header row number.
            data_row: Data row number.
        """
        print(f"\n  Sheet: {sheet_name}")
        print(f"    Target table: {target_table}")
        print(f"    Header row: {header_row}, Data row: {data_row}")

    def print_sheet_rows_read(self, count: int) -> None:
        """Print number of rows read from sheet.

        Args:
            count: Number of rows read.
        """
        print(f"    Rows read: {count}")

    def print_sheet_rows_written(self, count: int) -> None:
        """Print number of rows written to database.

        Args:
            count: Number of rows written.
        """
        print(f"    Rows written: {count}")

    def print_load_summary(self, results: list[WriteResult]) -> None:
        """Print summary of load results.

        Args:
            results: List of WriteResult objects from the load phase.
        """
        print("\n" + "-" * self.SUMMARY_WIDTH)
        print("Load Summary:")
        total_rows = 0
        for result in results:
            print(f"  {result.table_name}: {result.row_count} rows")
            total_rows += result.row_count
        print(f"Total: {len(results)} tables, {total_rows} rows")

    def print_transform_summary(self, results: list[TransformResult]) -> None:
        """Print summary of transform results.

        Args:
            results: List of TransformResult objects from the transform phase.
        """
        print("\n" + "-" * self.SUMMARY_WIDTH)
        print("Transform Summary:")
        success_count = sum(1 for r in results if r.success)
        fail_count = len(results) - success_count

        for result in results:
            status = "OK" if result.success else f"FAILED: {result.error}"
            print(f"  {result.sql_file}: {status}")

        print(f"Total: {success_count} succeeded, {fail_count} failed")

    def print_transform_abort_on_failure(self) -> None:
        """Print notice when aborting remaining SQL files after a failure."""
        print("  Aborting remaining SQL files due to failure")

    def print_publish_summary(self, results: list[PublishResult]) -> None:
        """Print summary of publish results.

        Args:
            results: List of PublishResult objects from the publish phase.
        """
        print("\n" + "-" * self.SUMMARY_WIDTH)
        print("Publish Summary:")
        success_count = sum(1 for r in results if r.success)
        fail_count = len(results) - success_count
        total_rows = sum(r.rows_written for r in results if r.success)

        for result in results:
            if result.success:
                print(f"  {result.sheet_name}: {result.rows_written} rows")
            else:
                print(f"  {result.sheet_name}: FAILED - {result.error}")

        print(f"Total: {success_count} succeeded, {fail_count} failed, {total_rows} rows written")

    # Workbook publishing methods

    def print_workbook_source(self, src_path: Path) -> None:
        """Print source template path.

        Args:
            src_path: Path to the source template workbook.
        """
        print(f"\n  Source template: {src_path}")

    def print_workbook_target(self, tgt_path: Path) -> None:
        """Print target workbook path.

        Args:
            tgt_path: Path to the target workbook.
        """
        print(f"  Target workbook: {tgt_path}")

    def print_workbook_deleting_existing(self, tgt_path: Path) -> None:
        """Print message when deleting existing target.

        Args:
            tgt_path: Path to the target being deleted.
        """
        print(f"  Deleting existing target: {tgt_path}")

    def print_workbook_cloning(self) -> None:
        """Print message when cloning template."""
        print("  Cloning template to target...")

    def print_workbook_saving(self) -> None:
        """Print message when saving workbook."""
        print("  Saving workbook...")

    def print_workbook_saved(self, tgt_path: Path) -> None:
        """Print message when workbook is saved.

        Args:
            tgt_path: Path to the saved workbook.
        """
        print(f"  Workbook saved: {tgt_path}")

    def print_publish_sheet_start(self, sheet_name: str, table_name: str) -> None:
        """Print message when starting to publish a sheet.

        Args:
            sheet_name: Name of the sheet being published.
            table_name: Source table name.
        """
        print(f"\n    Sheet: {sheet_name}")
        print(f"    Source table: {table_name}")

    def print_publish_rows_from_table(self, count: int) -> None:
        """Print number of rows fetched from table.

        Args:
            count: Number of rows fetched.
        """
        print(f"    Rows from table: {count}")

    def print_publish_rows_written(self, count: int) -> None:
        """Print number of rows written to sheet.

        Args:
            count: Number of rows written.
        """
        print(f"    Rows written: {count}")

    def print_publish_sheet_error(self, error: str) -> None:
        """Print error message for sheet publish.

        Args:
            error: The error message.
        """
        print(f"    ERROR: {error}")

    # Database writer methods

    def print_table_dropped(self, table_name: str) -> None:
        """Print message when a table is dropped.

        Args:
            table_name: Name of the table that was dropped.
        """
        print(f"    Dropped table: {table_name}")

    def print_reconciliation_header(self, table_name: str) -> None:
        """Print header for reconciliation section.

        Args:
            table_name: Name of the source table being reconciled.
        """
        print("\n" + "-" * self.SUMMARY_WIDTH)
        print(f"Ingestion Reconciliation ({table_name}):")
        print("-" * self.SUMMARY_WIDTH)

    def print_reconciliation_row(
        self,
        business_unit: str,
        ingested_rows: int,
        merged_rows: int,
        deduped_rows: int,
        status: str,
    ) -> None:
        """Print a row of reconciliation results.

        Args:
            business_unit: The business unit identifier.
            ingested_rows: Total rows ingested from raw tables.
            merged_rows: Rows in the merged table.
            deduped_rows: Rows removed by deduplication.
            status: Validation status message.
        """
        print(f"  {business_unit}:")
        print(f"    Ingested:    {ingested_rows:>8} rows")
        print(f"    Merged:      {merged_rows:>8} rows")
        print(f"    Deduplicated:{deduped_rows:>8} rows")
        print(f"    Status:      {status}")

    def print_reconciliation_totals(
        self,
        total_ingested: int,
        total_merged: int,
        total_deduped: int,
    ) -> None:
        """Print totals for reconciliation.

        Args:
            total_ingested: Total rows ingested across all business units.
            total_merged: Total rows in merged table.
            total_deduped: Total rows removed by deduplication.
        """
        print("-" * self.SUMMARY_WIDTH)
        print(f"  TOTAL:")
        print(f"    Ingested:    {total_ingested:>8} rows")
        print(f"    Merged:      {total_merged:>8} rows")
        print(f"    Deduplicated:{total_deduped:>8} rows")
