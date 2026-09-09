"""Sheet processor for extracting and loading data from various file types."""

from pathlib import Path
from typing import TYPE_CHECKING

import pandas as pd

from ..models import FileType, SheetConfig, WorkbookConfig
from ..writers import DuckDBWriter, SaveMode, WriteResult
from .excel_reader import ExcelReader

if TYPE_CHECKING:
    from ..reporting import PipelineReporter


class SheetProcessor:
    """Processes sheets from workbooks and loads them into DuckDB.

    Handles different file types (Excel, delimited) and delegates
    to appropriate readers for data extraction.
    """

    def __init__(
        self,
        data_path: Path,
        data_file_name: str,
        workbook_config: WorkbookConfig,
        save_mode: SaveMode = SaveMode.RECREATE,
        reporter: "PipelineReporter | None" = None,
    ):
        """Initialize the sheet processor.

        Args:
            data_path: Path to the data files directory.
            data_file_name: Name of the data file to process.
            workbook_config: Configuration for the workbook.
            save_mode: How to handle existing tables.
            reporter: Optional reporter for output. If None, no output is produced.
        """
        self.data_path = Path(data_path)
        self.data_file_name = data_file_name
        self.workbook_config = workbook_config
        self.save_mode = save_mode
        self.reporter = reporter
        self.file_path = self.data_path / self.data_file_name

    def process_sheets(
        self,
        sheets: list[SheetConfig],
        writer: DuckDBWriter,
    ) -> list[WriteResult]:
        """Process multiple sheets and write to DuckDB.

        Args:
            sheets: List of sheet configurations to process.
            writer: DuckDBWriter instance for database operations.

        Returns:
            List of WriteResult objects for each sheet processed.
        """
        results = []
        for sheet_config in sheets:
            result = self.process_sheet(sheet_config, writer)
            if result:
                results.append(result)
        return results

    def process_sheet(
        self,
        sheet_config: SheetConfig,
        writer: DuckDBWriter,
    ) -> WriteResult | None:
        """Process a single sheet.

        Args:
            sheet_config: Configuration for the sheet to process.
            writer: DuckDBWriter instance for database operations.

        Returns:
            WriteResult if data was written, None otherwise.
        """
        if self.reporter:
            self.reporter.print_sheet_start(
                sheet_name=sheet_config.sheet_name,
                target_table=sheet_config.target_table_name,
                header_row=sheet_config.header_row,
                data_row=sheet_config.data_row,
            )

        if self.workbook_config.file_type == FileType.EXCEL:
            return self._process_excel_sheet(sheet_config, writer)
        elif self.workbook_config.file_type == FileType.DELIMITED:
            return self._process_delimited_file(sheet_config, writer)
        else:
            raise ValueError(f"Unsupported file type: {self.workbook_config.file_type}")

    def _process_excel_sheet(
        self,
        sheet_config: SheetConfig,
        writer: DuckDBWriter,
    ) -> WriteResult:
        """Process an Excel sheet.

        Args:
            sheet_config: Configuration for the sheet to process.
            writer: DuckDBWriter instance for database operations.

        Returns:
            WriteResult with details of the write operation.
        """
        reader = ExcelReader(
            file_path=self.file_path,
            sheet_name=sheet_config.sheet_name,
            header_row=sheet_config.header_row,
            data_row=sheet_config.data_row,
            column_headers=sheet_config.column_headers,
            dtype=str,
        )
        df = reader.load()

        if self.reporter:
            self.reporter.print_sheet_rows_read(len(df))

        # Write to DuckDB
        result = writer.write(df, sheet_config.target_table_name, self.save_mode)

        if self.reporter:
            self.reporter.print_sheet_rows_written(result.rows_written)

        return result

    def _process_delimited_file(
        self,
        sheet_config: SheetConfig,
        writer: DuckDBWriter,
    ) -> WriteResult:
        """Process a delimited file (CSV, TSV, etc.).

        Args:
            sheet_config: Configuration for the file to process.
            writer: DuckDBWriter instance for database operations.

        Returns:
            WriteResult with details of the write operation.

        """
        suffix = self.file_path.suffix.lower()
        if self.workbook_config.delimiter is not None:
            sep = self.workbook_config.delimiter
        elif suffix == ".tsv":
            sep = "\t"
        elif suffix == ".psv":
            sep = "|"
        else:
            sep = ","

        raw = pd.read_csv(
            self.file_path,
            header=None,
            dtype=str,
            sep=sep,
            na_filter=False,
            keep_default_na=False,
        )

        reader = ExcelReader(
            file_path=self.file_path,
            sheet_name=sheet_config.sheet_name,
            header_row=sheet_config.header_row,
            data_row=sheet_config.data_row,
            column_headers=sheet_config.column_headers,
            dtype=str,
        )
        df = reader._finalize_raw(raw)

        if self.reporter:
            self.reporter.print_sheet_rows_read(len(df))

        result = writer.write(df, sheet_config.target_table_name, self.save_mode)

        if self.reporter:
            self.reporter.print_sheet_rows_written(result.rows_written)

        return result
