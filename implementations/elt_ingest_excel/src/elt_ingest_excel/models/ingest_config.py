"""Configuration models for workbook and sheet definitions."""

from dataclasses import dataclass, field
from enum import Enum


class FileType(Enum):
    """Supported file types for ingestion."""
    EXCEL = "EXCEL"
    DELIMITED = "DELIMITED"


@dataclass
class SheetConfig:
    """Configuration for a single worksheet to be loaded.

    Attributes:
        sheet_name: Name of the worksheet in the Excel workbook.
        target_table_name: Name of the DuckDB table to load data into.
        header_row: Row number containing headers (1-indexed). Use 0 for "no headers".
        data_row: Row number where data starts (1-indexed, default is header_row + 1 when header_row > 0, else 1).
        column_headers: Optional explicit column names to use when header_row is 0 (no header row in the sheet).
    """
    sheet_name: str
    target_table_name: str
    header_row: int = 1
    data_row: int | None = None
    column_headers: list[str] | None = None

    def __post_init__(self):
        """Set data_row default if not provided."""
        if self.data_row is None:
            self.data_row = (self.header_row + 1) if self.header_row > 0 else 1


@dataclass
class WorkbookConfig:
    """Configuration for a single workbook/file.

    Attributes:
        workbook_file_name: Path to the workbook/file.
        workbook_path_name: Optional base directory for workbook_file_name.
        sheets: List of sheet configurations to process.
        file_type: Type of file (EXCEL, DELIMITED, etc.). Defaults to EXCEL.
    """
    workbook_file_name: str
    workbook_path_name: str | None = None
    sheets: list[SheetConfig] = field(default_factory=list)
    file_type: FileType = FileType.EXCEL
    delimiter: str | None = None
