"""Configuration models for publish/output definitions."""

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class PublishSheetConfig:
    """Configuration for a single sheet to publish.

    Attributes:
        src_table_name: DuckDB table to query for data.
        sheet_name: Name of the worksheet in the target workbook.
        header_row: Row number containing headers (1-indexed, informational).
        data_row: Row number where data starts (1-indexed).
    """
    src_table_name: str
    sheet_name: str
    header_row: int = 1
    data_row: int = 2
    use_column_name_index: bool = False


@dataclass
class PublishWorkbookConfig:
    """Configuration for publishing data to a workbook.

    Attributes:
        src_workbook_path: Path to the source/template workbook directory.
        src_workbook_file_name: Name of the source/template workbook file.
        tgt_workbook_path: Path to the target workbook directory.
        tgt_workbook_file_name: Name of the target workbook file (without extension).
        sheets: List of sheet configurations to publish.
    """
    src_workbook_path: str
    src_workbook_file_name: str
    tgt_workbook_path: str
    tgt_workbook_file_name: str
    sheets: list[PublishSheetConfig] = field(default_factory=list)

    @property
    def src_workbook_full_path(self) -> Path:
        """Full path to source workbook."""
        return Path(self.src_workbook_path).expanduser() / self.src_workbook_file_name

    @property
    def tgt_workbook_full_path(self) -> Path:
        """Full path to target workbook (with extension from source)."""
        extension = Path(self.src_workbook_file_name).suffix
        return Path(self.tgt_workbook_path).expanduser() / f"{self.tgt_workbook_file_name}{extension}"


@dataclass
class PublishConfig:
    """Top-level configuration for publishing.

    Attributes:
        workbooks: List of workbook configurations to publish.
    """
    workbooks: list[PublishWorkbookConfig] = field(default_factory=list)
