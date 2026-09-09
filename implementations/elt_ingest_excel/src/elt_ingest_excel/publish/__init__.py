"""Publish module for writing data to Excel workbooks."""

from ..models import PublishResult
from .base import BaseExcelPublisher
from .excel_publisher_openpyxl import ExcelPublisherOpenpyxl
from .excel_publisher_xlwings import ExcelPublisherXlwings

# Default publisher (xlwings - requires Excel, preserves all features)
ExcelPublisher = ExcelPublisherXlwings

__all__ = [
    "BaseExcelPublisher",
    "PublishResult",
    "ExcelPublisher",
    "ExcelPublisherOpenpyxl",
    "ExcelPublisherXlwings",
]
