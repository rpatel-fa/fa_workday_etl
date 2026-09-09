"""Excel publisher using xlwings for writing DuckDB data to Excel workbooks.

This implementation uses xlwings which controls Excel directly via COM (Windows)
or AppleScript (Mac). This preserves all Excel features including:
- Drawing shapes
- VBA macros
- Charts
- Formatting

Requires Excel to be installed on the machine.
"""

from decimal import Decimal
from pathlib import Path
from typing import TYPE_CHECKING, Any, Union

import xlwings as xw

from ..models import PublishWorkbookConfig
from .base import BaseExcelPublisher, PublishResult

if TYPE_CHECKING:
    from ..reporting import PipelineReporter


class ExcelPublisherXlwings(BaseExcelPublisher):
    """Publisher for writing DuckDB data to Excel workbooks using xlwings.

    This class handles:
    1. Cloning a template workbook to target location
    2. Querying DuckDB tables for data
    3. Writing data to specific sheets in the workbook

    Uses xlwings to control Excel directly, preserving all features.
    """

    def __init__(
        self,
        database_path: Union[str, Path],
        reporter: "PipelineReporter | None" = None,
    ):
        """Initialize the publisher.

        Args:
            database_path: Path to the DuckDB database file.
            reporter: Optional reporter for output messages.
        """
        super().__init__(database_path, reporter)

    def _open_and_process_workbook(
        self,
        workbook_config: PublishWorkbookConfig,
        tgt_path: Path,
    ) -> list[PublishResult]:
        """Open workbook with xlwings and process all sheets.

        Args:
            workbook_config: Configuration for the workbook.
            tgt_path: Path to the target workbook file.

        Returns:
            List of PublishResult for each sheet processed.
        """
        # Open Excel in background (not visible, silent mode)
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False

        # On Mac, try to set macro security to enable all macros
        # This avoids the "Enable Macros" prompt
        try:
            # msoAutomationSecurityLow = 1 (enable all macros)
            app.api.AutomationSecurity = 1
        except (AttributeError, Exception):
            # May not be available on all platforms/versions
            pass

        results = []
        try:
            # Open the target workbook
            wb = app.books.open(str(tgt_path))

            try:
                # Process each sheet
                results = self._process_sheets(wb, workbook_config)

                # Save the workbook
                self._report_saving(tgt_path)
                wb.save()
                self._report_saved(tgt_path)

            finally:
                wb.close()

        finally:
            # Restore Excel settings before quitting
            try:
                app.display_alerts = True
                app.screen_updating = True
                # msoAutomationSecurityByUI = 2 (use UI settings)
                app.api.AutomationSecurity = 2
            except (AttributeError, Exception):
                pass
            app.quit()

        return results

    def _sheet_exists(self, workbook: Any, sheet_name: str) -> bool:
        """Check if a sheet exists in the workbook.

        Args:
            workbook: xlwings Book object.
            sheet_name: Name of the sheet to check.

        Returns:
            True if sheet exists, False otherwise.
        """
        sheet_names = [s.name for s in workbook.sheets]
        return sheet_name in sheet_names

    def _get_sheet(self, workbook: Any, sheet_name: str) -> Any:
        """Get a sheet from the workbook.

        Args:
            workbook: xlwings Book object.
            sheet_name: Name of the sheet to get.

        Returns:
            xlwings Sheet object.
        """
        return workbook.sheets[sheet_name]

    def _write_data_to_sheet(
        self,
        sheet: Any,
        rows: list[tuple],
        start_row: int,
        col_indexes: list[int] | None,
    ) -> None:
        """Write data to the sheet using range assignment.

        Args:
            sheet: xlwings Sheet object.
            rows: Data rows to write.
            start_row: Starting row number (1-indexed).
        """
        if not rows:
            return

        rows = [
            tuple(float(v) if isinstance(v, Decimal) else v for v in row)
            for row in rows
        ]

        if col_indexes is None:
            sheet.range((start_row, 1)).value = rows
            return

        for col_idx, excel_col in enumerate(col_indexes):
            col_values = [[row[col_idx]] for row in rows]
            sheet.range((start_row, excel_col)).resize(len(rows), 1).value = col_values
