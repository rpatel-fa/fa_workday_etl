# examples/base_validator.py
"""
Base validator for running VBA validation macros on Excel workbooks.

This module contains the shared logic for running VBA validations.
Individual instance validators should import and call run_validation() with their
specific configuration values.
"""

import argparse

from elt_ingest_excel.macro.vba_runner import run_validation


def create_parser(description: str) -> argparse.ArgumentParser:
    """Create an argument parser with common validation options."""
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "--workbook",
        help="Path to the .xlsm workbook to validate (overrides config default)",
    )
    parser.add_argument(
        "--macro",
        help="Validation macro to run (overrides config default)",
    )
    parser.add_argument(
        "--unhide-sheet",
        help="Sheet to unhide after validation (overrides config default)",
    )
    parser.add_argument(
        "--excel-visible",
        action="store_true",
        help="Show Excel during macro execution",
    )
    parser.add_argument(
        "--no-save",
        action="store_true",
        help="Do not save the workbook after validation",
    )
    parser.add_argument(
        "--close",
        action="store_true",
        help="Close the workbook after validation",
    )
    return parser


def run_validation_wrapper(
    *,
    workbook_path: str,
    macro_name: str = "runAllValidationsFromSheet",
    unhide_sheet: str = "Validation Results",
    excel_visible: bool = False,
    save: bool = True,
    close: bool = False,
):
    """
    Run VBA validation on a workbook.

    Args:
        workbook_path: Path to the .xlsm workbook
        macro_name: Validation macro to run
        unhide_sheet: Sheet to unhide after validation
        excel_visible: Show Excel during macro execution
        save: Save the workbook after validation
        close: Close the workbook after validation
    """
    run_validation(
        workbook_path=workbook_path,
        macro_name=macro_name,
        unhide_sheet=unhide_sheet,
        excel_visible=excel_visible,
        save=save,
        close=close,
    )
