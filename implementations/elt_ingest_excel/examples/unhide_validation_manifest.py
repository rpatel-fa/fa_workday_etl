# examples/unhide_validation_manifest.py
"""Unhide the ValidationManifest sheet in the workbook."""

import argparse
from pathlib import Path

from elt_ingest_excel.macro.excel_macro_runner import run_excel_vba_macro

# Configuration
WORKBOOK_PATH = "~/Documents/workday_fin_creditor_supplier_active_v1.xlsm"
SHEETS_TO_UNHIDE = ["ValidationManifest", "Validation Results"]


def create_parser(description: str) -> argparse.ArgumentParser:
    """Create argument parser."""
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "--workbook",
        default=WORKBOOK_PATH,
        help=f"Path to the .xlsm workbook (default: {WORKBOOK_PATH})",
    )
    parser.add_argument(
        "--sheets",
        nargs="+",
        default=SHEETS_TO_UNHIDE,
        help=f"Sheets to unhide (default: {SHEETS_TO_UNHIDE})",
    )
    parser.add_argument(
        "--excel-visible",
        action="store_true",
        help="Show Excel during macro execution",
    )
    return parser


if __name__ == "__main__":
    parser = create_parser("Unhide ValidationManifest sheet in Finance Supplier/Creditor workbook")
    args = parser.parse_args()

    wb_path = Path(args.workbook).expanduser().resolve()

    print(f"Unhiding sheets: {args.sheets}")
    print(f"Workbook: {wb_path}")
    print()

    run_excel_vba_macro(
        workbook_path=wb_path,
        macro_name=f"{wb_path.name}!Unhide_Multiple_Sheets",
        unhide_sheets=args.sheets,
        save=True,
        close=False,
        excel_visible=args.excel_visible,
    )

    print("✓ Sheets are now visible")
    print()
    print("To add validation rules:")
    print("  1. Open the 'ValidationManifest' sheet in Excel")
    print("  2. Add rows for 'Supplier Name' with validation rules")
    print("  3. Save the workbook")
