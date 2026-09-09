#!/usr/bin/env python3
"""
Run VBA validation macros in the Workday Supplier/Creditor workbook.

This module provides a simple interface to run the validation macros
that 3+ uses to validate the workbook data.

The validation macros are embedded in the .xlsm workbook and perform
data quality checks defined in the ValidationManifest sheet.

Usage:
    # From command line
    uv run python -m elt_ingest_excel.macro.vba_runner \
        --workbook ~/Documents/workday_fin_creditor_supplier_active_v1.xlsm

    # Run specific validations (single sheet)
    uv run python -m elt_ingest_excel.macro.vba_runner \
        --workbook ~/Documents/workday_fin_creditor_supplier_active_v1.xlsm \
        --macro runSpecificValidationsFromSheet

    # From Python
    from elt_ingest_excel.macro.vba_runner import run_validation

    result = run_validation(
        workbook_path="~/Documents/workday_fin_creditor_supplier_active_v1.xlsm"
    )
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from elt_ingest_excel.macro.excel_macro_runner import run_excel_vba_macro


def run_validation(
    workbook_path: str | Path,
    macro_name: str = "runAllValidationsFromSheet",
    unhide_sheet: str = "Validation Results",
    excel_visible: bool = False,
    save: bool = True,
    close: bool = False,
) -> None:
    """
    Run the VBA validation macro in the workbook.

    This executes the same validation that 3+ runs on the workbook.
    The validation results are written to the "Validation Results" sheet.

    Args:
        workbook_path: Path to the .xlsm workbook
        macro_name: Validation macro to run (default: runAllValidationsFromSheet)
        unhide_sheet: Sheet to unhide after validation (default: Validation Results)
        excel_visible: Show Excel during macro execution (default: False)
        save: Save the workbook after validation (default: True)
        close: Close the workbook after validation (default: False)

    Available macros:
        - runAllValidationsFromSheet: Run all validations (default)
        - runSpecificValidationsFromSheet: Run validations for selected sheet only
        - runValidationsFromRibbon: Entry point for ribbon button

    Example:
        >>> run_validation("~/Documents/workday_fin_creditor_supplier_active_v1.xlsm")
        Validation complete. Check 'Validation Results' sheet for issues.
    """
    path = Path(workbook_path).expanduser().resolve()

    if not path.exists():
        raise FileNotFoundError(f"Workbook not found: {path}")

    run_excel_vba_macro(
        workbook_path=path,
        macro_name=f"{path.name}!{macro_name}",
        unhide_sheets=[unhide_sheet],
        save=save,
        close=close,
        excel_visible=excel_visible,
    )

    print(f"Validation complete. Check '{unhide_sheet}' sheet for issues.")


def main() -> int:
    """Main entry point for command-line usage."""
    parser = argparse.ArgumentParser(
        description="Run VBA validation macros in Workday Supplier/Creditor workbook"
    )
    parser.add_argument(
        "--workbook",
        required=True,
        help="Path to the .xlsm workbook",
    )
    parser.add_argument(
        "--list-macros",
        action="store_true",
        help="List available macros in the workbook (no execution)",
    )
    parser.add_argument(
        "--macro",
        default="runAllValidationsFromSheet",
        choices=[
            "runAllValidationsFromSheet",
            "runSpecificValidationsFromSheet",
            "runValidationsFromRibbon",
        ],
        help="Validation macro to run (default: runAllValidationsFromSheet)",
    )
    parser.add_argument(
        "--unhide-sheet",
        default="Validation Results",
        help="Sheet to unhide after validation (default: Validation Results)",
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

    args = parser.parse_args()

    try:
        # List macros mode
        if args.list_macros:
            from elt_ingest_excel.macro.vba_analyzer import VbaMacroAnalyzer

            analyzer = VbaMacroAnalyzer(args.workbook)
            result = analyzer.analyze()

            print(f"Macros in {Path(args.workbook).name}:")
            print("-" * 50)
            for macro in result.macros:
                print(f"  {macro.macro_type:8} {macro.name}")
            print(f"\nTotal: {len(result.macros)} macros")

            if result.validation_sheets:
                print("\nValidation-related sheets/ranges:")
                for name in result.validation_sheets:
                    print(f"  • {name}")

            return 0

        # Run validation mode
        run_validation(
            workbook_path=args.workbook,
            macro_name=args.macro,
            unhide_sheet=args.unhide_sheet,
            excel_visible=args.excel_visible,
            save=not args.no_save,
            close=args.close,
        )
        return 0

    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except RuntimeError as e:
        print(f"Error running macro: {e}", file=sys.stderr)
        return 2
    except Exception as e:
        print(f"Unexpected error: {e}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
