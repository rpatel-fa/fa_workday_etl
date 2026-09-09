# examples/fin_supplier_creditor_create_eib.py
"""Run the EIB workbook creation macro for Finance Supplier/Creditor workbook.

The TempCreator VBA module is marked Option Private Module, which blocks
external macro calls. This script patches that flag out of the workbook binary
automatically (one-time, saves a .bak backup) before running the macro.
"""

import argparse
import subprocess
import sys
from pathlib import Path

from elt_ingest_excel.macro.excel_macro_runner import _escape_applescript_string
from elt_ingest_excel.macro.vba_patcher import remove_module_private

# Configuration
WORKBOOK_PATH = "~/Documents/workday_fin_creditor_supplier_active_v1.xlsm"
MACRO_NAME = "InitiateALLAdvanceLoad"
UNHIDE_SHEET = "EIB View"
PRIVATE_MODULE = "TempCreator"


def create_parser(description: str) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "--workbook",
        default=WORKBOOK_PATH,
        help=f"Path to the .xlsm workbook (default: {WORKBOOK_PATH})",
    )
    parser.add_argument(
        "--macro",
        default=MACRO_NAME,
        help=f"Macro to run (default: {MACRO_NAME})",
    )
    parser.add_argument(
        "--excel-visible",
        action="store_true",
        help="Show Excel during macro execution",
    )
    parser.add_argument(
        "--no-save",
        action="store_true",
        help="Do not save the workbook after macro execution",
    )
    parser.add_argument(
        "--close",
        action="store_true",
        help="Close the workbook after macro execution",
    )
    parser.add_argument(
        "--skip-patch",
        action="store_true",
        help="Skip the Option Private Module patch (use if already patched)",
    )
    return parser


if __name__ == "__main__":
    parser = create_parser("Run EIB workbook creation macro for Finance Supplier/Creditor workbook")
    args = parser.parse_args()

    wb_path = Path(args.workbook).expanduser().resolve()

    if not wb_path.exists():
        print(f"Error: workbook not found: {wb_path}", file=sys.stderr)
        sys.exit(1)

    # Patch Option Private Module out of TempCreator (one-time, idempotent)
    if not args.skip_patch:
        try:
            remove_module_private(wb_path, PRIVATE_MODULE)
        except ValueError as e:
            # Already patched or module not found - fine to continue
            if "no MODULEPRIVATE record" in str(e) or "already public" in str(e).lower():
                print(f"Note: {e} — skipping patch")
            else:
                print(f"Patch warning: {e}", file=sys.stderr)
        except Exception as e:
            print(f"Warning: could not patch workbook: {e}", file=sys.stderr)
            print("Attempting to run macro anyway...")

    print(f"Workbook: {wb_path}")
    print(f"Macro:    {args.macro}")
    print()

    try:
        _run_eib_macro(wb_path, args.macro, args.excel_visible, not args.no_save)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(2)


def _run_eib_macro(wb_path: Path, macro_name: str, excel_visible: bool, save: bool) -> None:
    """
    Open the workbook, activate EIB View, run the macro, then save any new
    workbooks that appeared (the exported EIB files).

    The macro uses WScript.Shell (Windows-only) to save output files, which
    fails on macOS.  This AppleScript catches that error gracefully and saves
    any newly created workbooks to ~/Desktop instead.
    """
    wb_name = wb_path.name
    p = _escape_applescript_string(str(wb_path))
    m = _escape_applescript_string(f"{wb_name}!{macro_name}")
    visible = "true" if excel_visible else "false"
    save_stmt = "save active workbook" if save else ""

    script = f"""
tell application "Microsoft Excel"
    activate
    set display alerts to false
    set visible to {visible}

    -- Note which workbooks are already open
    set existingNames to {{}}
    repeat with wb in workbooks
        set end of existingNames to (name of wb)
    end repeat

    -- Open the target workbook
    set targetWB to open workbook workbook file name (POSIX file "{p}")

    -- Activate EIB View sheet so the macro operates on the right sheet
    try
        activate object worksheet "EIB View" of targetWB
    end try

    -- Run the macro; ignore WScript.Shell errors (Windows-only SaveAs path)
    try
        run VB macro "{m}"
    end try

    -- Save any newly opened workbooks (the EIB output files) to ~/Desktop
    set savedFiles to {{}}
    repeat with wb in workbooks
        if (name of wb) is not in existingNames and (name of wb) is not equal to "{_escape_applescript_string(wb_name)}" then
            set wbName to name of wb
            set outPath to (POSIX file (POSIX path of (path to desktop))) & wbName & ".xlsx"
            try
                save workbook as wb filename outPath file format xlsx file format
                close wb saving no
                set end of savedFiles to outPath
            on error
                -- Leave open for manual saving if save fails
            end try
        end if
    end repeat

    {save_stmt}
end tell
"""

    proc = subprocess.run(["osascript", "-e", script], text=True, capture_output=True)
    if proc.returncode != 0:
        err = (proc.stderr or "").strip()
        raise RuntimeError(err or f"osascript failed (code {proc.returncode})")

    print("Done. EIB files saved to ~/Desktop (or left open in Excel if save failed).")
