"""
Run VBA macros in Excel workbooks (macOS only).

This module provides low-level AppleScript-based execution of VBA macros
in Excel workbooks on macOS.
"""

import os
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Iterable


def list_vba_entry_points(workbook_path: str | Path) -> list[str]:
    """
    List VBA macro entry points in a workbook.

    This extracts macro names from the compiled VBA project binary.
    Note: Results may include some false positives from string fragments.

    Args:
        workbook_path: Path to the .xlsm workbook

    Returns:
        List of macro names
    """
    path = Path(workbook_path).expanduser()
    if not path.exists():
        raise FileNotFoundError(str(path))
    if path.suffix.lower() not in {".xlsm", ".xlam", ".xlsb"}:
        return []

    with zipfile.ZipFile(path) as zf:
        try:
            vba = zf.read("xl/vbaProject.bin")
        except KeyError:
            return []

    txt = "\n".join(_iter_printable_strings(vba, min_len=4))
    names: set[str] = set()

    for m in re.finditer(r"\b(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\b", txt):
        n = m.group(1)
        if n and n.lower() not in {"auto_open", "auto_close"}:
            names.add(n)

    return sorted(names, key=str.lower)


def _iter_printable_strings(data: bytes, min_len: int = 4) -> Iterable[str]:
    """Extract printable ASCII strings from binary data."""
    current: list[int] = []
    for b in data:
        if 32 <= b <= 126:
            current.append(b)
            continue
        if len(current) >= min_len:
            yield bytes(current).decode("ascii", errors="ignore")
        current = []
    if len(current) >= min_len:
        yield bytes(current).decode("ascii", errors="ignore")


def run_excel_vba_macro(
    workbook_path: str | Path,
    macro_name: str,
    unhide_sheets: Iterable[str] = ("Validation Results",),
    save: bool = True,
    close: bool = True,
    excel_visible: bool = False,
) -> None:
    """
    Run a VBA macro in an Excel workbook.

    Args:
        workbook_path: Path to the .xlsm workbook
        macro_name: Name of the macro to run
        unhide_sheets: Sheets to unhide after macro execution
        save: Save the workbook after macro execution
        close: Close the workbook after macro execution
        excel_visible: Show Excel during macro execution
    """
    if sys.platform != "darwin":
        raise RuntimeError("Excel VBA macro runner is only implemented for macOS")

    path = Path(workbook_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(str(path))

    macro = macro_name.strip()
    macro_candidates: list[str] = []

    if "!" in macro:
        macro_candidates.append(macro)
        macro = macro.split("!", 1)[1].strip()
    if not macro:
        raise ValueError("macro_name is empty")
    macro_candidates.append(macro)
    workbook_file_name = path.name
    macro_candidates.append(f"{workbook_file_name}!{macro}")

    sheet_names = [s for s in (str(x).strip() for x in unhide_sheets) if s]
    sheet_list = "{" + ",".join(f"\"{_escape_applescript_string(s)}\"" for s in sheet_names) + "}"
    macro_list = "{" + ",".join(f"\"{_escape_applescript_string(m)}\"" for m in dict.fromkeys(macro_candidates)) + "}"

    script = f"""
tell application "Microsoft Excel"
    activate
    set display alerts to false
    set visible to {"true" if excel_visible else "false"}
    open (POSIX file "{_escape_applescript_string(str(path))}")
    set macroCandidates to {macro_list}
    set macroRan to false
    repeat with m in macroCandidates
        try
            run VB macro (contents of m)
            set macroRan to true
            exit repeat
        end try
    end repeat
    if macroRan is false then
        error ("Cannot run VB macro. Tried: " & macroCandidates as string)
    end if
    repeat with sheetName in {sheet_list}
        try
            set visible of worksheet (contents of sheetName) of active workbook to true
        end try
    end repeat
    if {"true" if save else "false"} then
        save active workbook
    end if
    if {"true" if close else "false"} then
        close active workbook saving no
    end if
end tell
"""
    _run_osascript(script)


def _run_osascript(script: str) -> None:
    """Execute an AppleScript command."""
    proc = subprocess.run(
        ["osascript", "-e", script],
        text=True,
        capture_output=True,
    )
    if proc.returncode != 0:
        out = (proc.stdout or "").strip()
        err = (proc.stderr or "").strip()
        msg = "\n".join([x for x in [out, err] if x])
        raise RuntimeError(msg or f"osascript failed with code {proc.returncode}")


def _escape_applescript_string(value: str) -> str:
    """Escape a string for use in AppleScript."""
    return value.replace("\\", "\\\\").replace('"', '\\"')
