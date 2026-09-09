"""
VBA Macro Analyzer for Excel Workbooks.

This module provides tools to extract and document VBA macro names
from Excel .xlsm, .xlam, and .xlsb files.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class VbaMacro:
    """Represents a VBA macro (Sub or Function)."""

    name: str
    macro_type: str  # "Sub" or "Function"


@dataclass
class VbaAnalysisResult:
    """Result of VBA analysis on a workbook."""

    workbook_path: str
    macros: list[VbaMacro] = field(default_factory=list)
    sheet_names: list[str] = field(default_factory=list)
    named_ranges: list[str] = field(default_factory=list)
    validation_sheets: list[str] = field(default_factory=list)


class VbaMacroAnalyzer:
    """Analyzer for VBA macros in Excel workbooks."""

    def __init__(self, workbook_path: str | Path) -> None:
        """Initialize the analyzer with a workbook path."""
        self.workbook_path = Path(workbook_path).expanduser().resolve()
        if not self.workbook_path.exists():
            raise FileNotFoundError(str(self.workbook_path))

        valid_suffixes = {".xlsm", ".xlam", ".xlsb"}
        if self.workbook_path.suffix.lower() not in valid_suffixes:
            raise ValueError(
                f"Expected workbook with suffix {valid_suffixes}, "
                f"got '{self.workbook_path.suffix}'"
            )

    def analyze(self) -> VbaAnalysisResult:
        """Perform comprehensive analysis of the workbook."""
        result = VbaAnalysisResult(workbook_path=str(self.workbook_path))

        with zipfile.ZipFile(self.workbook_path) as zf:
            result.macros = self._extract_macros(zf)
            result.sheet_names = self._extract_sheet_names(zf)
            result.named_ranges = self._extract_named_ranges(zf)
            result.validation_sheets = self._find_validation_sheets(
                result.sheet_names, result.named_ranges
            )

        return result

    def _extract_vba_strings(self, zf: zipfile.ZipFile, min_len: int = 2) -> list[str]:
        """Extract printable strings from VBA project."""
        try:
            vba = zf.read("xl/vbaProject.bin")
        except KeyError:
            return []

        strings: list[str] = []
        current: list[int] = []

        for b in vba:
            if 32 <= b <= 126:
                current.append(b)
            else:
                if len(current) >= min_len:
                    strings.append(bytes(current).decode("ascii", errors="ignore"))
                current = []

        if len(current) >= min_len:
            strings.append(bytes(current).decode("ascii", errors="ignore"))

        return strings

    def _extract_macros(self, zf: zipfile.ZipFile) -> list[VbaMacro]:
        """Extract VBA macro definitions from the workbook."""
        strings = self._extract_vba_strings(zf, min_len=2)
        txt = "\n".join(strings)

        macros: list[VbaMacro] = []
        seen: set[str] = set()

        # Known validation macro patterns
        known_patterns = [
            r"runSpecificValidationsFromSheet",
            r"runAllValidationsFromSheet",
            r"runValidationsFromRibbon",
            r"deleteValidationSheet",
            r"Worksheet_Change",
            r"exportWorkbooks",
        ]

        for pattern in known_patterns:
            if re.search(pattern, txt, re.IGNORECASE):
                match = re.search(rf"({pattern}[A-Za-z0-9_]*)", txt, re.IGNORECASE)
                if match:
                    name = match.group(1)
                    if name and name not in seen:
                        seen.add(name)
                        macros.append(VbaMacro(name=name, macro_type="Sub"))

        # Also extract from Sub/Function patterns
        pattern = r"\b(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(|:|\n|$)"

        for match in re.finditer(pattern, txt, re.IGNORECASE):
            macro_type = match.group(1)
            macro_name = match.group(2)

            if macro_name.lower() in {"auto_open", "auto_close"}:
                continue
            if macro_name in seen:
                continue
            if len(macro_name) < 3:
                continue

            seen.add(macro_name)
            macros.append(
                VbaMacro(name=macro_name, macro_type=macro_type.capitalize())
            )

        return sorted(macros, key=lambda m: (m.macro_type, m.name.lower()))

    def _extract_sheet_names(self, zf: zipfile.ZipFile) -> list[str]:
        """Extract worksheet names from workbook.xml."""
        try:
            wb_xml = zf.read("xl/workbook.xml").decode("utf-8", errors="ignore")
        except KeyError:
            return []

        names = re.findall(r'name="([^"]+)"', wb_xml)
        return sorted([n for n in names if not n.startswith("_xlnm.")])

    def _extract_named_ranges(self, zf: zipfile.ZipFile) -> list[str]:
        """Extract named range definitions from workbook.xml."""
        try:
            wb_xml = zf.read("xl/workbook.xml").decode("utf-8", errors="ignore")
        except KeyError:
            return []

        names = re.findall(r'name="([^"]+)"', wb_xml)
        return sorted([n for n in names if not n.startswith("_xlnm.")])

    def _find_validation_sheets(
        self, sheet_names: list[str], named_ranges: list[str]
    ) -> list[str]:
        """Identify validation-related sheets and ranges."""
        keywords = ["valid", "manifest", "result"]
        all_names = sheet_names + named_ranges
        return sorted(set(n for n in all_names if any(k in n.lower() for k in keywords)))


def list_vba_entry_points(workbook_path: str | Path) -> list[str]:
    """List VBA macro entry points in a workbook."""
    analyzer = VbaMacroAnalyzer(workbook_path)
    result = analyzer.analyze()
    return [macro.name for macro in result.macros]
