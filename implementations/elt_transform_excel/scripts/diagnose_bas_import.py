#!/usr/bin/env python3
"""Per-file deep VBA hygiene diagnostic that mirrors what Excel/VBProject.Import
chokes on at import time.  Run this on a specific .bas file to see why
BootstrapImporter is raising 80040714 on that file specifically."""

from __future__ import annotations

import argparse
import pathlib
import re
import sys


def analyse(path: pathlib.Path) -> int:
    raw = path.read_bytes()
    errors: list[str] = []
    warnings: list[str] = []

    print(f"File: {path}  ({len(raw)} bytes)")

    # 1. BOMs / leading whitespace-bytes
    if raw.startswith(b"\xef\xbb\xbf"):
        errors.append(f"UTF-8 BOM at start of file (Excel VBE Import BOM-UTF-8 unsupported for .BAS; "
                      f"VBComponents.Import fails with 80040714 on BOM-prefixed .BAS files). "
                      f"Strip with: sed -i '1s/^\\xEF\\xBB\\xBF//' or python strip.")
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        errors.append("UTF-16 BOM at start of file; .BAS imports are ASCII/UTF-8 plain only.")

    # 2. Component-name length limit (31 chars for VBProject.VBComponents)
    component_name = path.stem
    if len(component_name) > 31:
        errors.append(f"Component name '{component_name}' is {len(component_name)} chars; "
                      f"VBComponent collection has 31-char limit.")

    # 3. Attribute block: for .BAS files VBProject.Import demands the file begin
    #    with Attribute VB_Name = "...".  If the file has NO Attribute header or
    #    it appears after some blank line / comment block, Import on some
    #    versions of Excel (Mac especially) raises the opaque 80040714 error.
    #    Check 32-line header.
    header_lines = raw.split(b"\n")[:32]
    has_vb_name_attr = False
    bad_attribute_positions: list[int] = []
    for i, ln in enumerate(header_lines, 1):
        stripped = ln.strip()
        if stripped.startswith(b"Attribute"):
            if b"VB_Name" in stripped:
                has_vb_name_attr = True
            if i > 8:
                bad_attribute_positions.append(i)
    if not has_vb_name_attr:
        errors.append("No 'Attribute VB_Name = \"...\"' line found in first 32 lines. "
                      "Excel VBComponents.Import() REQUIRES this header on standard modules; "
                      "Mac Excel in particular rejects .BAS files without it. "
                      "Fix: insert 'Attribute VB_Name = \"modApStationerySupplierInvoice\"' as "
                      "the very first line (before any ' comment header).")
    if bad_attribute_positions:
        warnings.append(f"Attribute blocks appear late in the header (lines {bad_attribute_positions}). "
                        "On Mac Excel, VBComponents.Import sometimes rejects files where the first "
                        "Attribute line is after a large comment block. Move the Attribute VB_Name line "
                        "to line 1 for robustness.")

    # 4. Non-ASCII (same rule as check_vba_hygiene.py, but we show exact lines).
    non_ascii_hits: list[tuple[int, int, int]] = []  # (line, col, byte)
    line_no = 1
    col_no = 1
    for i, b in enumerate(raw):
        if b == 0x0A:
            line_no += 1
            col_no = 1
            continue
        if b > 127:
            non_ascii_hits.append((line_no, col_no, b))
        col_no += 1
    if non_ascii_hits:
        errors.append(f"{len(non_ascii_hits)} NON-ASCII byte(s) found. Excel's VBComponents.Import() "
                      "on Mac treats multi-byte UTF-8 inside .BAS files as corrupt: same 80040714. "
                      "First few:")
        for (ln, col, byte) in non_ascii_hits[:8]:
            errors.append(f"  line {ln}, col {col}  byte 0x{byte:02X}")

    # 5. Long lines
    long_line_hits: list[tuple[int, int]] = []
    for ln_idx, ln in enumerate(raw.split(b"\n"), 1):
        if len(ln) > 250:
            long_line_hits.append((ln_idx, len(ln)))
    if long_line_hits:
        warnings.append(f"{len(long_line_hits)} line(s) exceed 250 chars. Lines longer than ~1024 bytes "
                        "can cause Import() to fail silently on Mac Excel. Longest 5:")
        for ln, size in sorted(long_line_hits, key=lambda x: -x[1])[:5]:
            warnings.append(f"  line {ln}: {size} chars")

    # 6. CRLF / bare CR inconsistency
    cr = raw.count(b"\r")
    lf = raw.count(b"\n")
    crlf = raw.count(b"\r\n")
    bare_cr = cr - crlf
    if bare_cr != 0:
        warnings.append(f"Bare CR (not part of CRLF) detected: {bare_cr} occurrence(s). "
                        f"Mac VBProject.Import() historically mishandles mixed line endings. "
                        f"Normalise to LF-only with `dos2unix` or similar.")

    # 7. Trailing control / NUL bytes (a write-encoding mistake that breaks Import()).
    trailing_bad = raw.rstrip(b"\n")[-8:] if len(raw) >= 8 else raw
    if any(b == 0x00 for b in trailing_bad):
        errors.append("NUL byte found at end of file. Corrupt write/encoding artifact.")
    if re.search(rb"Attribute\s+VB_", raw) is None and b"Attribute" not in raw[:200]:
        warnings.append("No VB_* attribute declarations at all (not even VB_Name). "
                        "Standard .BAS files exported from VBE carry at least Attribute VB_Name; "
                        "without it, Mac Excel VBComponents.Import() frequently raises 80040714.")

    print("---")
    if warnings:
        print(f"Warnings ({len(warnings)}):")
        for w in warnings:
            print(f"  ! {w}")
    if errors:
        print(f"Errors ({len(errors)}):")
        for e in errors:
            print(f"  * {e}")
        print("\nLikely cause for the 80040714 VBComponents.Import() error listed above.")
        return 2
    else:
        print("No structural/encoding issues detected.  If Import() still fails, suspect Excel's "
              "'Trust access to VBA project object model' toggle or macOS sandbox.")
        return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("bas_file", type=pathlib.Path, help="Path to a .bas/.cls file to diagnose")
    args = parser.parse_args()
    if not args.bas_file.is_file():
        print(f"File not found: {args.bas_file}", file=sys.stderr)
        return 1
    return analyse(args.bas_file)


if __name__ == "__main__":
    raise SystemExit(main())
