"""Verify INGEST output tables for supplier_purchase_order_receipt."""

from __future__ import annotations

import json
import os
import zipfile
from pathlib import Path
from typing import Iterable

import duckdb


def _list_xlsx_sheets(path: str) -> list[str]:
    with zipfile.ZipFile(path) as z:
        xml = z.read("xl/workbook.xml")

    import xml.etree.ElementTree as ET

    root = ET.fromstring(xml)
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    sheets = root.find("m:sheets", ns)
    if sheets is None:
        return []
    return [s.attrib.get("name", "") for s in sheets.findall("m:sheet", ns)]


def _chunks(xs: Iterable[str], n: int) -> list[list[str]]:
    out: list[list[str]] = []
    chunk: list[str] = []
    for x in xs:
        chunk.append(x)
        if len(chunk) >= n:
            out.append(chunk)
            chunk = []
    if chunk:
        out.append(chunk)
    return out


if __name__ == "__main__":
    repo_root = Path(__file__).parent.parent.resolve()
    cfg_path = repo_root / "config" / "ingest" / "finance" / "supplier_purchase_order_receipt.json"
    runtime_cfg_path = repo_root / "config" / "runtime" / "default.json"

    cfg = json.loads(cfg_path.read_text())
    workbook_cfg = cfg[0]
    workbook_path = os.path.expanduser(workbook_cfg["workbookPathName"])
    workbook_file = workbook_cfg["workbookFileName"]
    workbook_full_path = str(Path(workbook_path) / workbook_file)

    print("ingest_config:", str(cfg_path))
    print("workbook:", workbook_full_path)

    config_sheets = workbook_cfg["sheets"]
    expected_tabs = [s["sheetName"] for s in config_sheets]
    expected_tables = [s["targetTableName"] for s in config_sheets]

    xlsx_tabs = _list_xlsx_sheets(workbook_full_path)
    xlsx_tab_set = set(xlsx_tabs)
    expected_tab_set = set(expected_tabs)

    print("")
    print("Workbook tab presence (configured tabs only):")
    missing_tabs = [t for t in expected_tabs if t not in xlsx_tab_set]
    if missing_tabs:
        for tab in missing_tabs:
            print("  MISSING_TAB", tab)
    else:
        print("  OK all configured tabs exist in workbook")

    extra_configured = [t for t in expected_tabs if expected_tabs.count(t) > 1]
    if extra_configured:
        for c in sorted(set(extra_configured)):
            print("  DUPLICATE_CONFIG_TAB", c)

    print("")
    print("Workbook tabs (for reference):")
    for chunk in _chunks(xlsx_tabs, 4):
        print("  " + " | ".join(chunk))

    runtime_cfg = json.loads(runtime_cfg_path.read_text())
    db_path = os.path.expanduser(runtime_cfg["databasePath"])

    print("")
    print("db_path:", db_path)
    con = duckdb.connect(db_path, read_only=True)
    present_tables = set(
        r[0]
        for r in con.execute(
            "select table_name from information_schema.tables where table_schema='main'"
        ).fetchall()
    )

    print("")
    print("DuckDB table presence + row counts (configured tables only):")
    missing_tables = [t for t in expected_tables if t not in present_tables]
    for tab, tbl in zip(expected_tabs, expected_tables, strict=True):
        if tbl not in present_tables:
            print("  MISSING_TABLE", tbl, "(tab:", tab + ")")
            continue
        row_count = con.execute(f"select count(*) from {tbl}").fetchone()[0]
        col_count = con.execute(
            f"select count(*) from information_schema.columns where table_schema='main' and table_name='{tbl}'"
        ).fetchone()[0]
        print(f"  OK {tbl}: rows={row_count:,} cols={col_count} (tab: {tab})")

    if missing_tables or missing_tabs:
        raise SystemExit(2)

