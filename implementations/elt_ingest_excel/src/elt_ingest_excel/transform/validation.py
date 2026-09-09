from pathlib import Path
from typing import Iterable

import duckdb
import pandas as pd


def _normalize(values: Iterable[str]) -> set[str]:
    out: set[str] = set()
    for v in values:
        if v is None:
            continue
        s = str(v).strip()
        if s:
            out.add(s.casefold())
    return out


def validate_counties_against_master(
    database_path: Path,
    master_workbook_path: Path,
    sheet_name: str = "Country States-Regions",
) -> tuple[bool, set[str], set[str]]:
    if not master_workbook_path.exists():
        return True, set(), set()
    with duckdb.connect(str(database_path)) as conn:
        rows = conn.execute(
            """
            SELECT DISTINCT county
            FROM ref_post_code_county
            WHERE NULLIF(TRIM(county),'') IS NOT NULL
              AND LOWER(country_name) IN ('england','scotland','wales','northern ireland')
            """
        ).fetchall()
    counties = [r[0] for r in rows]
    try:
        df = pd.read_excel(
            str(master_workbook_path),
            sheet_name=sheet_name,
            header=2,
            engine=None,
        )
    except Exception:
        return True, set(), set()
    df = df[["Country", "Instance"]].dropna(subset=["Country", "Instance"])
    df["Country_norm"] = df["Country"].astype(str).str.strip().str.casefold()
    uk = df[df["Country_norm"] == "united kingdom"]["Instance"].astype(str).str.strip()
    allowed_list: list[str] = uk.tolist()
    allowed_norm = _normalize(allowed_list)
    # Also allow base forms of '(obsolete)' entries
    obsolete_suffix = " (obsolete)"
    for inst in allowed_list:
        if inst.endswith(obsolete_suffix):
            allowed_norm.add(inst[: -len(obsolete_suffix)].strip().casefold())
    current = _normalize(counties)
    missing = {c for c in current if c not in allowed_norm}
    ok = len(missing) == 0
    return ok, missing, allowed_norm
