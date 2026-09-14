#!/usr/bin/env python3
"""Deterministic manifest/label summary for stakeholder reporting.

Reads ONLY:
  - config/build_manifest.yaml   (active_profiles / defunct_profiles lists)
  - config/transforms/<key>.yaml (transform.label, or shell_defaults.variants[].profile_label
                                  — never touches transform_specs/ decision records)

Outputs partitioned counts + labels for AP (Supplier Invoice) / AR (Customer Invoice) /
Defunct, plus a consolidated view (Amex Trains V1+V2 → 1 line item) on request.

Usage:
    uv run --with pyyaml --no-project python \\
        implementations/elt_transform_excel/scripts/manifest_summary.py [--consolidate-amex-trains] [--json]
    # Or: cd implementations/elt_transform_excel && uv run --with pyyaml --no-project python scripts/manifest_summary.py
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

try:
    import yaml  # type: ignore
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "ERROR: PyYAML required. Run with: uv run --with pyyaml --no-project python ...\n"
    )
    sys.exit(2)


SCRIPT_PATH = Path(__file__).resolve()
MODULE_ROOT = SCRIPT_PATH.parents[1]
CONFIG_DIR = MODULE_ROOT / "config"
MANIFEST_PATH = CONFIG_DIR / "build_manifest.yaml"
TRANSFORMS_DIR = CONFIG_DIR / "transforms"


@dataclass
class Entry:
    profile_key: str
    archetype_key: str
    label: str
    partition: str  # "AP" | "AR" | "DEFUNCT"
    archetype_label: str | None


def _partition_of(key: str, *, is_defunct: bool) -> str:
    if is_defunct:
        return "DEFUNCT"
    if key.startswith("ap_"):
        return "AP"
    if key.startswith("ar_"):
        return "AR"
    return "UNKNOWN"


def _load_yaml(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def _resolve_entries_for(key: str, *, is_defunct: bool) -> list[Entry]:
    """Return one Entry per Runner-dropdown row for a single manifest key.

    Resolution order:
      1. If the transform profile defines shell_defaults.variants -> one Entry per
         uncommented (i.e. actually present in the YAML dict) variant, using
         profile_label.
      2. Otherwise -> single Entry using transform.label.
    """
    partition = _partition_of(key, is_defunct=is_defunct)

    profile_path = TRANSFORMS_DIR / f"{key}.yaml"
    if not profile_path.exists():
        return [
            Entry(
                profile_key=key,
                archetype_key=key,
                label=f"[{key} — MISSING transforms/{key}.yaml]",
                partition=partition,
                archetype_label=None,
            )
        ]

    doc = _load_yaml(profile_path) or {}
    transform_block = doc.get("transform", {}) or {}
    archetype_label = transform_block.get("label") or key

    shell_defaults = doc.get("shell_defaults", {}) or {}
    variants = shell_defaults.get("variants")
    if isinstance(variants, list) and variants:
        entries: list[Entry] = []
        for v in variants:
            if not isinstance(v, dict):
                continue
            vk = v.get("profile_key") or key
            vl = v.get("profile_label") or archetype_label
            entries.append(
                Entry(
                    profile_key=str(vk),
                    archetype_key=key,
                    label=str(vl),
                    partition=partition,
                    archetype_label=str(archetype_label),
                )
            )
        if entries:
            return entries

    return [
        Entry(
            profile_key=key,
            archetype_key=key,
            label=str(archetype_label),
            partition=partition,
            archetype_label=str(archetype_label),
        )
    ]


def _consolidate_amex_trains(entries: list[Entry]) -> list[Entry]:
    """Merge ap_amex_trains + ap_amex_trains_v2 into one consolidated label (stakeholder view)."""
    out: list[Entry] = []
    seen_amex_consolidated = False
    for e in entries:
        if e.archetype_key in {"ap_amex_trains", "ap_amex_trains_v2"} and e.partition == "AP":
            if seen_amex_consolidated:
                continue
            seen_amex_consolidated = True
            out.append(
                Entry(
                    profile_key="ap_amex_trains (consolidated)",
                    archetype_key="ap_amex_trains",
                    label="AP AMEX Trains (V1 + V2, consolidated)",
                    partition="AP",
                    archetype_label="AP AMEX Trains (consolidated)",
                )
            )
            continue
        out.append(e)
    return out


def build() -> dict[str, list[Entry]]:
    manifest = _load_yaml(MANIFEST_PATH) or {}
    active: list[str] = list(manifest.get("active_profiles", []) or [])
    defunct: list[str] = list(manifest.get("defunct_profiles", []) or [])

    result: dict[str, list[Entry]] = {"AP": [], "AR": [], "DEFUNCT": [], "UNKNOWN": []}
    for k in active:
        for e in _resolve_entries_for(k, is_defunct=False):
            result.setdefault(e.partition, []).append(e)
    for k in defunct:
        for e in _resolve_entries_for(k, is_defunct=True):
            result.setdefault(e.partition, []).append(e)
    return result


def print_text(by_part: dict[str, list[Entry]], *, consolidate_amex: bool) -> None:
    ap = list(by_part.get("AP", []))
    ar = list(by_part.get("AR", []))
    df = list(by_part.get("DEFUNCT", []))

    if consolidate_amex:
        ap = _consolidate_amex_trains(ap)

    total = len(ap) + len(ar)

    print(f"TOTAL ACTIVE RUNNER ENTRIES: {total}  ({len(ap)} AP / {len(ar)} AR)")
    if df:
        print(f"DEFUNCT (not in dropdown): {len(df)}")
    print()

    def _emit(title: str, entries: list[Entry]) -> None:
        if not entries:
            return
        print(f"--- {title} ({len(entries)}) ---")
        for i, e in enumerate(entries, 1):
            suffix = ""
            if e.archetype_label and e.label != e.archetype_label:
                suffix = f"  (archetype: {e.archetype_label})"
            print(f"{i:>2}. {e.label}{suffix}")
        print()

    _emit("AP — Supplier Invoice", ap)
    _emit("AR — Customer Invoice", ar)
    _emit("Defunct (retained on disk, not in runner)", df)

    if consolidate_amex:
        print(
            "[Note: Amex Trains V1 + V2 consolidated per stakeholder reporting directive; "
            "run without --consolidate-amex-trains for the raw dropdown count.]"
        )


def print_json(by_part: dict[str, list[Entry]], *, consolidate_amex: bool) -> None:
    out: dict[str, Any] = {}
    for part, entries in by_part.items():
        es = list(entries)
        if part == "AP" and consolidate_amex:
            es = _consolidate_amex_trains(es)
        out[part] = [
            {
                "profile_key": e.profile_key,
                "archetype_key": e.archetype_key,
                "label": e.label,
                "archetype_label": e.archetype_label,
            }
            for e in es
        ]
    ap_n = len(out.get("AP", []))
    ar_n = len(out.get("AR", []))
    out["_summary"] = {
        "total_active": ap_n + ar_n,
        "ap": ap_n,
        "ar": ar_n,
        "defunct": len(out.get("DEFUNCT", [])),
        "consolidated_amex_trains": consolidate_amex,
        "manifest": str(MANIFEST_PATH),
    }
    json.dump(out, sys.stdout, indent=2, sort_keys=False)
    sys.stdout.write("\n")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "--consolidate-amex-trains",
        action="store_true",
        help="Merge ap_amex_trains + ap_amex_trains_v2 into one stakeholder-facing line item.",
    )
    parser.add_argument("--json", action="store_true", help="Emit JSON instead of text.")
    args = parser.parse_args(argv)

    by_part = build()
    if args.json:
        print_json(by_part, consolidate_amex=args.consolidate_amex_trains)
    else:
        print_text(by_part, consolidate_amex=args.consolidate_amex_trains)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
