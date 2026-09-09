#!/usr/bin/env python3
import sys, pathlib, importlib.util

ROOT = pathlib.Path(__file__).resolve().parent.parent

spec = importlib.util.spec_from_file_location("build_runner", ROOT / "scripts" / "build_runner.py")
br = importlib.util.module_from_spec(spec)
sys.modules["build_runner"] = br
spec.loader.exec_module(br)

profile_dir = ROOT / "config" / "transforms"
eib_defaults = br._load_eib_defaults(profile_dir.parent / "eib_population_defaults.yaml")
print("EIB defaults OK:", sorted(eib_defaults.keys()))
for k, v in eib_defaults.items():
    print(f"  {k:20s} -> sheet={v['target_sheet_name']!r} start_row={v['target_data_start_row']} end_col={v['clear_published_rows_end_col']!r}")

bindings = br._load_transform_bindings(profile_dir, br.DEFAULT_BUILD_MANIFEST)
print(f"\nBindings parsed: {len(bindings)}")
for b in sorted(bindings, key=lambda x: x.profile_key):
    wr = b.worksheet_resolution or "???"
    hd = b.header_detection or "???"
    ws = f'"{b.worksheet_name}"' if b.worksheet_name else "<first/only>"
    rh = len(b.required_headers) if b.required_headers else 0
    ep = b.eib_population_type or "???"
    print(f"  {b.profile_key:22s} | ep={ep:18s} | {wr:11s} + {hd:20s} | ws={ws:34s} | header={b.header_row} data={b.data_start_row} | rh={rh}")

# Cross-check eib_population_type references a valid key
print("\nCross-check eib_population_type against defaults:")
for b in sorted(bindings, key=lambda x: x.profile_key):
    ok = "OK" if b.eib_population_type in eib_defaults else "MISSING"
    print(f"  {b.profile_key:22s} {b.eib_population_type!r:22s} -> {ok}")
