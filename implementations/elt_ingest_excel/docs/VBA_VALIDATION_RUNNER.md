# VBA Validation Runner

Run and analyze VBA validation macros in your Workday Supplier/Creditor Excel workbook.

## Overview

This module provides tools to run the validation macros that 3+ uses to validate workbooks.

## Quick Start

### One-time Setup: Unhide ValidationManifest

The ValidationManifest sheet contains validation rules but is hidden by default.

```bash
cd elt_ingest_excel

# Unhide ValidationManifest sheet
uv run python examples/unhide_validation_manifest.py --excel-visible
```

Then in Excel:
1. Open the `ValidationManifest` sheet (at the end of sheet tabs)
2. Add validation rules for your data sheets
3. Save the workbook

### Run Validation

```bash
cd elt_ingest_excel

# Run all validations (default)
uv run python examples/fin_supplier_creditor_validate_vba.py

# Run specific sheet validation only
uv run python examples/fin_supplier_creditor_validate_vba.py \
    --macro runSpecificValidationsFromSheet

# Run with Excel visible (to watch progress)
uv run python examples/fin_supplier_creditor_validate_vba.py \
    --excel-visible
```

### Override Defaults

```bash
# Different workbook
uv run python examples/fin_supplier_creditor_validate_vba.py \
    --workbook /path/to/other_workbook.xlsm

# Different macro
uv run python examples/fin_supplier_creditor_validate_vba.py \
    --macro runAllValidationsFromSheet
```

### Your Complete Workflow

```bash
# Step 1: Run your ETL pipeline (unchanged)
uv run python examples/fin_supplier_creditor.py --run-to-phase PUBLISH

# Step 2: Run validation (NEW - removes 3+ latency)
uv run python -m elt_ingest_excel.macro.vba_runner \
    --workbook ~/Documents/workday_fin_creditor_supplier_active_v1.xlsm

# Step 3: Review results
# Open the workbook and check the "Validation Results" sheet
# Fix any SQL transforms if needed, then repeat from Step 1
```

## Available Macros

| Macro | Description |
|-------|-------------|
| `runAllValidationsFromSheet` | Run all validations across all sheets (default) |
| `runSpecificValidationsFromSheet` | Run validations for selected sheet only |
| `runValidationsFromRibbon` | Entry point for Excel ribbon button |

## Validation Output

The validation writes results to the **"Validation Results"** sheet:

| Column | Description |
|--------|-------------|
| sheet | Which sheet has the issue |
| rule_type | DUPLICATE, MISSING, MATCH, or ENSURE_DATA |
| col | Which column failed validation |
| supplier_id | Related supplier ID |
| detail | The problematic value |
| message | Human-readable error message |

## Validation Rule Types

| Rule Type | Description |
|-----------|-------------|
| **DUPLICATE** | Checks for duplicate values in unique key columns |
| **MISSING** | Checks that required fields are populated |
| **MATCH** | Validates values against allowed drop-down lists |
| **ENSURE_DATA** | Conditional requirements (if X is set, Y must also be set) |

## Command-Line Options

```
--workbook WORKBOOK   Path to the .xlsm workbook (required)
--list-macros         List available macros (no execution)
--macro MACRO         Validation macro to run (default: runAllValidationsFromSheet)
--unhide-sheet NAME   Sheet to unhide after validation (default: Validation Results)
--excel-visible       Show Excel during macro execution
--no-save             Do not save the workbook after validation
--close               Close the workbook after validation
```

## Python API

### Run Validation

```python
from elt_ingest_excel.macro import run_validation

# Run all validations
run_validation(
    workbook_path="~/Documents/workday_fin_creditor_supplier_active_v1.xlsm"
)

# Run specific sheet validation
run_validation(
    workbook_path="~/Documents/workday_fin_creditor_supplier_active_v1.xlsm",
    macro_name="runSpecificValidationsFromSheet",
    excel_visible=True,
)
```

### List Macros

```python
from elt_ingest_excel.macro import list_vba_entry_points

macros = list_vba_entry_points(
    workbook_path="~/Documents/workday_fin_creditor_supplier_active_v1.xlsm"
)
print(f"Found {len(macros)} macros: {macros}")
```

## Requirements

- **macOS** - Uses AppleScript to control Excel
- **Microsoft Excel** - Must be installed
- **Macros enabled** - The workbook must allow macro execution

## Module Structure

```
elt_ingest_excel/src/elt_ingest_excel/macro/
├── __init__.py            # Module exports
├── vba_runner.py          # Run validation macros
├── vba_analyzer.py        # List/analyze macros
└── excel_macro_runner.py  # Low-level AppleScript execution
```
