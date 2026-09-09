# Reference export (`fin_workday_ref.py`)

The script [fin_workday_ref.py](examples/fin_workday_ref.py) is an intentional outlier.

Its purpose is to produce an analyst-friendly Excel workbook containing the reference/mapping tables used in the transforms (source values → Workday-approved values).

This is useful for:

- Reviewing and editing mapping inputs with a business user
- Validating that a given source value will land on a valid Workday target value
- Sharing the “contract” of acceptable values without requiring access to DuckDB

## How it works

- Loads [publish_workday_ref.json](config/publish/finance/publish_workday_ref.json) directly (manual JSON parsing)
- Opens the DuckDB database read-only
- For each configured table:
  - Creates a new sheet named `sheetName`
  - Writes a header row + all data rows from `SELECT * FROM <srcTableName>`
  - Applies basic formatting (header styling, auto-filter, freeze panes, auto column widths)
- Saves a new `.xlsx` workbook to the configured output path

## Why it does not use the standard publisher

The standard publish step is designed for “clone a macro-enabled template workbook and fill in data rows”.

The reference export is the opposite:

- It creates a brand-new workbook (no template)
- It adds convenience formatting
- It is intended to be shared and edited, not loaded into Workday directly

