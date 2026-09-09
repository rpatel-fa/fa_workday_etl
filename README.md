# fa_workday_etl

ETL implementations for migrating Football Association finance data into Workday, extracted
from a larger internal platform so they can be worked on standalone. Each implementation is
self-contained and carries its own canonical docs — this repo root only orients and routes.

> **This repo is not PCO-governed.** It has no injected governance/policy layer. The working
> conventions a coding agent should follow are distilled in [AGENTS.md](AGENTS.md); the real
> per-task playbooks live inside each implementation's docs (below).

## Implementations

Both live under [`implementations/`](implementations/):

### [`implementations/elt_ingest_excel`](implementations/elt_ingest_excel)
Config-driven Python/DuckDB ELT pipeline: read source Excel extracts → load to DuckDB → ordered
SQL transforms → publish into cloned Workday load workbooks. Self-contained `uv` project
(`pyproject.toml` + `uv.lock`).

- **Start here:** [README](implementations/elt_ingest_excel/README.md) →
  [docs/ARCHITECTURE.md](implementations/elt_ingest_excel/docs/ARCHITECTURE.md),
  [docs/CLI.md](implementations/elt_ingest_excel/docs/CLI.md),
  [docs/pipeline.md](implementations/elt_ingest_excel/docs/pipeline.md)
- **New data-migration task:** the `docs/finance-*.md` files are the canonical per-dataset
  contracts (supplier, supplier-invoice, supplier-po-receipt, customer, customer-invoice).
  Read the relevant one first, then [docs/CONFIGURATION.md](implementations/elt_ingest_excel/docs/CONFIGURATION.md)
  and [docs/sql-transform.md](implementations/elt_ingest_excel/docs/sql-transform.md).

### [`implementations/elt_transform_excel`](implementations/elt_transform_excel)
Finance-user-run, **native-VBA-only** Excel workflow: a Runner workbook turns a source workbook
into a populated Workday EIB workbook. No Python/DuckDB in the shipped solution; the Python under
`scripts/` only builds the Runner from YAML transform profiles. Scripts run standalone via
`uv run --with pyyaml --no-project ...` (no repo-level environment needed).

- **Start here:** [README](implementations/elt_transform_excel/README.md) →
  [docs/ARCHITECTURE.md](implementations/elt_transform_excel/docs/ARCHITECTURE.md)
- **Extend the EIB activities / add a transform:** the two canonical playbooks are
  [docs/AGENT_COLD_START_PROMPT_CONTRACT.md](implementations/elt_transform_excel/docs/AGENT_COLD_START_PROMPT_CONTRACT.md)
  (what a cold-start agent must do, ambiguity handling, the deterministic build+reconcile pipeline)
  and [docs/ADDING_A_TRANSFORM.md](implementations/elt_transform_excel/docs/ADDING_A_TRANSFORM.md)
  (the end-to-end authoring playbook).

## Setup

The repo is a **`uv` workspace** (root `pyproject.toml`); both implementations are members and
share one lockfile at the root (`uv.lock`). `uv` is the only Python tooling used (no pip/poetry/conda).

```bash
# from repo root — sync the whole workspace
uv sync

# or sync a single member
uv sync --package elt-ingest-excel
uv sync --package elt-transform-excel
```

- **`elt_ingest_excel`** — the pipeline runners:
  ```bash
  uv run python implementations/elt_ingest_excel/examples/fin_supplier_creditor.py   # see its docs/CLI.md
  ```
- **`elt_transform_excel`** — Python only builds/inspects the native-VBA Runner (the shipped
  solution is VBA). Either use the workspace env, or the docs' standalone form:
  ```bash
  uv run --package elt-transform-excel python implementations/elt_transform_excel/scripts/manifest_summary.py
  # equivalently, no env needed:
  uv run --with openpyxl --with pyyaml --no-project python implementations/elt_transform_excel/scripts/manifest_summary.py
  ```

## Path caveat (post-extraction)

These implementations were extracted from a repo where each module sat at the repo root. Some
module docs still show historical absolute paths (e.g. `.../elt_lake/elt_transform_excel`, or
`cd elt_lake`). Here the module directory is `implementations/<module>` — treat those paths as
repo-relative to the module and adjust `cd` targets accordingly. Some docs also point to a former
`ai_context/pco/...` governance tree that does not exist in this repo; those pointers are historical.
