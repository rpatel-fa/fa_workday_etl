# Working Contract for Coding Agents

This repo has no injected governance layer. These are the few conventions the implementations
here were built under and depend on. Each rule states its consequence — a rule without one gets
overridden under pressure. The authoritative per-task playbooks are the module cold-start docs
(see [README.md](README.md)); read the relevant one before touching code.

## Before you start a task
- **Read the canonical doc first.** Ingest → the matching `docs/finance-*.md`; transform → `docs/AGENT_COLD_START_PROMPT_CONTRACT.md`. *Consequence:* these encode finance decisions already made; re-deriving them from data produces wrong mappings that look plausible.
- **Surface ambiguity, don't fill it.** If a source column, mapping, or target field is unspecified, ask — do not infer from domain terms. *Consequence:* silent assumption-filling is the leading source of rework and produces migration errors with no error signal.

## Config purity (both modules are config-driven)
- Domain strings — entity names, dataset identifiers, column names, thresholds, mappings, prompt/label text — live in YAML/JSON config, never hard-coded in Python or VBA. *Consequence:* domain context baked into code cannot be reused for the next dataset; you get a new near-duplicate implementation instead of a config change.
- Adding a dataset or transform = new/edited config, not new bespoke code, unless the mechanism genuinely doesn't exist yet.

## Surgical changes
- Change only what the task requires; match existing style; don't refactor adjacent code. *Consequence:* unrequested changes widen diffs and inject unreviewed risk into a migration where correctness is auditable against source.

## Verify, don't assert
- Prove results with live queries/output (e.g. before/after DuckDB row counts for ingest; the Runner's colour-coded result cells + reconcile for transform), not "the SQL/mapping looks right". *Consequence:* weak success criteria produce false "done" signals on data that is wrong.
- For transform, the build gate is non-negotiable: a code-complete transform is **pending manual Excel import + verify** until confirmed in real Excel.

## Docs currency
- Any code/config change updates the affected module docs in the same change (README, ARCHITECTURE, CONFIGURATION, the relevant `finance-*.md`/transform doc, TODO). *Consequence:* stale docs are a bug in the change — the next cold-start agent trusts them and acts on wrong facts.

## Tooling
- Python: `uv` only (`uv sync`, `uv run`). Ingest is a uv project; transform scripts run via `uv run --with <deps> --no-project`. No pip/poetry/conda.
