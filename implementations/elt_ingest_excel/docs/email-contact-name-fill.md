# Email → Contact Name fill-in (standalone tool)

**Purpose.** Populate `First name` + `Surname` columns on a 3-column (email, First name, Surname)
contact workbook *from the email local-part only*, for Gayle/3+ style EIB prep:
> "If it is not 100% clear what the First or Surname is → **leave both blank**.
> Ambiguous rows cannot be loaded into Workday via EIB and must be handled manually."

This is a **standalone utility** — it sits *next to* the ingest / transform pipelines,
not *inside* them. It reads a raw contact workbook, writes a filled copy, and does not
touch DuckDB, ingest config, or any publish workbook. When Finance/3+ say "here is a
list of emails, can you fill First/Surname for EIB?", this doc is the playbook.

**Why this lives in the repo.** The 2026-09 original spreadsheet ("Customer Contact
Emails and Names.xlsx", 4478 rows) was the first instance, but the request pattern is
recurring (supplier contact lists, customer contact lists, etc.). Repo-capturing the
script + review-playbook means the next cold-start agent runs the *same* conservative
pipeline rather than improvising a new parser that silently ships swapped names.

---

## Input / Output contract

| Item | Value |
|---|---|
| Source workbook | 3-column sheet: col A = email, col B = `First name`, col C = `Surname`. Any rows with blank col A are **preserved unchanged** (blank rows, label rows). Any row whose B **and** C were already non-empty in the source is **copied verbatim** (never overwritten). |
| Output workbook | Same shape, same row order, same sheet name (`Sheet1`). No new columns. |
| Sheet name | Must be `Sheet1` (the tool defaults to the active sheet and keeps its name). |
| Fill rate | Typical 35–45% of data rows filled; 55–65% left deliberately blank. The tool is tuned to leave a row blank *well before* it risks shipping a swapped or hallucinated name. |

---

## Pipeline

Four phases. **Phases 1 + 3 + 4 run offline from a single script.** Phase 2 is an
agent-review pass and is optional for small runs (see the `--skip-agent-review`
equivalent: omit `--emit-review-bundle` and treat Phase-1 output as final).

```
source workbook
   │
   ▼
[1] PARSE          email_to_contact_names.py  →  filled_candidates  (~40%)
                                                + blank_candidates   (~60%)
   │
   ▼
[2] AGENT REVIEW   optional. 2 parallel sub-agents each review ½ the candidates
                    and emit ok:true / ok:false+reason JSON per row.
   │
   ▼
[3] REPAIR / BLANK deterministic repair rules apply only to flagged rows; any row
                    still failing the zero-defect gate → blanked both sides.
   │
   ▼
[4] QUALITY GATE   structural checks across every filled row; 0 failures required
                    before "send to Gayle/3+".
   │
   ▼
 PARSED.xlsx  —  send
```

### Phase 1 — Parse (`email_to_contact_names.py`)

Run from repo root (the script is standalone; no workspace env required):

```bash
# Minimal — no agent review (fine for 100-row runs, or the "quick eyeball" pass)
uv run --with openpyxl --no-project \
    python implementations/elt_ingest_excel/scripts/email_to_contact_names.py \
    --src  "~/Downloads/Customer Contact Emails and Names.xlsx" \
    --out  "~/Downloads/Customer Contact Emails and Names - PARSED.xlsx"

# Same, but also emit a review bundle so we can run Phase 2 after
uv run --with openpyxl --no-project \
    python implementations/elt_ingest_excel/scripts/email_to_contact_names.py \
    --src  "~/Downloads/Customer Contact Emails and Names.xlsx" \
    --out  "~/Downloads/Customer Contact Emails and Names - PARSED.xlsx" \
    --emit-review-bundle /tmp/email_name_review
```

The script internally applies, in priority order:

1. **Strip leading/trailing digits** (`rolandtaylor25` → `rolandtaylor`, `1paul` → `paul`).
2. **Strip honorific prefixes** (`mr`/`mrs`/`ms`/`miss`/`dr`/`prof`/`rev`/`mx`/`mstr`) when the remaining tail is ≥ 4 chars.
3. **Blocklist instant-blank rule.** Any separator-token from [BLOCKLIST_TOKENS](file:///Users/Rakesh.Patel/Documents/__code/git/rakeshpatelfa/fa_workday_etl/implementations/elt_ingest_excel/scripts/email_to_contact_names.py#L79-L100) (accounts, payable, purchase, invoicing, secretary, info, club, football, etc.) on either side → row blanked immediately.
4. **Separator 2-token split** (`first.last`, `first_last`, `first-last`): try forward (prefix=first) and reverse. If *both* fit → tie-break to forward per UK convention; if *neither* fits → blank.
5. **Separator 3-token split:** allowed ONLY for the `first.[1-char-middle].last` / `first_[1char]_last` pattern (e.g. `jack.7.bray`). Any other 3-token pattern → blank.
6. **Separator 4+ token:** blanked. Three names / initials-middle / role aliases all happen here; there is no safe inference.
7. **Single-token compound (no separators):** bidirectional sweep at every split point →
   - H1: `prefix=first`, `suffix=last`
   - H2: `prefix=last`,  `suffix=first` (surname-first order, some EU/old orgs)
   - each candidate scored on first-name dictionary match, surname-shape gating (≥4 chars + vowel, not blocklist, no club suffix), onset-initial frequency weight (S/B/C/W/H/M/… common; X/Q/Z rare), typical first-name length bonus (3–7 chars).
   - top-2 candidates of opposite direction within ≤0.05 score → **blanked** (ambig).
   - top-2 opposite-direction within 0.05–0.15 score → **H1 (prefix-first) tie-break** per UK convention.
8. **Zero-defect filter applied immediately before writing.** If a candidate fails structural checks (punct in name, digits, 2-char first not in safe-list, 3-char surname not in safe-list, no vowels, first==last, etc.), it is blanked with a `zdrop:` reason rather than emitted.

### Phase 2 — optional 2-agent parallel review

When `--emit-review-bundle /tmp/email_name_review` is passed, Phase 1 writes two files:

- `/tmp/email_name_review/chunk_a.json` — filled rows 0, 2, 4, … (≈½)
- `/tmp/email_name_review/chunk_b.json` — filled rows 1, 3, 5, … (≈½)

Each file is a JSON array of `{email, first, last, parse_reason}`. To review:

1. Launch two parallel `general_purpose_task` sub-agents, one per chunk.
2. Each sub-agent emits one JSON object **per input row** of the form:
   ```json
   {"email":"paulrhocking@icloud.com","ok":true}
   {"email":"salesjohnmcooper@hotmail.com","ok":false,"reason":"Role word used"}
   ```
3. Use reason tags drawn from:
   `Bad surname parse`, `Role word used`, `Generic first name`, `Period in name`,
   `Digits in name`, `First/last swapped`, `Surname too short`,
   `First name too long`, `Bad first parse`, `title merged`, `misc`.
4. Conservative rule: if in any doubt, flag. A false-flag is repaired by Phase 3's
   deterministic rules; a false-OK passes a hallucinated name to finance.

Save the agent outputs next to each chunk: `/tmp/email_name_review/chunk_a_flags.json`,
`/tmp/email_name_review/chunk_b_flags.json`.

### Phase 3 — Repair / Blank (for agent-flagged rows)

A deterministic rule set applies **only to rows flagged by either agent.** Unflagged
rows pass through untouched. The current rules (captured inside the script's
`extract_names()` + `zero_defect_filter()` when run as part of the same combined
parse-blank pass; re-run same script to re-apply after manual flag JSON assembly):

| Rule | Trigger | Action |
|---|---|---|
| R1 | Any separator token in blocklist | BLANK |
| R2 | `mr`/`dr`/`mrs`/`miss` etc. title merged into start of local part | Strip title, then re-run the 2-token / single-token split (recovers `mrdannyburke`, `mrdavestyles`). |
| R3 | ≥ 3 separator tokens, middle is NOT exactly 1 char | BLANK |
| R4 | Exactly 2 separator tokens, one is 1 char (initials) | Drop the middle initial; re-evaluate. |
| R5 | 2-token separated where one side is 1 single letter (initials) | BLANK (e.g. `a.shoesmith` — we can't guess first name from `A`). |
| R6 | Single-token compound with swapped-order tie-break score between (0.05, 0.15] | Keep **H1** (prefix = first) — UK `first.last` convention wins. |
| R7 | Any of the structural `zero_defect_filter` checks still fail after repair | BLANK |

### Phase 4 — Quality gate (mandatory before sending)

The gate is **non-negotiable.** Every one of these must read `0 / filled_count` before
handing off:

| Check | Meaning |
|---|---|
| blocklist_anywhere | No role/generic token on first or last. |
| punct_first_dot_under / punct_last_dot_under | No `.` or `_` inside either name (hyphens in surnames are fine). |
| punct_edge | No leading/trailing `-` / `_`. |
| digits_first / digits_last | No digits in either name (legacy year suffixes are stripped from the *email*; if they bleed into a name, it's a bad split). |
| case_first / case_last | First letter upper, rest lower. Hyphenated parts get own cap. |
| first_too_short | ≥ 2 chars; 2-char only if in the `REAL_2_CHAR_FIRST_NAMES` safe-list (Al/Ed/Jo/…), **never** if it's 2 random caps (Kc, Sk, Nj, …). |
| last_too_short_or_not_in_safe3 | 2-char last → REAL_2_CHAR_SURNAMES only (Ng/Oh/Li/Au/…); 3-char last → REAL_3_CHAR_SURNAMES only (Cox/Day/Lee/Ray/…). Everything else ≤ 3 chars = blank. |
| first_no_vowel_len3plus | Real given names have vowels. `Gkr` etc. = not a name. |
| first_eq_last | No one has `Mick Mick`; it's a split failure on a doubled token. |
| club_suffix | Surname ends `fc / afc / utd / town / rovers / rangers / …` → it's a club address. Blank. |

The script applies these gates internally *before writing*, so the output xlsx on disk
already passes them. The 2026-09 original 4478-row run: 1899 filled, 2579 blank, zero
structural defects. `first_not_in_dict` is **not** a fail-reason — it's an info flag.
186 / 1899 filled rows fall into this on the baby-names corpus because it only covers
US English names; international names (Sanjay, Ajay, Raihan, Keturah, Alicja, Bunyamin,
Dai, Debs, Rheanna, …) all correctly survive the gate.

---

## Anatomy of a real run (2026-09, 4478 rows)

- Source: `~/Downloads/Customer Contact Emails and Names.xlsx` — 1 pre-filled row (`lee.ashmore1`), preserved.
- Phase-1 filled: 2055 rows.
- 2-agent review flags: 113 rows (Chunk A) + 47 rows (Chunk B) = 160 unique emails flagged.
- Phase-3 repair: 40 flagged rows successfully re-split with the strict rules above; 1 flagged row still ambiguous (`purchasing@hospitalityfinder…`) → blanked.
- Phase-4 second sweep (post-parse zero-defect filter): another 36 borderline rows blanked — mostly 3-char junky surnames (Mcc, Pkr, Cvp, Jcc, Hj, Pkr, Iqi, Sk, Ca), 2-char junk initials (Kc, Sk, Nj, Ph, Tg), trailing underscores in the split, and bleed-through of digits into surnames.
- Final: **1899 filled, 2579 blank.** No structural defects.

### Spot-check legend

On every run, at minimum confirm the following behaviour classes by filtering the
output for these patterns:

| Pattern | Expected | Why |
|---|---|---|
| `paulrhocking@icloud.com` | Paul / Rhocking | Single-token dict split, H1. |
| `paulashton@…` | Paul / Ashton | Both orders plausible → H1 tie-break. |
| `adrianreynolds@…` | Adrian / Reynolds | Single-token H1. |
| `jamieedmunds@…` | Jamie / Edmunds | Single-token H1. |
| `mrdannyburke@…` / `mrdavestyles@…` | Danny / Burke, Dave / Styles | Honorific strip then 2-token. |
| `chrisdharris@…` | Chris / Dharris | Middle-initial H1 (common UK). |
| `sianywilliams4@…` | Siany / Williams (last-first concat) → H2. | Last-first order OK if only order that fits dict+shape. |
| `abi.invoicing@…` / `purchase.ledger@…` / `accounts.payables@…` / `bpo-cbl.invoicereceipt@…` | BLANK | Blocklist R1. |
| `a.shoesmith@…` | BLANK | 1-char first, R5. |
| `alan.denis.smith1947@…` / `charles.macinnes.25@…` | BLANK | 4-token / year bleed → R3 R7. |
| `rolandtaylor25@…` / `wright.ian24@…` / `docker.andrew@…` | BLANK | Digit-suffix / reversed-dot ambiguity → no fit. |
| `martin@1stbusstop.co.uk` (single-token, only a first) | BLANK | Cannot invent a surname. |
| `jack.7.bray@gmail.com` (middle-init sep-3) | Jack / Bray | R4. |

If any of the "BLANK" rows above come out filled on a new run, the parser has
regressed — stop and adjust the rules before handing the file to finance.

---

## Files created for this tool (the source of truth)

| Path | Role |
|---|---|
| [scripts/email_to_contact_names.py](file:///Users/Rakesh.Patel/Documents/__code/git/rakeshpatelfa/fa_workday_etl/implementations/elt_ingest_excel/scripts/email_to_contact_names.py) | The whole 4-phase script. Run this; doc the invocation in the PR. |
| [docs/email-contact-name-fill.md](file:///Users/Rakesh.Patel/Documents/__code/git/rakeshpatelfa/fa_workday_etl/implementations/elt_ingest_excel/docs/email-contact-name-fill.md) | This file. Cold-start agent should read this before touching anything. |
| `~/.cache/contact_first_names.json` | Local cache of the 6782-name baby-names corpus after first network hit; embedded fallback (~320 names) ships inside the script if the network is unavailable. |

## Extending this

The rules above are intentionally conservative; we blank rather than guess. If Finance
starts seeing a repeated pattern (e.g. `initials.surname` where initials decode to
first+middle from a trusted internal Workday CCS reference), add it as a new
**high-specificity rule, not as a relaxed split.** Example of a bad change: "allow
2-char surnames not in safe-list". Example of a good change: "when the 2-char surname
token exactly matches `Ng / Oh / Li / Au …`, keep, else keep blank". Keep the 2026-09
spot-check list above as a regression list — re-run on the original workbook after
every rule change and confirm every row still behaves exactly as in the legend.
