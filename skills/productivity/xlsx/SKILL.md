---
name: xlsx
description: Create, read, edit Excel .xlsx workbooks and CSVs.
version: 1.3.0
author: Nous Research
license: MIT
platforms: [linux, macos, windows]
metadata:
  hermes:
    tags: [excel, spreadsheet, xlsx, csv, openpyxl, productivity]
    category: productivity
    related_skills: [docx, pdf, powerpoint]
---

# Xlsx Skill

Work with Excel .xlsx workbooks using Python and openpyxl: build styled
multi-sheet workbooks with formulas and charts, inspect or dump existing
files, edit cells and structure, and convert to/from CSV. All helper
scripts are argparse CLIs that print JSON and use explicit UTF-8 I/O.
Writes run through a guarded chain (backup -> temp -> verification ->
atomic replace) and every run prints a structured JSON envelope.
A read-only Document Map builder (`xlsx_understand.py`, Phase 2A) adds
structural understanding: sheets, regions, headers, columns, tables,
merges, hidden state, formulas and repeated structures - it can never
write to the workbook it reads.

## When to Use

- Creating .xlsx reports: multiple sheets, number formats, styling,
  merged cells, freeze panes, autofilter, conditional formatting,
  charts, data-validation dropdowns, native Excel tables, defined
  names, hyperlinks, cell notes, sheet protection.
- Reading a workbook: sheet inventory, dumping data as JSON or CSV,
  listing formulas vs cached values, notes, defined names, tables.
- Editing existing files: set cells, append rows, insert/delete
  rows/columns (reference-aware via `xlsx_restructure.py`),
  copy/rename sheets, tables, names, notes, protection.
- Recalculating formulas headlessly via LibreOffice
  (`xlsx_recalc.py`).
- Understanding an unfamiliar workbook structurally: a deterministic,
  read-only **Document Map** with `xlsx_understand.py` (regions, headers
  with confidence + evidence, per-column profiles, tables, merges,
  hidden rows/columns, formula inventory with cached-value availability,
  repeated structures). Phase 2A: detection only - no template matching,
  no filling, no writing.
- CSV interop with type inference and non-UTF-8 encodings.
- Safely writing into existing workbooks: every save takes a backup,
  verifies the temp output by re-opening it, replaces atomically, and
  records the operation (see "Write Safety & Operation Record").
- Not for the legacy .xls binary format (use LibreOffice to convert
  first: `soffice --headless --convert-to xlsx old.xls`).

## Prerequisites

- Python 3.10+ with `openpyxl` pinned by this skill in
  `scripts/requirements.txt`: `openpyxl>=3.1.5,<3.2`. The `<3.2` pin is
  deliberate - the fidelity matrix, the chart/image round-trip
  guarantee and the D4/D5 corrections were measured against 3.1.5.
- Recommended runner (uses the pin without touching global packages):
  `uv run --with-requirements scripts/requirements.txt python scripts/...`
- If openpyxl is missing, every script exits 1 with a structured
  `DEPENDENCY_MISSING` JSON error whose `recovery` field shows the real
  install command for this environment. Scripts never `pip install`
  anything by themselves.
- Optional: LibreOffice (`soffice`) for headless recalculation or
  format conversion.

## How to Run

Run the helper scripts with the `terminal` tool from this skill's
`scripts/` directory (every script supports `--help`). Every script
prints one JSON object on stdout; failures print the same JSON payload
on stderr with `"ok": false` and exit 1:

```bash
# recommended runner (keeps the pinned openpyxl)
uv run --with-requirements scripts/requirements.txt python scripts/xlsx_read.py report.xlsx --sheets

python scripts/xlsx_create.py spec.json report.xlsx   # build from JSON spec
python scripts/xlsx_read.py report.xlsx --sheets      # inventory
python scripts/xlsx_read.py report.xlsx --json --sheet Data
python scripts/xlsx_read.py report.xlsx --formulas
python scripts/xlsx_validate.py report.xlsx --expect-sheets Data,Ozet
python scripts/xlsx_edit.py report.xlsx --sheet Data --set B2=42 --recalc
python scripts/xlsx_restructure.py report.xlsx --sheet Data --insert-rows 3:2
python scripts/xlsx_recalc.py report.xlsx
python scripts/csv_to_xlsx.py data.csv out.xlsx --encoding utf-8
python scripts/xlsx_to_csv.py report.xlsx out.csv --sheet Data
```

Author the JSON spec with `write_file`, inspect script JSON output with
`read_file` or directly from stdout.

## Quick Reference

| Task | Command |
|---|---|
| Create workbook from spec | `xlsx_create.py spec.json out.xlsx` |
| Sheet names + dimensions | `xlsx_read.py f.xlsx --sheets` |
| Dump sheet as JSON | `xlsx_read.py f.xlsx --json --sheet S` |
| Dump sheet as CSV | `xlsx_read.py f.xlsx --csv --out d.csv` |
| List formulas + cached values | `xlsx_read.py f.xlsx --formulas` |
| Set a cell / formula | `xlsx_edit.py f.xlsx --set "A1==SUM(B:B)"` |
| Append a row | `xlsx_edit.py f.xlsx --append '[1,"x",true]'` |
| Insert 2 rows, refs NOT shifted | `xlsx_edit.py f.xlsx --insert-rows 3:2` |
| Insert 2 rows, refs shifted | `xlsx_restructure.py f.xlsx --insert-rows 3:2` |
| Delete a column, refs shifted | `xlsx_restructure.py f.xlsx --delete-cols B` |
| Create a native table | `xlsx_edit.py f.xlsx --add-table Sales:A1:C9` |
| Append inside a table | `--table-append 'Sales=["West",5]'` |
| List tables | `xlsx_edit.py f.xlsx --list-tables` |
| Defined names | `--define-name "Rates='Data'!$B$2:$B$9"` / `--delete-name Rates` / `xlsx_read.py f.xlsx --names` |
| Hyperlink | `--hyperlink "A1=https://example.com|Docs"` |
| Cell note | `--note "B2=Check this|Reviewer"`; read via `xlsx_read.py f.xlsx --notes` |
| Protect sheet (see Pitfalls) | `--protect your-password --unlock B2:B9` |
| Recalculate via LibreOffice | `xlsx_recalc.py f.xlsx` |
| Document Map (read-only) | `xlsx_understand.py f.xlsx`; `--pretty`, `--samples N`, `--formula-list-cap N` |
| Semantic roles + blocks | `xlsx_semantics.py f.xlsx`; `--pretty`, `--doc-map map.json` |
| Template profile (deterministic) | `xlsx_semantics.py f.xlsx --emit profile --out f.profile.json` |
| Match a new form to a template | `xlsx_semantics.py new.xlsx --match f.profile.json`; `--review-file review.json` |
| Source→target plan (writes nothing) | `xlsx_mapping.py --profile f.profile.json --source data.csv --emit plan --out plan.json` |
| Dry-run fill plan (writes nothing) | `xlsx_mapping.py --profile f.profile.json --plan plan.json --source data.csv --target f.xlsx --emit dry-run` |
| Copy / rename sheet | `--copy-sheet Src:New --rename-sheet Old:New` |
| Force recalc on open | `xlsx_edit.py f.xlsx --recalc` |
| Validate a written file | `xlsx_validate.py f.xlsx --expect-sheets Data,Ozet` |
| CSV -> styled xlsx | `csv_to_xlsx.py in.csv out.xlsx` |
| xlsx -> CSV | `xlsx_to_csv.py f.xlsx out.csv --encoding utf-8` |

## Structured Output

Every script prints one JSON object on stdout and exits 0 on success,
1 on failure (errors repeat the same object on stderr with
`"ok": false`). The success envelope always carries:

- `warnings[]` - degradations that did not stop the operation, each
  with `code` + `message` (e.g. `UNKNOWN_SPEC_KEY`, `FEATURE_PARTIAL`);
- `unsupported[]` - features that were requested or encountered but
  could not be handled (never silently skipped);
- `diagnostics[]` - non-fatal findings (e.g. unreadable manifest lines).

On failure the payload carries `error_code`, `error` (alias `message`,
both hold the human text), `context` (structured details such as the
offending spec path or requested sheet) and `recovery` (what to do
next; for `DEPENDENCY_MISSING` it is the real install command).

Spec validation is strict by default: unknown spec keys fail with
`SPEC_UNKNOWN_KEY` and a closest-match hint; pass
`--allow-unknown-keys` to downgrade them to a warning.

## Document Map (Phase 2A)

`xlsx_understand.py f.xlsx [--pretty] [--samples N]` prints the
EXCEL DOCUMENT MAP: a deterministic, JSON-serializable, read-only
structural description. Exit codes and the error envelope match the
other scripts; a successful run prints the map itself with `warnings`,
`unsupported` and `diagnostics` arrays.

Top level: `{ok, workbook, sheets[], locale, warnings, unsupported,
diagnostics}`.

Read-only architecture (three passes, no write path exists):

- **Pass A** - `read_only=True, data_only=False`: cell values, formula
  text, per-row profiles (type signature, non-empty counts, style
  summary) and per-column statistics. Gap rows are reconstructed from
  real coordinates, so read-only padding never invents cells.
- **Pass B** - `read_only=True, data_only=True`: cached values for the
  formula cells found in Pass A only. Availability is reported as
  `cached_value_available`; the map always states
  `values_recalculated: false` - cached values are prior-save
  artifacts and openpyxl never recomputes anything.
- **Pass S** - structural metadata read straight from the raw OOXML
  parts (zip + XML): tables, merges, hidden rows/columns, validations,
  conditional formats, freeze panes, autofilter, defined names,
  external links, calculation settings, chart/image counts. No workbook
  object is created on this path at all; the no-write guarantee is
  covered by tests that SHA-256 the input before and after every run,
  and by an oracle test that proves the raw reads agree field-for-field
  with openpyxl's own normal-mode access (see Testing).

Per sheet, the map carries: `regions[]` (`title`, `header`,
`subheader`, `data`, `input_region` (styled-but-empty), `section_header`,
`subtotal`, `total`, `notes`, `footer`, `formula_region`), `headers[]`
(single/multi-row/merged, labels, blank cells, `confidence_score` and
`confidence_level`, evidence strings, repeat links), `columns[]`
(`data_type` over data rows plus `data_type_full_column` over the whole
column, counts, samples, number format, style summary, header and
normalized header, region), `tables[]` (name, ref, columns, row_count,
totals), `merged_cells[]`, `hidden` (rows/columns, outline levels kept
separate from hidden flags), row heights / column widths,
`freeze_panes`, `autofilter`, `validations[]`, `conditional_formats[]`
(unknown rule kinds go to `unsupported`, never silently skipped),
`formulas[]` (cell, text, coarse `kind`, functions, operators, parsed
refs with absolute flags and sheet refs, cached availability) and
`repeated_structures[]` (repeated headers, repeated row blocks,
periodic patterns - DETECTION confidence only, never template
matching).

Determinism: canonical JSON (`sort_keys`, compact separators), UTF-8
with `ensure_ascii=False`, no timestamps, sorted collections - two runs
on the same file are byte-identical. Uncertain detections are reported
with LOW confidence plus `evidence` and a `LOW_CONFIDENCE_DETECTION`
warning; declared sheet dimensions and the computed used range stay
separate fields with an inflation diagnostic. Nothing is silently
dropped or auto-fixed.

Scale honesty: `formulas[]` lists at most `--formula-list-cap`
entries per sheet (default 5000); the per-sheet `formula_summary`
aggregates (kinds, top functions, cross-sheet/absolute counts, cached
availability) always cover ALL formulas, and a
`FORMULA_LIST_TRUNCATED` warning is emitted when the list is capped.
Per-row label values are kept for the first 500000 non-empty cells per
sheet (`ROW_VALUES_PARTIAL` warning beyond); row heights and column
widths are capped at 1000 entries with `STRUCTURE_CAP_REACHED`. Row
values feed label/region heuristics only - formulas, tables and
structure are never approximated by these budgets. Measured locally: a 3.6 MB real workbook maps in ~1 s; a 5.3 MB workbook with
256k formulas maps in ~52 s (dominated by the two read-only scans).

Phase 2A deliberately does NOT do: template matching, source->target
mapping, confidence-based filling, row expansion, formula propagation,
R1C1 semantics or any workbook modification. Those are later phases.

Phase 2B (semantic roles, template profiles, matching, mapping and
dry-run plans) is documented below, together with **Phase 3A** — the
template-execution layer that applies an approved plan. Row expansion, formula propagation and lookup execution are handled by the hardened execution layer (Phase 3C).

## Semantic Layer, Profiles & Mapping (Phase 2B)

Phase 2B builds on the Document Map: it labels what the cells *mean*,
turns a form into a reusable template identity, matches a new workbook
against that identity, and produces a mapping plan plus a dry-run fill
plan. **It never writes.** Execution is Phase 3.

| Command | Purpose |
|---|---|
| `xlsx_semantics.py BOOK.xlsx` | B1+B2: semantic roles, blocks, slots |
| `xlsx_semantics.py BOOK.xlsx --emit profile --out b.profile.json` | B3: deterministic template profile |
| `xlsx_semantics.py NEW.xlsx --match b.profile.json` | B4: match report (+ `--review-file review.json`) |
| `xlsx_mapping.py --profile b.profile.json --source data.csv --emit plan --out plan.json` | B5: source→target plan |
| `xlsx_mapping.py --profile b.profile.json --plan plan.json --source data.csv --target BOOK.xlsx --emit dry-run` | B6: what *would* be written |

`--doc-map map.json` reuses a saved 2A map instead of re-running 2A on a
workbook (identical output — measured by a test).

### Semantic model (B1/B2)

Per column: `semantic_role` (identifier, label, date, period, quantity,
currency, percent, code, address, person_name, free_text, summary, flag,
formula_derived, empty_slot, unknown), `value_kind` (string/integer/
decimal/date/datetime/boolean/formula/blank/mixed), `unit_guess` (only
with evidence, else null), `role_score`, `role_level`, `candidates[]`,
`evidence[]`.

Evidence comes from three families and every decision carries at least
one item (`header_term:…`, `value_affinity:…`, `region:…`,
`number_format:…`, `validation:…`, `table_column:…`):

1. header text vs `references/semantic-dictionary.tr-en.json` (`tr-en.1`);
2. data type / number format / sample values from the Document Map;
3. structure and context (region type, position, repeated block,
   validation, neighbouring columns).

Confidence reuses the 2A thresholds (HIGH ≥ 80, MEDIUM ≥ 65, LOW < 65).
**Ambiguity margin:** `margin = best − second`; when `margin < 10` the
decision is listed in `ambiguity[]` and `requires_review = true`. Both
directions are ranked: a template column with two look-alike instance
columns *and* an instance column contested by two template columns.
LOW and ambiguous decisions always emit a warning.

Blocks (B2) carry `archetype` (record_block, record_row, input_slot_row,
subtotal_row, total_row, note_row) and a `record_key_candidate` with
`null_rate`, `distinct_count`, `duplicate_count` and
`candidate_key_score` — an identifier role is **not** assumed to be a
unique key; duplicates are reported as evidence.

Slots carry `required` **and** `required_source` (`heuristic` here; a
profile round-trip preserves `user`). A heuristic "required" is never
presented as user-approved.

### Template profile (B3)

`profile_version`, `profile_id`, `profile_hash`, `content_identity`,
`structure_fingerprint`, `dictionary_version`, `provenance`, `sheets[]`
(regions, columns, blocks, slots, formula_summary, tables, visibility,
merged layout, hidden state).

Identity chain: `content_identity` = SHA-256 over the canonical semantic
projection, so it is independent of the file path and of `mtime`
(measured: same bytes saved with a new timestamp → byte-identical
profile); `profile_id` = SHA-256 of content identity + structure
fingerprint + profile version + dictionary version. No timestamps
anywhere in profile, plan or report payloads.

### Match report (B4)

Components `region_score`, `column_score`, `formula_score`,
`structural_score` → `overall_score` + `overall_level`, plus
`unmatched_template[]`, `extra_instance[]`, `conflicts[]` (value_kind and
number_format mismatches are conflicts, never conversions),
`sheets[].column_matches[]` with per-match evidence, and
`review_state`.

`--review-file review.json` records human decisions next to the machine
decision:

```json
{"decisions": {"column_competition|instance:D, matched template:D, rival template:C":
               {"state": "accepted", "note": "checked by hand"}}}
```

States: `accepted` / `rejected` / `pending`; the report's
`review_state.state` becomes `unreviewed`, `partially_reviewed`,
`reviewed` or `nothing_to_review`. Unknown subjects warn
(`UNKNOWN_REVIEW_SUBJECT`) and a review **never** changes scores,
levels or ambiguity.

### Mapping plan and dry-run (B5/B6)

Plan: `plan_id` (deterministic), `mappings[]` (`source_key`, `target`
anchor, `mode` = `scalar`/`row_block`, confidence, evidence,
preconditions), `unresolved[]`, `conflicts[]`, `policy`
(`on_low_confidence`, `on_conflict`, `missing_required`, `max_rows`),
`complete` (false while any `unresolved`/`conflicts` remain). `lookup`
validations are reported through `unsupported[]` — never executed.

Dry-run partitions every planned action into exactly one class:
`would_write[]` (per cell: sheet/cell/source/preview/type/`checks`
{validation, merged, format_compatible, role_compatible}/`write_policy`),
`blocked[]` (failed precondition), `preserved[]` (formula-derived target
cells — the formula is never touched), `unresolved[]` (no source data —
nothing is invented) and `unsupported[]`. Row shortage never adds rows:
`row_expansion.status = "planned_only"`, `actual_rows_added = 0`.

### No-write guarantee (2B)

Both modules are scanned by AST tests that forbid `save`, `ZipFile`,
`open(..., "w"/"a")` and a direct `openpyxl` import; 2A is invoked
in-process and forwards its warnings/unsupported/diagnostics. Real files
are SHA-256-fingerprinted before and after every run in the tests and in
the corpus sweep.

### Scale honesty

The Document Map carries no row-level data, so `available_rows` can only
come from `input_region` bounds; otherwise it is `null` plus
`ROW_SLOTS_UNKNOWN` — never a guess. Long lists use
`count_total / returned_count / truncated` with `*_TRUNCATED` warnings
(profiles 2000 blocks/slots/columns, matches 2000 entries); beyond the
column cap a `COLUMNS_TRUNCATED` warning still carries the totals, and
columns that are entirely blank/unused are counted (`blank_columns`)
rather than listed (`BLANK_COLUMNS_SUMMARISED`). Profile columns are
capped because a transposed workbook can carry tens of thousands of
real columns (measured: 16384 columns -> 8.2 MB uncapped vs 1.5 MB
capped, AC-B26). Matching caps the *instance* side the
same way (`INSTANCE_COLUMNS_TRUNCATED`), because the pair scan is
O(template x instance) - measured on a 16384-column sheet: 306 s
uncapped vs 47 s capped, same match result. Role accuracy
(AC-B03): the oracle runs against `tests/fixtures/labels/labels.json`,
authored by the agent on the user's 2026-09-29 delegation (provisional,
not an independent human oracle) - measured **58/61 = 95.1%** against
the >=90% assertion. The three misses are documented: English "Total"
lands on `formula_derived` (HIGH) instead of `summary`, and two
placeholder-header columns ("Fazla Kolon", "Bilinmeyen Alan") stay
LOW-flagged guesses instead of `unknown`. Replacing `labels.json` with
independent labels re-runs the metric unchanged.

## Template Execution (Phase 3A)

Phase 3A applies an **approved** dry-run fill plan to the template. It
is the first (and only) layer allowed to write, and only to the cells
the plan names.

| Command | Purpose |
|---|---|
| `xlsx_execute.py --fill-plan dry.json --profile p.json --source data.csv --target BOOK.xlsx --approve-token TOKEN` | execute (default output `BOOK_filled.xlsx`; `--out` overrides) |
| `... --emit preflight --approve-token TOKEN` | read-only gate report, byte-deterministic |
| `... --in-place --approve-token TOKEN` | overwrite the target itself (backup mandatory, fail-closed) |

- **Approval:** plans with `requires_approval` need the token = the
  first 12 characters of `plan_id`. Missing/invalid token ->
  `APPROVAL_REQUIRED` / `APPROVAL_INVALID`; approval state is never
  invented.
- **Scope:** only `would_write[]` cells, each re-checked live before
  writing (sheet exists, cell addressable, not a merged inner cell, the
  current cell holds no formula, the source reference still resolves,
  the live source value matches the plan preview, type/format kind
  compatible, live dropdown validation satisfied). Any failed
  precondition stops the whole run before a single write — never a
  partial fill. Codes: `STALE_PLAN`, `PLAN_INVALID`,
  `MERGED_CELL_WRITE_FORBIDDEN`, `FORMULA_MODIFICATION_FORBIDDEN`,
  `TYPE_MISMATCH`, `VALIDATION_MISMATCH`, `SOURCE_CHANGED`,
  `ROW_EXPANSION_REQUIRED`, `PLAN_CONFLICT`, `QA_FAILED`,
  `UNEXPECTED_WRITE`, `UNEXPECTED_MODIFICATION`.
- **Safe chain:** preflight -> staged temp write -> reopen -> deep QA ->
  manifest -> atomic replace. QA compares the staged file against the
  pre-write snapshot (values, formulas, styles, merges, tables, defined
  names, validations, conditional formatting, hidden state): formulas
  must survive, and only planned cells may differ. QA failure = no
  commit; the staged file stays next to the output for diagnosis. A
  failed post-commit verification restores the backup (in-place runs).
- **Idempotency:** a prior execution of the same plan is found in the
  manifest — `already_applied` (the region already holds the result;
  zero writes), `safe_to_reapply` (region in its pre-execution state:
  empty or byte-identical layout), otherwise `PLAN_CONFLICT`. Planned
  cells that already hold the planned value are reported
  (`CELL_ALREADY_EQUAL`), never silent.
- **Not in 3A:** row insertion/expansion (`row_expansion` stays
  `planned_only`, `actual_rows_added: 0`), formula propagation or a
  formula-pattern engine, R1C1 normalization, lookup execution,
  automatic remapping.
- Every committed run appends an operation entry (`plan_id`, approval
  state + source, file/region fingerprints before/after, QA and
  validation status, output, counts) to `.xlsx_ops/manifest.jsonl` and
  emits a structured payload with `status`, `write_count`,
  `changes_count`, `noop_count`, `original_unchanged` and the QA
  summary.

## Write Safety & Operation Record

Every write (create / edit / restructure / csv_to_xlsx) runs the same
chain: **backup -> temp file -> reopen & verify -> atomic replace**.
The target is replaced only after the temp copy re-opens cleanly with
the expected sheets; on a failure the target is untouched and the temp
file is kept next to it for diagnosis. In-place saves always place a
timestamped `.bak-...` backup beside the file first.

Hard guards (library level, `xlsx_common.py`):

- `data_only=True` loads return a values-only handle; accessing its
  `.wb` or attempting any save is refused (`DATA_ONLY_WRITE_FORBIDDEN`).
- `read_only=True` handles can never be saved (`READ_ONLY_FILE`).
- In-place writes require explicit `in_place=True` at the API; the CLI
  keeps its legacy in-place default but records
  `approval: "caller-explicit"`, takes the backup, and stays atomic.
- Raw `wb.save(...)` outside `xlsx_common.py` is forbidden (enforced
  by an AST test in the suite).

Fingerprints: `fingerprint_file` (whole file), `fingerprint_region`
(range content at read time) and `fingerprint_structure` (sheet names +
table refs + used ranges) - stored with every write for verification
and for the idempotency/conflict checks of later phases.

Operation record: each committed write appends one line to
`.xlsx_ops/manifest.jsonl` (next to this SKILL.md) with `operation_id`,
`plan_hash`, source/target/region fingerprints, backup path, status and
timestamp. The log is trimmed to the last 200 entries; unreadable lines
never break reads - they surface as diagnostics.

`xlsx_validate.py FILE [--expect-sheets A,B,C] [--expect-tables N]`
re-opens a file and reports independently: sheet names, formula-text
count, `formula_text_ok`, and whether cached values were actually
recalculated (`values_recalculated` - false unless a spreadsheet app or
LibreOffice ran; a missing `soffice` is reported honestly, not hidden).

## Execution Hardening (Phase 3C)

Phases 3A/3B answer "does it work when the situation is right". Phase 3C
answers "does it stay safe when the situation is wrong": interrupted runs,
re-runs, missing or corrupt evidence, locked targets, huge files. It adds
**no new Excel intelligence** — only hardening, failure semantics and
honest diagnostics.

### Idempotency states

Before any write the executor reads `.xlsx_ops/manifest.jsonl` and decides:

| state | meaning | action |
|---|---|---|
| `first_run` | evidence exists, this plan was never applied | commit |
| `already_applied` | identical plan + identical post-region state | commit nothing |
| `safe_to_reapply` | plan applied, but the region was reset to its pre-state | commit |
| `unknown` | fresh workspace (no manifest) | commit + `MANIFEST_MISSING` warning |
| `unknown` | history exists but the record is gone/corrupt | **fail closed** (`IDEMPOTENCY_UNVERIFIED`) |
| `conflict` | region is neither the pre- nor the post-state | **fail closed** (`PLAN_CONFLICT`) |

A re-run never duplicates rows: `already_applied` returns
`write_count: 0`. If evidence is missing while the target region is
already populated, the run is refused unless the plan opts in with
`policy.allow_unverified_region_overwrite: true` (D3).

### Failure semantics

- **No partial commit.** All operations succeed or nothing is written:
  work happens on a staged temp file, then `atomic_replace`
  (`ATOMIC_COMMIT_FAILED` if the replace cannot be proven).
- **No success before success.** The `committed` record is appended only
  after the atomic replace is verified; intermediates are recorded as
  `staged` and are never treated as evidence.
- **Failure records carry what happened**: `operation_id`, `plan_id`,
  `stage`, `error_code`, `original_fingerprint`, `temp`, `backup`,
  `committed: false`, `timestamp`. A run blocked at preflight writes no
  manifest record at all (nothing was attempted) — the payload carries
  `stage: "preflight"`, `commit_state: "NOT_STARTED"`, `write_count: 0`
  and the blocker list instead.
- **Fail closed** whenever the outcome is unknowable: unreadable
  manifest, unmatched source fingerprint, unverifiable staged output,
  ambiguous replace, unknown recovery state -> no commit.
- **Original file safety** is the top invariant: a failing execution
  leaves the original byte-identical (verified by SHA-256 in the suite
  and in `smoke_pipeline_3c_recovery.py`).

`commit_state` values: `NOT_STARTED`, `STAGED`, `COMMITTED`, `FAILED`.

### Manifest hardening

`manifest_read` distinguishes `missing` / `corrupt` (with
`corrupt_count` + a `MANIFEST_CORRUPT` diagnostic, never a silent
partial read). Appends are serialised with a file lock; a held lock
returns `MANIFEST_LOCKED` instead of interleaving a line. Trim to the
last 200 entries is reported (`MANIFEST_TRIM_REPORTED`).

### Recovery / rollback

Temp and backup files are never deleted by a failing run — they are the
recovery material (kept beside the target, path recorded in the
manifest). After any failure: check `commit_state`, restore from the
`.bak-...` file if the target was mid-replace, then re-run the same plan
— idempotency will refuse to duplicate what already landed.

### Fault injection (tests only)

`xlsx_common.FAULTS` maps a point name to the error code to raise
(`backup`, `temp_create`, `staged_save`, `reopen_validate`,
`pre_commit_validate`, `expansion`, `formula_propagation`, `lookup`, `qa`,
`manifest_append`, `atomic_replace` — 11 points). It is the
supported way to prove the failure paths without breaking real files;
production code never sets it.

### Evidence scripts

- `scripts/smoke_pipeline_3c_recovery.py [--only NAME] [--json OUT]` — real
  workbooks, on copies only: success + injected failure + rerun + in-place
  idempotency, printing the SHA of every file before/after.
- `scripts/bench_pipeline_3c.py [--json OUT]` — per-stage timings
  (semantics/plan/dry-run/execute) on real files of increasing size plus the
  marginal cost of row expansion.
- Fault injection in tests: set `xlsx_common.FAULTS["<point>"] = "<ERROR_CODE>"`
  inside a try/finally — never in production code.

### Performance expectations

Honest, measured locally:
execution itself is fast (small files < 2 s; row expansion adds cost
linear in inserted rows). The expensive parts are the
**understanding** stages, dominated by loading large workbooks
(`read_only=True` for inventory) and building profiles — tens of seconds
to minutes on 10+ MB files. `timings_s` in the execution payload reports
each stage so a slow run can be attributed instead of guessed at.

## Formula Intelligence (Phase 4)

`scripts/xlsx_formula.py` — formülleri string olarak değil **yapısal** olarak
anlar: tokenize → normalize → deterministik R1C1 imzası → family discovery →
window classification → intent inference → bounded dependency graph →
kontrollü synthesis (yalnızca kanıtlanmış durumlarda).

```bash
# read-only intelligence (deterministic JSON; file SHA'sı değişmez)
python scripts/xlsx_formula.py analyze --workbook kitap.xlsx --out analiz.json

# read-only synthesis planı (9 ön koşul kapısı; hepsi bloklanırsa exit 3)
python scripts/xlsx_formula.py plan --workbook kitap.xlsx --targets "SAYFA!H13" --out plan.json

# onaylı planın Faz 1/3C safe-write zinciriyle uygulanması
python scripts/xlsx_formula.py execute --plan plan.json --workbook kitap.xlsx --out cikti.xlsx
python scripts/xlsx_formula.py execute --plan plan.json --workbook kitap.xlsx --in-place --approve <plan_id>
```

Kurallar (testlerle sabit):

* **Kanıt önce**: tek formülden family/window/intent çıkmaz; belirsizlik
  `UNKNOWN_WINDOW` / `unknown` olarak **gerçek sonuç** döner, hiçbir şeye
  zorlanmaz (D4, §1.18).
* Her inference `evidence[]` + `ambiguity[]` + HIGH/MEDIUM/LOW confidence
  taşır; fonksiyon adı tek başına kanıt değildir (§1.17).
* **Synthesis yalnızca HIGH-confidence, destekli, deterministik ailelerde**:
  üretilen formül yeniden parse edilir ve kaynak ailenin imzasını **birebir**
  üretmek zorunda; aksi `FORMULA_SYNTHESIS_BLOCKED` (D3, §1.12).
* Yazma zinciri Faz 1/3C'tir: ham `wb.save` yok; başarısızlıkta orijinal dosya
  byte-identical, partial commit = 0 (AC-4.21).
* `--in-place` için `--approve <plan_id>` zorunlu (`APPROVAL_REQUIRED`).
* `values_recalculated` kalıcı false'dur; hesaplama ancak ayrı recalc adımıyla
  (D9).

Ayrıntılar: `tests/test_xlsx_formula_oracle.py` (20 senaryoluk oracle).

## Procedure

1. **Create**: write a JSON spec (schema documented in
   `xlsx_create.py --help` and its docstring). Each sheet supports
   `rows` (scalars or styled cell objects), sparse `cells` overrides,
   `column_widths`, `row_heights`, `merges`, `freeze_panes`,
   `autofilter`, `conditional_formats` (cell_is rules and color
   scales), `charts` (bar/line/pie from cell ranges),
   `validations` (list dropdowns), `tables` (native Excel tables with
   a style name), and `protection`. Workbook-level `defined_names`
   maps names to refs. Cell objects also take `hyperlink` and `note`.
   Typed values: JSON numbers/bools
   pass through; dates use `{"value": "2026-01-31", "type": "date"}`.
   Number formats are Excel format strings: currency `"$#,##0.00"`,
   percent `"0.0%"`, date `"yyyy-mm-dd"`.
2. **Formulas**: set with `"formula": "SUM(B2:B9)"` in the spec or
   `--set "C1==SUM(A:A)"` in the editor. When writing formulas, add
   `"full_calc_on_load": true` (spec) or `--recalc` (editor); this sets
   the workbook's `fullCalcOnLoad` flag so Excel/LibreOffice recompute
   everything on open. openpyxl itself NEVER evaluates formulas.
3. **Read**: `--sheets` for inventory (names, dimensions, merged
   ranges, chart count, tables, protection, defined names),
   `--json`/`--csv` for data, `--formulas` to
   pair each formula string with its cached result, `--notes` for
   cell comments, `--names` for defined names. Cached results
   exist only if the file was last saved by a real spreadsheet app;
   files fresh from openpyxl return `null` there. To materialize
   results headlessly run `xlsx_recalc.py file.xlsx` (uses
   LibreOffice; prints `{"recalculated": false, ...}` and exits 0
   when `soffice` is absent), then reload with `--data-only`.
   `xlsx_validate.py FILE` gives an independent open-and-check report
   (sheet list, formula text count, honest recalc status).
4. **Edit**: `xlsx_edit.py` applies renames/copies first, then
   structural row/column changes, then `--set`/`--append`. It edits in
   place unless `--out` is given; every write follows the safe chain
   above (backup -> temp -> verify -> atomic replace), so the original
   survives any failure — but use `--out` when you want to leave the
   source file untouched.
5. **Restructure**: for insert/delete on sheets that have formulas,
   merges, tables, or filters, use `xlsx_restructure.py` instead of
   `xlsx_edit.py`. It rewrites formula references on ALL sheets
   (absolute `$` refs, ranges, cross-sheet refs), shifts merges,
   autofilter, freeze panes, validation and conditional-format
   ranges, table refs, defined names, and row/column dimensions, then
   prints a JSON report including a `not_shifted` list. Rules and
   limits: `references/restructuring.md`.
6. **CSV interop**: `csv_to_xlsx.py` infers int/float/bool/ISO-date
   per cell and styles the header row; `xlsx_to_csv.py` writes ISO
   dates and blank strings for empty cells. Both default to UTF-8 and
   accept `--encoding` (e.g. `utf-8-sig` for Excel-friendly BOM,
   `cp1252` for legacy Windows exports).

## Converting to PDF

LibreOffice converts headlessly (also works for CSV export of a single
sheet):

```bash
soffice --headless --convert-to pdf report.xlsx --outdir out/
soffice --headless --convert-to csv report.xlsx --outdir out/  # 1st sheet only
```

Only the first sheet lands in a CSV; for other sheets use
`xlsx_to_csv.py --sheet NAME`. If `soffice` is missing, install
LibreOffice or hand the file to the user unconverted.

## Pitfalls

- **openpyxl does not calculate.** Formula results are available only
  via `load_workbook(path, data_only=True)` and only when the file was
  previously saved by Excel/LibreOffice. Otherwise you get `None`.
- **`xlsx_edit.py` insert/delete does not shift references** (raw
  openpyxl behavior). Use `xlsx_restructure.py`, which does — but even
  it cannot move chart anchors, images, or conditional-format RULE
  formulas; read its JSON report's `not_shifted` list and
  `references/restructuring.md`.
- **Sheet protection is NOT security.** `--protect` sets the standard
  xlsx sheet-protection hash: it signals "don't edit this" to
  well-behaved apps and nothing more. Anyone can strip it by editing
  the zip's XML or unchecking it in LibreOffice. Never rely on it for
  confidentiality or integrity; it does not encrypt anything.
- **`data_only=True` then save** silently discards all formulas
  (cached values replace them). The scripts guard this: values-only
  handles refuse to save (`DATA_ONLY_WRITE_FORBIDDEN`). Never save a
  raw-openpyxl data_only workbook unless replacing formulas with their
  cached values is exactly the goal.
- **Foreign workbooks may hold exotic drawings.** openpyxl 3.1 parses
  and re-writes the charts/images it understands (the fidelity test
  verifies a charted fixture round-trips cleanly), but advanced Excel
  drawings — text boxes, shapes, arrows, SmartArt, sparklines — cannot
  be generated by openpyxl and their re-save behaviour is NOT_TESTED.
  Prefer reading, or edit on a copy, when a workbook uses them.
- **CSV locale traps**: always pass explicit encodings (the scripts
  already do) and remember European CSVs often use `;` delimiters and
  decimal commas — use `--delimiter ';'` and expect strings like
  `"12,5"` to stay strings.
- **Dates are datetimes**: Excel stores dates as serial numbers;
  openpyxl returns `datetime`/`date` objects. Dumps here emit ISO
  strings.
- Sheet names are capped at 31 chars and reject `[ ] : * ? / \`.

## Verification

- After creating: `xlsx_read.py out.xlsx --sheets` and confirm sheet
  names, dimensions, merged ranges, and chart counts match intent.
- Dump data with `--json` and compare against the source values.
- After edits: re-dump the touched range; if formulas were written,
  confirm `--formulas` lists them and that `--recalc` was applied.
- Every write result carries `verified: true`, a `plan_hash` and a
  `backup` path — confirm the backup file exists, and check the tail of
  `.xlsx_ops/manifest.jsonl` for the committed entry.
- `xlsx_validate.py out.xlsx --expect-sheets ...` re-opens the output
  independently before declaring success.
- After `xlsx_restructure.py`: read its JSON report, then re-run
  `--formulas` and `--sheets` to confirm references and ranges landed
  where expected.
- For a full visual check, open in LibreOffice:
  `soffice --headless --convert-to pdf out.xlsx` and inspect the PDF.

## Testing

```bash
# full suite with the pinned runner; --with pillow enables the image
# fidelity check (otherwise that one check is skipped)
uv run --with-requirements scripts/requirements.txt --with pytest --with pillow python -m pytest tests/ -q
```

- `tests/test_xlsx_skill.py` — CLI regression suite.
- `tests/test_xlsx_hardening.py` — safety guards, structured errors,
  D4/D5 fixes, fingerprints, manifest round-trip.
- `tests/test_xlsx_fidelity.py` — 18-feature round-trip through the
  safe chain plus OOXML part checks.
- `tests/test_xlsx_semantics.py` — Phase 2B semantics: roles,
  evidence, blocks/slots, profile identity (path- and mtime-
  independent), dictionary validity.
- `tests/test_xlsx_mapping.py` — Phase 2B mapping: plan dry-run
  partitions, `planned_only` row expansion, precondition cross-checks.
- `tests/test_xlsx_execute.py` — Phase 3A execution: plan validation,
  approval enforcement, scalar/row-block writes, type and dropdown
  safety, merged-cell and formula protection, QA-gate (no commit on
  failure), rollback, idempotency triad, manifest fields, fidelity,
  byte-deterministic preflight, AST scope guard.
- `tests/test_xlsx_understand.py` — Phase 2A Document Map: fixtures
  for every region/header/table/repeat case, no-write SHA checks,
  byte-determinism, UTF-8 output, structured error paths, plus a
  fixture oracle proving the raw-OOXML structure pass matches openpyxl
  field-for-field (merges, hidden state, tables, validations, CF,
  freeze/filter, sheet state, chart/image counts, structure
  fingerprint).

Latest local run: 161 passed, 1 skipped (the chart/image
fidelity check skips because `pillow` is not installed in the pinned
runner; add `--with pillow` to enable it).
