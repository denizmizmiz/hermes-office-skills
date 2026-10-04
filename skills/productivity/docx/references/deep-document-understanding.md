# Deep Document Understanding — sense-making and context-building for .docx corpora

This skill's single-file commands handle ONE document you already understand.
This reference is for the opposite case: a SET of hundreds/thousands of Word
files (a QMS, a policy corpus, an audit archive) where the job is to work out
what each file IS, which revision is CURRENT, how the documents RELATE, and
what the system actually DOES — without modifying anything.

Distilled from a read-only deep analysis of a large real-world QMS
archive (thousands of files, hundreds of canonical documents). The full
corpus engine (name parsing, office handlers) is not included here —
adapt the recipes below to your own stack.

## The six-step loop

### 1. Identity — parse the file NAME first
The name is the cheapest, highest-signal metadata. Parse: code, prefix,
family, revision, effective date, language, title.

Key trap — **glued names**: `PR.01DocumentControlProcedure-R6&01.01.2026.docx`
has NO separator between code and title. A strict `(?![0-9a-z])` guard drops
~1,000 files (the whole PR.01/02/05/07 family, TB.302, …) silently into
"unregistered". Proven regex — allows a glued title, still requires any
single-letter revision suffix (`FR.508a`) to be lowercase:

```python
CODE_RE = re.compile(r'^(?i:KEK|OEK|QM|PR|TL|TB|LS|FR|DD|SD|PG)[._\- ]?'
                     r'(\d{2,3}(?:[a-z](?![A-Za-z]))?)(?![0-9])')
```

(no global re.I — the scoped `(?i:…)` keeps `[a-z]` genuinely lowercase.)

Other pieces: revisions — last plausible token of `Rev 3`, `R6`, `REV2&…`
(see pattern list in `docparse.py`); dates `dd.mm.yyyy` (Turkish order —
`01.08.2026` is 1 August) preferring the token AFTER the revision; language
via explicit `TR`/`EN` tokens plus Turkish/English word hints; title = name
minus code/rev/date noise. Family from prefix: PR→PROCEDURE, TL→INSTRUCTION,
FR→FORM, LS→LIST, TB→TABLE, QM/KEK/OEK→MANUAL, DD/SD→by title (job
description / process / policy) else SUPPORT_DOC.

### 2. Package facts (cheap, no style parsing)
From the zip: part count; `word/comments.xml` present?; `word/vbaProject.bin`
(macros!); `word/media/*` count; count `<w:ins `/`<w:del ` in
`word/document.xml` for tracked changes; `<w:hyperlink` count; core
properties (title/author/created/modified).

### 3. Structure
- Heading tree: match paragraph style NAME `Heading N` OR style id
  `HeadingN` (Word stores either); Title/Subtitle → level 0. First resolve
  custom style IDs through `word/styles.xml` — Turkish documents use IDs
  like `Balk1` (w:name = "heading 1") or `KonuBal` (w:name = "Title") that
  a naive style check silently misses.
- Style-less documents: fall back to a numbered-line heuristic (`1.`,
  `2.3.` short lines → `numbered_outline_guess`); report it as a guess,
  never as authored structure.
- Tables: transcribe row-major with caps (≤80 tables, ≤40 rows, ≤20 cols,
  ≤300 chars/cell) — full fidelity for structure, capped for size.
- Headers/footers per section (joined paragraph text).

### 4. Fallback ladder (read-only, never "repair" the file)
1. python-docx fails `KeyError: "There is no item named 'word/NULL'…"`
   (broken relationship) → **raw-XML fallback**: read `word/document.xml`
   from the zip, iterate `w:p`/`w:tbl` (ElementTree); heading levels still
   recoverable from `w:pPr/w:pStyle/@val` (`Heading1`, …).
2. `BadZipFile` → sniff: 0 bytes = EMPTY; anything else = NOT_ZIP (record).
3. Legacy `.doc` `NotOleFileError` → file is not OLE either → corrupt; a real
   `.doc` needs Word COM (see the skill's COM section).
A corpus pipeline must carry all three or its failure list will lie.

### 5. Corpus context — duplicates and revisions
- **Deduplicate FIRST** by SHA-256; byte-identical copies get a stub status
  and are never deep-parsed twice (in the real corpus ~60% of files were
  archive mirrors). Keep every path, parse each unique byte-stream once.
- Group by code → canonical register. Within a code:
  - `Rev/`(or `rev`) folders = historical revisions; `Arşiv/`/`Archive` =
    mirrors; `_to_delete/` = pending deletion — excluded from the current pool.
  - current = newest (effective date, then revision number). **Date beats
    revision** — revision series get reset (a real R10 → R1 re-issue).
  - Flag and mark TENTATIVE — never auto-promote — when: revision order and
    date order disagree; TR and EN copies sit on different revisions; a
    history folder holds something newer than the current folder; same rev
    seen with several dates; missing rev/date metadata.

### 6. Cross-references → the context map
Scan extracted text (paragraphs + table cells) for other codes
(`\b(?:KEK|OEK|QM|PR|TL|TB|LS|FR|DD|SD|PG)[._ ]?\d{2,3}[a-z]?(?![0-9])` and
external `(?:ASR|CCS|GRS|OCS|RCS)[-_.]?\d{2,3}(?![0-9])`), normalize
identifiers before comparing (`-`/`_`/`.` vary: `CCS-102` must equal
`CCS.102`), drop self-references. The edge list is the corpus graph: who
references whom, most-referenced docs, orphans, and "referenced but NOT
present in the share" codes (a genuinely useful finding). For procedures,
the text then supports sense-making outputs (roles, decision points, process
steps) — always stored with the source path (provenance) and labeled as
derived.

## Pitfalls (hard-won)
- Glued titles (step 1) — the single biggest silent failure mode.
- Subcodes: `FR.508` vs `FR.508a` — keep the letter suffix; treat a
  `FR.508` mention as matching an `FR.508a` file at mention level.
- Dates are Turkish `dd.mm.yyyy` — do not read them as US month-first.
- Copies everywhere: expect 2-3× duplication (archive mirrors, per-year folders).
- Don't OCR scanned PDFs (static analysis, no side files): record
  SCANNED_NO_TEXT and move on. 0-byte files are EMPTY, not failures.
- Keep every record's relative path — provenance is the deliverable's spine.
- Never write to the corpus; outputs (index, JSON, maps) live in a separate
  workspace and are quoted back with `rel` paths.
