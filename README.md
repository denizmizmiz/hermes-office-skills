# Hermes Office Skills

Enhanced **office file skills** for [Hermes Agent](https://github.com/NousResearch/hermes-agent):

| Skill | Version | What it does |
|---|---|---|
| [`xlsx`](skills/productivity/xlsx) | 1.3.0 | Excel `.xlsx` / CSV toolkit — read, create, edit, restructure, validate. Safe-write chain (backup → temp → verify → atomic commit) with an operation manifest; Document Map → semantic profiles → mapping / dry-run → approved execution pipeline; Formula Intelligence (parse / family / graph / gated synthesis). |
| [`docx`](skills/productivity/docx) | 1.3.0 | Word `.docx` toolkit — create, read, edit, template, validate; comments and tracked changes; **Deep Document Understanding** layer for document sets (identity parsing, structure, cross-references). |
| [`word-com-automation`](skills/productivity/word-com-automation) | 1.0.0 | Windows-only: drive Microsoft Word via COM — legacy `.doc` → `.docx` conversion, PDF export, and format-preserving in-place edits. |

## Install

Copy a skill folder into your Hermes skills directory:

```bash
# default home
cp -r skills/productivity/xlsx ~/.hermes/skills/productivity/

# or a specific profile
cp -r skills/productivity/xlsx ~/.hermes/profiles/<profile>/skills/productivity/
```

Dependencies: `openpyxl` (xlsx), `python-docx` (docx) — see
`skills/productivity/xlsx/scripts/requirements.txt`.
`word-com-automation` needs Microsoft Word installed (it falls back to
LibreOffice when that is present).

## What's different from the bundled skills

- **xlsx 1.1.0 → 1.3.0** — 16 new modules on top of the stock scripts: safe-write execution layer with recovery + fault injection, Document Map, semantic profiles & mapping, Formula Intelligence, richer validation/charts/tables round-trips, and a full pytest suite (160+ tests).
- **docx 1.1.0 → 1.3.0** — new `docx_deep_extract.py` (document-set understanding), plus expanded guides for revisions/comments and Windows Word COM workflows.
- **word-com-automation** — new companion skill covering the Windows-only gaps (legacy `.doc`, PDF export, format-preserving edits).

## Upstream

These versions are offered upstream to the Hermes project:
[NousResearch/hermes-agent#132875](https://github.com/NousResearch/hermes-agent/issues/132875)

## Credits & license

- `xlsx` / `docx` are derived from the bundled skills in
  [NousResearch/hermes-agent](https://github.com/NousResearch/hermes-agent)
  (MIT, © Nous Research — see `LICENSE` inside each skill folder).
- v1.3.0 enhancements and `word-com-automation`:
  © 2026 Deniz Mizmizlioğlu ([@denizmizmiz](https://github.com/denizmizmiz)), MIT.
- Community-maintained; not an official Nous Research product.
