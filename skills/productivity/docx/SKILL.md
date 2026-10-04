---
name: docx
description: "Create, read, edit, template, and review Word .docx files (Windows: convert legacy .doc/.rtf and export PDF via Word COM)."
version: 1.3.0
author: Nous Research
license: MIT
platforms: [linux, macos, windows]
metadata:
  hermes:
    tags: [word, docx, documents, office, templates, revisions, comments]
    category: productivity
    related_skills: [pdf, xlsx, powerpoint]
---

# Docx Skill

Create, read, edit, and template Microsoft Word `.docx` files with
python-docx via small CLIs. It handles text, styles, lists, tables,
images, headers/footers, `{{token}}` templating, tracked changes
(list/accept/reject), comments (list/add/delete), TOC and page-number
fields, and package health checks. It does not render documents itself
(PDF needs LibreOffice — see Converting to PDF) or edit legacy `.doc`.

On **Windows**, a Word COM layer (v1.2) also converts legacy `.doc` /
`.rtf` to `.docx`, exports PDFs, and does format-preserving edits —
see "Legacy .doc, .rtf, and PDF via Word COM (Windows)".

## When to Use

- The user asks to generate a Word document (report, letter, contract).
- You need the text, outline, styles, or embedded images of a `.docx`.
- You must change an existing `.docx`: replace text, edit table cells,
  insert/delete paragraphs, apply styles, merge fragmented runs.
- You have a `.docx` template with `{{placeholders}}` to fill from data.
- The document has tracked changes to review, accept, or reject.
- You need to read reviewers' comments, or add/delete comments.
- A `.docx` won't open or behaves oddly and you need corruption triage.
- The document needs a table of contents or "Page X of Y" footers.
- On Windows: a legacy `.doc` / `.rtf` file must be read or edited
  (python-docx cannot open it), or a PDF must be produced when
  LibreOffice is absent — see the Word COM section.
- A document **set** (QMS-style corpus) must be made sense of — parse
  codes/revisions/dates from file names, resolve current vs historical
  revisions, extract cross-references, dedupe archive mirrors: see
  "Deep Document Understanding".
- Not for: `.doc` (legacy) without Word COM available, `.odt`, or
  WYSIWYG layout work.

## Prerequisites

- Python 3.10+ with `python-docx` installed:
  `pip install python-docx` (import name is `docx`; lxml comes with it).
- Comments `add` uses the native API on python-docx >= 1.2 and an XML
  fallback on older versions — both are automatic.
- For image blocks: the image files must exist locally (PNG/JPEG).

## How to Run

All helpers live in `scripts/` next to this file. Run them with the
`terminal` tool; each supports `--help` and prints JSON to stdout.

```bash
python scripts/docx_create.py spec.json out.docx
python scripts/docx_read.py out.docx --text
python scripts/docx_edit.py replace out.docx --find old --replace new
python scripts/docx_template.py tpl.docx values.json filled.docx
python scripts/docx_revisions.py list out.docx
python scripts/docx_comments.py list out.docx
python scripts/docx_validate.py out.docx
python scripts/docx_deep_extract.py f.docx        # deep JSON (identity+structure+refs)
```

## Quick Reference

| Task | Command |
| --- | --- |
| Create from JSON spec | `docx_create.py spec.json out.docx` |
| Full text (body+tables+headers/footers) | `docx_read.py f.docx --text` |
| Heading outline + table shapes | `docx_read.py f.docx --structure` |
| Styles actually used | `docx_read.py f.docx --styles` |
| Extract embedded images | `docx_read.py f.docx --images outdir/` |
| Detect tracked changes/comments | `docx_read.py f.docx --revisions` |
| Find/replace (formatting kept) | `docx_edit.py replace f.docx --find A --replace B -o out.docx` |
| Set a table cell | `docx_edit.py set-cell f.docx --table 0 --row 1 --col 2 --text X` |
| Insert paragraph before index N | `docx_edit.py insert f.docx --index N --text X --style Normal` |
| Delete paragraph N | `docx_edit.py delete f.docx --index N` |
| Apply style to paragraph N | `docx_edit.py style f.docx --index N --style "Heading 1"` |
| Merge equal-format adjacent runs | `docx_edit.py normalize f.docx -o out.docx` |
| Insert TOC field before para N | `docx_edit.py toc f.docx --index N -o out.docx` |
| "Page X of Y" footer fields | `docx_edit.py page-numbers f.docx` |
| Fill `{{tokens}}` | `docx_template.py tpl.docx values.json out.docx --strict` |
| List revisions (id/author/date/text) | `docx_revisions.py list f.docx` |
| Accept / reject all revisions | `docx_revisions.py accept-all f.docx -o out.docx` (or `reject-all`) |
| Accept / reject one revision | `docx_revisions.py accept f.docx --id 3 -o out.docx` |
| List comments (+anchored text) | `docx_comments.py list f.docx` |
| Add comment anchored to text | `docx_comments.py add f.docx --target "phrase" --text "note" --author You` |
| Delete comment by id | `docx_comments.py delete f.docx --id 0` |
| Health-check the package | `docx_validate.py f.docx` (exit 1 on errors) |
| Deep-parse identity+structure+refs | `docx_deep_extract.py f.docx [f2 ...]` (JSON) |

## Procedure

1. **Create.** Write a JSON spec with `write_file`, then run
   `scripts/docx_create.py`. The spec supports: `page` (size + margins in
   mm), `header`/`footer` strings, `footer_page_numbers` (adds a
   "Page X of Y" field footer), `styles` (custom paragraph styles with
   font, size, bold/italic, hex `color`), and `blocks` — `heading`
   (level 1-9), `paragraph` (either `text` or a `runs` list where each run
   may set `bold`/`italic`/`underline`), `bullet_list`, `numbered_list`,
   `table` (`header` row rendered bold, `rows`, optional built-in table
   `style` such as `Table Grid`), `image` (`path`, optional `width_mm`),
   `toc` (Table of Contents field), and `page_break`. The full spec
   format is documented at the top of `scripts/docx_create.py`.
2. **Read.** Use `scripts/docx_read.py` with exactly one mode flag.
   `--text` returns body paragraphs, all table cell text, and
   header/footer text as JSON. `--structure` returns the heading outline
   plus paragraph/table/section counts. `--images DIR` copies every file
   under `word/media/` out of the package.
3. **Edit.** Use `scripts/docx_edit.py`. `replace` walks body, tables
   (nested included), headers and footers, and preserves run formatting;
   add `--body-only` to skip headers/footers. Pass `-o out.docx` to keep
   the original; omit it to edit in place. Paragraph indices for
   `insert`/`delete`/`style`/`toc` refer to `--structure`/`--text` body
   order. Run `normalize` first on documents that came out of heavy Word
   editing — it merges adjacent runs with identical formatting so later
   find-replace matches reliably.
4. **Review revisions.** `docx_revisions.py list` reports every `w:ins`
   and `w:del` (id, author, date, affected text) anywhere in body,
   tables, headers, or footers. `accept-all` / `reject-all` resolve them
   in bulk; `accept`/`reject --id N` handles a single revision. Accept
   keeps insertions and drops deleted text; reject does the reverse.
5. **Comments.** `docx_comments.py list` returns each comment's id,
   author, date, body text, and the document text it is anchored to.
   `add --target "some phrase"` anchors a new comment to the first
   occurrence of that phrase (runs are split as needed; formatting is
   preserved). `delete --id N` removes the comment and its markers
   without touching document text.
6. **Template.** Put `{{name}}`-style tokens in the document. Run
   `scripts/docx_template.py` with a JSON object of values. Use
   `--strict` to fail when tokens remain unfilled; the JSON output lists
   `filled` counts and `unfilled_tokens` either way.
7. **Verify** (always): re-read the output with `--text` or
   `--structure`, and run `docx_validate.py` on anything you produced
   via revision/comment surgery.

## Legacy .doc, .rtf, and PDF via Word COM (Windows)

On Windows, when LibreOffice is absent or the file is legacy `.doc` /
`.rtf` (python-docx cannot open those), drive **Microsoft Word via COM**
through PowerShell for conversion and PDF export, and use python-docx
for format-preserving `.docx` edits.

### Prerequisites / detection

1. Check for a headless renderer first: `command -v soffice || command -v libreoffice`.
   If present, use it — it is simpler and cross-platform.
2. Else check for Word:
   `ls "/c/Program Files/Microsoft Office/root/Office16/WINWORD.EXE"`.
   If Word exists, drive it via PowerShell COM (below). Word is often
   installed even when LibreOffice is not.

### 1. Run PowerShell from the bash terminal

Write the PowerShell to a file with write_file, run it with `-File`, and
pass a NATIVE path (bash `cd`/globs work, but the interpreter path must be
Windows-style). Convert an MSYS path with `cygpath -w`:

```bash
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(cygpath -w /tmp/conv.ps1)"
```

### 2. Convert `.doc` -> `.docx`

```powershell
$ErrorActionPreference='Stop'
$word=New-Object -ComObject Word.Application
$word.Visible=$false
$doc=$word.Documents.Open($src,$false,$true)   # ConfirmConversions, ReadOnly
$doc.SaveAs([ref]$dst,[ref]16)                  # 16 = wdFormatDocumentDefault (.docx)
$doc.Close($false); $word.Quit()
```

SaveAs format codes: **16** = `.docx`, **0** = legacy `.doc`,
**1** = `.dotx` template, **2** = plain text.

### 3. Export to PDF

```powershell
$doc.ExportAsFixedFormat($dstPdf, 17)   # 17 = wdExportFormatPDF
```

`ExportAsFixedFormat(OutputFileName, ExportFormat)` renders layout
faithfully; python-docx cannot produce PDF.

### 4. Inspect before editing

The converted `.docx` is a normal package — read it with the `docx`
skill scripts (`docx_read.py --text` / `--structure`) to map tables and
rows before touching anything.

### 5. In-place, format-preserving edits (python-docx)

Replace cell text WITHOUT resetting the run formatting (never use
`cell.text = ...`, it wipes runs):

```python
def set_cell(cell, text):
    p = cell.paragraphs[0]
    for extra in cell.paragraphs[1:]:          # drop extra paragraphs
        extra._element.getparent().remove(extra._element)
    runs = p.runs
    if runs:
        runs[0].text = text                     # keep first run's rPr
        for r in runs[1:]:
            r._element.getparent().remove(r._element)
    else:
        p.add_run(text)
```

Clone a table row (deep-copy the XML, insert after an anchor):

```python
import copy
from docx.oxml.ns import qn
from docx.table import _Cell
tr = copy.deepcopy(table.rows[template_idx]._tr)
anchor_tr.addnext(tr); anchor_tr = tr
# reach cloned cells by their <w:tc> children:
cell = _Cell(tr.findall(qn('w:tc'))[col], table)
```

### 6. Verify

- Re-read the output (`docx_read.py --text`) and assert the new strings
  are present and old ones gone.
- Convert/export, then read the PDF back (`read_file`) to confirm the
  rendered result — a successful COM call is not proof of correct layout.

### COM pitfalls

- **HOLD `_tr` ELEMENT REFERENCES, NOT ROW INDICES, across insertions.**
  After `addnext` the table's `rows[...]` indices shift, so an index-based
  "insert before row 15" silently targets the wrong row. Capture
  `anchor = table.rows[15]._tr` once and chain `anchor.addnext(new_tr)`.
- **`cell.text = x` destroys formatting.** Assign to `runs[0].text` and
  delete the remaining runs (see step 5). Apply the same rule when editing
  headers/footers.
- **A cloned row inherits the template row's text.** Always overwrite every
  cell of the clone (including the section header cell), or duplicate text
  leaks into the new block.
- **Word COM needs `[ref]` wrappers on `SaveAs`** and `Visible=$false`;
  open the source ReadOnly (`$true` as the 3rd `Open` arg) for pure
  conversions so the original is not locked or modified.
- **Only add a real, source-backed entry when consolidating documents.**
  When updating one document from others, replace stale fields with values
  found in the sources; do NOT invent absent dates/details. Flag residual
  conflicts (e.g. two sources disagreeing on a date, or a field one source omits) to the user instead of silently picking one.
- **Match the sources' date granularity.** When sources give month-level
  periods, write month-level (`10.2020 – 11.2024`) rather than fabricating
  a day.
- **Apply the user's stated correction literally, then surface any contradiction it
  creates.** If a corrected value makes entries overlap (e.g. an end date before
  the previous entry starts), either apply it and ask, or reconcile and say which
  value you changed — never silently "repair" a value the user supplied.
- **Back up the original before overwriting it** (dated copy beside it, or save
  the update under a new filename) and report where the backup is.

## Deep Document Understanding (document sets)

When the task is a SET of Word files — a QMS, a policy corpus, an audit
archive — the single-file commands above are the wrong level. Use
`scripts/docx_deep_extract.py` (per-file deep JSON: identity + structure +
cross-references) together with the rules in
`references/deep-document-understanding.md` (the six-step loop:
identity → package facts → structure → fallback ladder → revision
resolution → cross-references).

```bash
python scripts/docx_deep_extract.py "path/PR.01DocumentControlProcedure-R6&01.01.2026.docx"
python scripts/docx_deep_extract.py "corpus_dir" --out deep.json   # recursive *.docx
```

It returns `identity` (code/prefix/family/rev/rev_raw/eff_date/lang/title),
the heading tree and table transcription (counts + samples),
`numbered_outline_guess` (heuristic outline for style-less documents),
`refs` (other document codes mentioned in the text) and package flags
(tracked changes, comments, macros). Read-only; never modifies inputs.

Four rules that make or break corpus work (full list in the reference):

1. **Deduplicate by SHA-256 first** — archives mirror themselves 2-3×;
   parse each unique byte-stream once, but keep every path.
2. **Names are metadata — but glued names lie.** `PR.01DocumentControl...`
   (no separator) must still parse as `PR.01`; use the reference's regex,
   not `(?![0-9a-z])` guards, or a whole family of Turkish QMS documents
   silently drops out of the register.
3. **`Rev/` folders are history; `Arşiv/` are mirrors.** Current = newest
   (effective date, then revision); date beats revision because series get
   reset (R10 → R1 seen in the wild). Conflicts get flagged TENTATIVE —
   never auto-promoted.
4. **Carry the full fallback ladder**: python-docx → raw
   `word/document.xml` (survives the `word/NULL` KeyError) → NOT_ZIP/EMPTY
   classification. A pipeline that only knows the happy path will report
   wrong failure lists.

## Converting to PDF

No script needed. When LibreOffice is installed, convert headlessly:

```bash
soffice --headless --convert-to pdf --outdir outdir/ file.docx
```

Check availability first (`command -v soffice || command -v
libreoffice`). If neither exists, tell the user PDF conversion is
unavailable in this environment rather than improvising — python-docx
cannot render PDFs, and layout fidelity requires a real renderer.

## Pitfalls

- **Tokens split across runs.** Word often fragments text into several
  runs. The replace helpers collapse matched runs (replacement inherits
  the first run's formatting); running `docx_edit.py normalize` first
  reduces fragmentation for all later edits.
- **Revision coverage.** `docx_revisions.py` resolves run-level
  insertions and deletions (the overwhelming majority). Paragraph-mark
  and table-row revisions, format-change records, and moves are detected
  by `--revisions` but not auto-resolved — see
  `references/revisions-and-comments.md` and hand those to Word.
- **Comment threading.** Replies and "resolved" status live in
  `commentsExtended.xml`, which this skill ignores; comments it adds are
  plain top-level comments.
- **Field results are computed by Word.** `toc`, `page-numbers`, and the
  `toc`/`footer_page_numbers` spec options write *field codes*.
  Word/LibreOffice populates the actual entries and numbers when the
  file is opened (Word may prompt to update fields); python-docx never
  computes them, so placeholder text shows until then.
- **Validation is a health check, not schema validation.**
  `docx_validate.py` verifies the zip, required parts, relationship
  targets, image magic bytes, and referenced styles. It is NOT XSD
  validation — a file can pass and still contain XML Word dislikes.
- **Style names must exist.** Applying a style that isn't defined in the
  document raises `KeyError`. Built-ins like `Heading 1`, `List Bullet`,
  `List Number`, `Table Grid` exist in the default template; custom
  styles must be declared in the create spec first.
- **Numbered lists restart.** `List Number` relies on Word's default
  numbering; separate lists in one document may continue numbering
  instead of restarting. Warn users needing precise multi-list numbering.
- **Cell writes replace formatting.** `set-cell` uses `cell.text = ...`,
  which resets runs in that cell to plain formatting.
- **Encoding.** All JSON specs/values files are read as UTF-8 explicitly;
  never rely on locale defaults when writing your own glue code.
- **Don't unzip-and-sed the XML.** Edit through the scripts (or
  python-docx); raw text substitution in `document.xml` corrupts files
  easily. Use `patch`/`write_file` only for the JSON inputs, never on the
  `.docx` itself.

## Verification

- After create/edit/template, run `docx_read.py out.docx --text` and
  check the expected strings appear (and old strings are gone).
- After accept/reject, `docx_revisions.py list` should return `[]` (or
  only the ids you intentionally left); after comment surgery,
  `docx_comments.py list` should reflect the change and `--text` output
  must be unchanged.
- `docx_validate.py out.docx` exits 0 with `"ok": true` on a healthy
  package — run it after any revision/comment/field manipulation.
- For templates run with `--strict`, or check `unfilled_tokens == []`.
- Structure checks: `--structure` should show the expected heading
  outline and table shapes; `--styles` confirms custom styles applied.
- Deep parse: `docx_deep_extract.py f.docx` returns a precise `status`
  (OK / EMPTY / NOT_ZIP / FAILED), a sane `identity` (code/rev/date) and
  the heading outline; damaged packages still succeed via
  `fallback: raw_xml`.
