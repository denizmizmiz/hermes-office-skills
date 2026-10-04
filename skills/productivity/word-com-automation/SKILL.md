---
name: word-com-automation
description: Use when converting .doc or exporting PDF on Windows.
version: 1.0.0
author: denizmizmiz (with Hermes Agent)
license: MIT
platforms: [windows]
metadata:
  hermes:
    tags: [word, docx, doc, com, conversion, pdf, office, formatting]
    category: productivity
    related_skills: [docx, pdf, xlsx]
---

# Word COM Automation (Windows)

Convert between Word formats and export PDF by driving **Microsoft Word
via COM**, and make **format-preserving edits** to `.docx` with
python-docx. Use this when LibreOffice is not installed, or when the file
is legacy `.doc` (python-docx and the `docx` skill cannot open `.doc`).

## When to Use

- A `.doc` / `.rtf` / legacy Word file must be read or edited (python-docx
  cannot open it).
- Convert `.doc` -> `.docx`, or save an updated doc back to `.doc`.
- Produce a PDF from a `.docx` and LibreOffice (`soffice`) is absent.
- Edit an existing `.docx` while keeping its original layout/formatting
  (add table rows, replace cell text) — rebuilding from scratch would lose
  the template look.
- Not for: greenfield `.docx` creation (use the `docx` skill), `.odt`, or
  non-Windows hosts (Word COM does not exist there).

## Prerequisites / detection

1. Check for a headless renderer first: `command -v soffice || command -v libreoffice`.
   If present, use it — it is simpler and cross-platform.
2. Else check for Word:
   `ls "/c/Program Files/Microsoft Office/root/Office16/WINWORD.EXE"`.
   If Word exists, drive it via PowerShell COM (below). Word is often
   installed even when LibreOffice is not.

## Procedure

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

## Pitfalls

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

## Related bundled skills

`docx` (create/read/edit `.docx`), `pdf` (PDF ops), `xlsx`,
`powerpoint`. This skill only ADDS the COM conversion layer and the
row-cloning edit recipe those skills do not cover.
