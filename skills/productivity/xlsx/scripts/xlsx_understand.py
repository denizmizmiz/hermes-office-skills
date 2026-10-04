#!/usr/bin/env python3
# MIT License. Part of the Hermes xlsx skill.
"""Understand an .xlsx workbook and print a deterministic Document Map.

READ-ONLY by construction: this module has no write path and never calls
a workbook save (audited by the test suite). It analyses a file in three
passes whose state never mixes:

  Pass A  read_only=True, data_only=False
          values, formula text, coordinates, compact styles, row/column
          visibility profile, dimensions
  Pass B  read_only=True, data_only=True
          cached formula values and ONLY their availability; a cached
          value is never treated as a recalculation -- the workbook map
          always reports values_recalculated=false.
  Pass S  raw OOXML parts (zip + XML, no workbook object at all)
          tables, merged cells, hidden rows/columns, validations,
          conditional formats, freeze panes, autofilter, defined
          names, external links, calculation settings, chart/image
          counts -- read-only zip reads; the test suite proves the
          input SHA-256 is unchanged after every run.

Document Map root keys: ok, workbook, sheets, locale, warnings[],
unsupported[], diagnostics[]. Every inference carries evidence[] and a
confidence score/level (LOW detections always surface a warning -- no
silent detection, no silent skip).

Usage:
  xlsx_understand.py book.xlsx                  compact canonical JSON
  xlsx_understand.py book.xlsx --pretty         indented JSON
  xlsx_understand.py book.xlsx --samples 3      sample values per column

Determinism: same file + same options -> byte-identical stdout (the map
core carries no timestamps; execution timing is measured externally).

Limits (Phase 2A, by design): no template filling, no source->target
mapping, no row expansion, no formula propagation, no formula-pattern
engine -- discovery and detection only.
"""
from __future__ import annotations

import argparse
import json
import posixpath
import re
import sys
import unicodedata
import zipfile
from datetime import date, datetime
from pathlib import Path

try:  # defusedxml ships with the openpyxl test extra
    from defusedxml import ElementTree as _ET
except ImportError:  # pragma: no cover - environment guard
    import xml.etree.ElementTree as _ET  # noqa: S405

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_common as common  # noqa: E402

try:  # dependency guard: structured DEPENDENCY_MISSING, never a traceback
    from openpyxl.utils import (  # noqa: F401
        coordinate_to_tuple,
        get_column_letter,
        range_boundaries,
    )
    # Reuse the single A1-reference parser of xlsx_restructure.py
    # (REF_RE/COORD_RE/STRING_RE); Phase 2A never writes a second parser.
    from xlsx_restructure import STRING_RE, extract_refs  # noqa: F401
except ImportError as _exc:  # pragma: no cover - environment guard
    sys.exit(common.dependency_error(_exc))


# ---------------------------------------------------------------------------
# constants  (all thresholds live here so the tests can pin behaviour)
# ---------------------------------------------------------------------------

SAMPLE_DEFAULT = 5          # default --samples for column sample_values
SAMPLE_MAX = 100            # --samples is clamped here (with a warning)
FORMULA_LIST_CAP = 5000     # listed formula entries per sheet; aggregates
                            # (formula_summary) always cover ALL formulas
SAMPLE_CHAR_LIMIT = 60      # sample strings longer than this get "..."

SCAN_ROW_CAP = 250_000      # scan stops here WITH a warning (never silent)
ROW_VALUE_BUDGET = 500_000  # per-row label values kept in profiles

REGION_TYPES = (
    "title", "section_header", "header", "subheader", "data",
    "formula_region", "subtotal", "total", "footer", "notes",
    "input_region",
)

# compact, deterministic vocabularies for summary-row heuristics
TOTAL_WORDS = {
    "toplam", "toplamı", "toplami", "total", "totals", "genel toplam",
    "genel toplamı", "genel toplami", "grand total", "grand totals",
    "büyük toplam", "buyuk toplam",
}
SUBTOTAL_WORDS = {
    "ara toplam", "ara toplamı", "ara toplami", "ara toplamlar",
    "subtotal", "sub total", "subtotals", "ara toplam:",
}
NOTES_WORDS = (
    "not:", "not;", "not ", "notlar", "notes", "note:", "remark",
    "remarks", "açıklama", "açıklamalar", "aciklama", "aciklamalar",
)

FUNC_RE = re.compile(r"(?<![\w.])([A-Za-z_][A-Za-z0-9_.]*)\s*\(")
CMP_RE = re.compile(r"<=|>=|<>|[<>=]")
ARITH_CHARS = frozenset("+-*/^")

HEADER_SCORE_MIN = 55       # below this a band's first row is not a header
HEADER_SCORE_HIGH = 80
HEADER_SCORE_MEDIUM = 65

ROW_SIG_TYPE_CODES = {
    "string": "t", "integer": "i", "decimal": "n", "date": "d",
    "datetime": "d", "boolean": "b", "formula": "q", "blank": ".",
}


# ---------------------------------------------------------------------------
# value + text helpers  (deterministic, stdlib only)
# ---------------------------------------------------------------------------

def normalize_header(text):
    """Deterministic header normalization (Turkish-first, non-semantic).

    NFC -> odd spaces to plain space -> whitespace collapse -> casefold ->
    drop the combining dot that casefolding capital dotted-I introduces
    (so "DENETİM" and "Denetim" normalize alike) -> strip edge
    punctuation -> inner . _ - dashes to space -> collapse again.

    Dotless i stays distinct from i by design, and NO synonym mapping is
    done in Phase 2A ("ÜRÜN NO" / "Article Number" are not equated).
    """
    if text is None:
        return None
    s = str(text)
    if s.startswith("\ufeff"):
        s = s[1:]
    s = unicodedata.normalize("NFC", s)
    s = s.replace("\u00a0", " ").replace("\u202f", " ").replace("\u2007", " ")
    s = re.sub(r"\s+", " ", s).strip()
    s = s.casefold()
    s = s.replace("\u0307", "")
    s = s.strip(".:;,·•*\"'")
    s = re.sub(r"[._\-–—]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def truncate_text(text, limit=SAMPLE_CHAR_LIMIT):
    """String samples are truncated deterministically ('...' suffix)."""
    text = str(text)
    return text if len(text) <= limit else text[: limit - 3] + "..."


def sample_value(value):
    """JSON-safe, truncated sample of a cell value (formulas keep text)."""
    v = common.jsonable(value)
    if isinstance(v, str):
        return truncate_text(v)
    return v


def classify_scalar(value):
    """Scalar type tag used by column and row analysis (bool before int)."""
    if value is None:
        return "blank"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, int):
        return "integer"
    if isinstance(value, float):
        return "decimal"
    if isinstance(value, datetime):
        return "datetime"
    if isinstance(value, date):
        return "date"
    if isinstance(value, str):
        return "formula" if value.startswith("=") else "string"
    return "string"


def rollup_column_type(type_counts):
    """Collapse per-cell type counts into the column's single data_type.

    Formula cells are recorded separately (formula_count/formula_share)
    and never force a column to 'formula' unless it contains ONLY
    formulas. integer+decimal mix -> decimal; date+datetime mix ->
    datetime; any other multi-type mix -> mixed.
    """
    non_blank = {t: c for t, c in type_counts.items()
                 if t not in ("blank", "formula") and c > 0}
    has_formula = type_counts.get("formula", 0) > 0
    if not non_blank:
        return "formula" if has_formula else "blank"
    kinds = set(non_blank)
    if kinds == {"integer", "decimal"}:
        return "decimal"
    if kinds <= {"date", "datetime"}:
        return "datetime"
    if len(kinds) == 1:
        return next(iter(kinds))
    return "mixed"


def classify_summary_label(norm):
    """'total'/'subtotal' for a normalized left-most label, else None."""
    if not norm:
        return None
    if norm in SUBTOTAL_WORDS:
        return "subtotal"
    if norm in TOTAL_WORDS:
        return "total"
    if norm.startswith("ara toplam"):
        return "subtotal"
    for word in ("genel toplam", "grand total", "büyük toplam",
                 "buyuk toplam"):
        if norm.startswith(word):
            return "total"
    if norm.startswith("toplam") or norm.startswith("total"):
        return "total"
    return None


def looks_like_notes(norm):
    return bool(norm) and any(norm.startswith(w) for w in NOTES_WORDS)


def read_cell_style(cell, failures):
    """Compact style trio for one cell, per-property guarded (Phase-0.5).

    Unreadable sub-properties never take the whole map down: the failure
    kind is recorded once per sheet and surfaces as a STYLE_READ_PARTIAL
    warning. Returns (number_format|None, bold|None, italic|None,
    filled|None).
    """
    numfmt = bold = italic = filled = None
    try:
        numfmt = cell.number_format
    except Exception as exc:  # noqa: BLE001
        failures.add(("number_format", type(exc).__name__))
    try:
        font = cell.font
        bold = bool(font.b) if font is not None else None
        italic = bool(font.i) if font is not None else None
    except Exception as exc:  # noqa: BLE001
        failures.add(("font", type(exc).__name__))
    try:
        fill = cell.fill
        filled = bool(fill is not None and fill.patternType)
    except Exception as exc:  # noqa: BLE001
        failures.add(("fill", type(exc).__name__))
    return numfmt, bold, italic, filled


def cell_coordinate(cell):
    """Coordinate of a real cell, or None for read-only EmptyCell gaps."""
    return getattr(cell, "coordinate", None)


def cell_value(cell):
    return getattr(cell, "value", None)


# ---------------------------------------------------------------------------
# column accumulator  (Pass A)
# ---------------------------------------------------------------------------

class ColumnStats:
    """Per-column accumulator filled during the Pass A scan."""

    __slots__ = ("index", "letter", "type_counts", "non_null", "distinct",
                 "samples", "formula_count", "numfmt_counter", "bold_count",
                 "italic_count", "fill_count", "sample_cap")

    def __init__(self, index, sample_cap):
        self.index = index
        self.letter = get_column_letter(index)
        self.type_counts = {}
        self.non_null = 0
        self.distinct = set()
        self.samples = []
        self.formula_count = 0
        self.numfmt_counter = {}
        self.bold_count = 0
        self.italic_count = 0
        self.fill_count = 0
        self.sample_cap = sample_cap

    def feed(self, value, style):
        numfmt, bold, italic, filled = style
        kind = classify_scalar(value)
        self.type_counts[kind] = self.type_counts.get(kind, 0) + 1
        if kind == "formula":
            self.formula_count += 1
        if value is not None:
            self.non_null += 1
            if len(self.samples) < self.sample_cap:
                self.samples.append(sample_value(value))
            try:
                self.distinct.add(value)
            except TypeError:  # pragma: no cover - scalars are hashable
                self.distinct.add(str(value))
        if numfmt is not None:
            self.numfmt_counter[numfmt] = \
                self.numfmt_counter.get(numfmt, 0) + 1
        if bold:
            self.bold_count += 1
        if italic:
            self.italic_count += 1
        if filled:
            self.fill_count += 1

    def dominant_number_format(self):
        if not self.numfmt_counter:
            return None
        return sorted(self.numfmt_counter.items(),
                      key=lambda kv: (-kv[1], kv[0]))[0][0]

    def to_map(self):
        total_cells = sum(self.type_counts.values())
        formula_share = (round(self.formula_count / self.non_null, 4)
                         if self.non_null else 0.0)
        return {
            "letter": self.letter,
            "index": self.index,
            "data_type": rollup_column_type(self.type_counts),
            "type_counts": dict(sorted(self.type_counts.items())),
            "sample_values": list(self.samples),
            "null_count": self.type_counts.get("blank", 0),
            "cell_count": total_cells,
            "distinct_count": len(self.distinct),
            "formula_count": self.formula_count,
            "formula_share": formula_share,
            "number_format": self.dominant_number_format(),
            "style_summary": {
                "number_formats_distinct": len(self.numfmt_counter),
                "bold_count": self.bold_count,
                "italic_count": self.italic_count,
                "fill_count": self.fill_count,
            },
        }

# ---------------------------------------------------------------------------
# row bands + shared helpers
# ---------------------------------------------------------------------------

def compute_bands(row_profiles):
    """Contiguous runs of rows that contain at least one value."""
    bands, start, end = [], None, None
    for prof in row_profiles:
        if prof["nonempty"] > 0:
            if start is None:
                start = prof["row"]
            end = prof["row"]
        elif start is not None:
            bands.append({"start": start, "end": end})
            start = end = None
    if start is not None:
        bands.append({"start": start, "end": end})
    return bands


def profiles_in_band(profiles_by_row, band):
    return [profiles_by_row[r] for r in range(band["start"],
                                               band["end"] + 1)]


def band_metrics(profiles_by_row, band):
    """(width, min_col, max_col) of a band, from its non-empty cells."""
    width = 0
    min_col, max_col = None, None
    for prof in profiles_in_band(profiles_by_row, band):
        if prof["first_col"] is not None:
            min_col = prof["first_col"] if min_col is None else \
                min(min_col, prof["first_col"])
            max_col = prof["last_col"] if max_col is None else \
                max(max_col, prof["last_col"])
            width = max(width, prof["last_col"])
    return width or 0, min_col, max_col


def range_for_rows(profiles_by_row, r_start, r_end):
    """Canonical A1 range over the used span of rows r_start..r_end."""
    min_col = max_col = None
    for r in range(r_start, r_end + 1):
        prof = profiles_by_row.get(r)
        if prof and prof["first_col"] is not None:
            min_col = prof["first_col"] if min_col is None else \
                min(min_col, prof["first_col"])
            max_col = prof["last_col"] if max_col is None else \
                max(max_col, prof["last_col"])
    if min_col is None:
        return None
    return f"{get_column_letter(min_col)}{r_start}:{get_column_letter(max_col)}{r_end}"


def _conf(score):
    score = max(0, min(100, int(score)))
    if score >= HEADER_SCORE_HIGH:
        return score, "HIGH"
    if score >= HEADER_SCORE_MEDIUM:
        return score, "MEDIUM"
    return score, "LOW"


def _region(kind, r_start, r_end, profiles_by_row, score, evidence):
    score, level = _conf(score)
    return {
        "type": kind,
        "range": range_for_rows(profiles_by_row, r_start, r_end),
        "rows": [r_start, r_end],
        "confidence_score": score,
        "confidence_level": level,
        "evidence": list(evidence),
    }


def _sig_pad(sig, width):
    return sig + "." * max(0, width - len(sig))


def _first_band_is_title(profiles_by_row, bands, merged_rows):
    """Heuristic: the first band is a title block, not a header+data block."""
    if not bands:
        return False, []
    band = bands[0]
    height = band["end"] - band["start"] + 1
    if height > 2 or len(bands) < 2:
        return False, []
    evidence = []
    for r in range(band["start"], band["end"] + 1):
        prof = profiles_by_row[r]
        if prof["nonempty"] > 2:
            return False, []
        if prof["nonempty"] == 0:
            return False, []
        merged_here = [m for m in merged_rows.get(r, [])
                       if m["min_row"] == r and
                       (m["max_col"] - m["min_col"]) >= 1]
        if merged_here:
            evidence.append(f"merged span {merged_here[0]['range']}")
    evidence.append(f"sparse block rows {band['start']}-{band['end']} "
                    "before a second band")
    return True, evidence


# ---------------------------------------------------------------------------
# header detection  (multi-signal, deterministic, evidence-carrying)
# ---------------------------------------------------------------------------

def _header_score_rows(profiles_by_row, band, span, row_values, merged_rows,
                       table_rows, width, result, sheet_name, evidence):
    """Score a header candidate span; returns score after adjustments."""
    r0 = span[0]
    prof0 = profiles_by_row[r0]
    score = 0
    nonempty = prof0["nonempty"]
    text_ratio = prof0["text"] / nonempty if nonempty else 0.0
    fill_ratio = nonempty / width if width else 0.0

    if text_ratio >= 0.8 and nonempty >= 2:
        score += 30
        evidence.append(f"text ratio {int(text_ratio * 100)}%")
    if fill_ratio >= 0.8:
        score += 20
        evidence.append(f"fills {nonempty}/{width} width")
    if nonempty * 2 < width:
        score -= 20
        evidence.append(f"sparse row (only {nonempty} of {width} columns)")

    # type transition into data on the row(s) after the span
    below = span[-1] + 1
    transition_row = None
    for r in (below, below + 1):
        prof = profiles_by_row.get(r)
        if prof is None:
            break
        if prof["num"] + prof["datish"] + prof["bool"] + prof["formula"]:
            transition_row = r
            break
    if transition_row == below:
        score += 20
        evidence.append(f"type transition after row {span[-1]}")
    elif transition_row == below + 1:
        score += 15
        evidence.append(f"type transition after row {span[-1] + 1}")

    raws = [v for _, v in sorted(row_values(r0).items()) if v is not None]
    if len(raws) >= 2 and len({str(v) for v in raws}) == len(raws):
        score += 10
        evidence.append("labels unique")
    if prof0["bold"] and prof0["bold"] * 2 >= nonempty:
        score += 10
        evidence.append("bold header row")
    if r0 <= 3:
        score += 8
        evidence.append("near sheet top")

    merged_here = [m for m in merged_rows.get(r0, []) if m["min_row"] == r0]
    if merged_here:
        score += 8
        evidence.append("merged cells in candidate row")

    long_label = any(isinstance(v, str) and len(v) > 120 for v in raws)
    if long_label:
        score -= 20
        evidence.append("long text (>120 chars) suggests content, not header")
    if prof0["formula"]:
        score -= 25
        evidence.append("formulas in candidate row")

    table_name = table_rows.get(r0)
    floor = 0
    if table_name:
        score += 15
        floor = 60
        evidence.append(f"table context: {table_name}")

    # blank cells inside the header span (not covered by a merge)
    blank_cells, need_inference = [], False
    merged_cells = set()
    for r in span:
        for m in merged_rows.get(r, []):
            for mc in range(m["min_col"], m["max_col"] + 1):
                merged_cells.add((r, mc))
    for r in span:
        values = row_values(r)
        for c in range(1, width + 1):
            if c not in values or values.get(c) is None:
                if (r, c) not in merged_cells:
                    blank_cells.append(f"{get_column_letter(c)}{r}")
    if blank_cells:
        need_inference = True
        score -= 15
        shown = ", ".join(blank_cells[:5])
        if len(blank_cells) > 5:
            shown += f", ... (+{len(blank_cells) - 5})"
        evidence.append(f"blank header cell(s): {shown}")

    return max(score, floor), blank_cells, need_inference


def _is_title_prefix_row(prof, width):
    """A sparse text-only line glued above a header (no blank separator)."""
    if prof["nonempty"] == 0 or prof["nonempty"] > 2 or prof["formula"]:
        return False
    if prof["text"] != prof["nonempty"]:
        return False
    if prof["nonempty"] * 2 >= width:
        return False
    first = prof.get("first_value")
    if first is None:
        return False
    return bool(prof["bold"] >= 1 or str(first).isupper())


def detect_headers(profiles_by_row, bands, first_title, row_values,
                   merged_rows, table_meta, result, sheet_name):
    """Detect header blocks (single-row / multi-row / merged / repeated)."""
    headers = []
    table_rows = {}
    for t in table_meta:
        table_rows[t["header_row"]] = t["name"]

    for index, band in enumerate(bands):
        if index == 0 and first_title:
            continue
        r0, r1 = band["start"], band["end"]
        width, min_col, max_col = band_metrics(profiles_by_row, band)
        if width < 1:
            continue

        # leading sparse text-only lines glued above the header (no blank
        # separator) are skipped for scoring; they become a title region
        band_start = r0
        while (r0 < r1 and r0 not in table_rows and
               _is_title_prefix_row(profiles_by_row[r0], width)):
            nxt = profiles_by_row.get(r0 + 1)
            if nxt is None or nxt["nonempty"] == 0:
                break
            r0 += 1
        skipped_prefix = r0 > band_start

        span = [r0]
        prof0 = profiles_by_row[r0]
        below1 = profiles_by_row.get(r0 + 1)
        below2 = profiles_by_row.get(r0 + 2)
        merged0 = [m for m in merged_rows.get(r0, []) if m["min_row"] == r0]
        wide_merge = [m for m in merged0
                      if (m["max_col"] - m["min_col"]) >= 1]

        # multi-row span: the row below is still all-text and fills the band
        if below1 and r0 + 1 <= r1:
            b1 = below1
            b1_texty = (b1["nonempty"] > 0 and
                        b1["text"] == b1["nonempty"] and
                        b1["formula"] == 0)
            b2 = below2 if (below2 and r0 + 2 <= r1) else None
            b2_dataish = bool(b2 and (b2["num"] + b2["datish"] +
                                      b2["bool"] + b2["formula"]))
            if b1_texty and (b2_dataish or wide_merge):
                span = [r0, r0 + 1]

        evidence = []
        score, blank_cells, need_inference = _header_score_rows(
            profiles_by_row, band, span, row_values, merged_rows, table_rows,
            width, result, sheet_name, evidence)
        if skipped_prefix:
            evidence.append(
                f"leading title line(s) in rows {band_start}-{r0 - 1} "
                "skipped above the header")

        # repeated-header signal: compare with previously confirmed headers
        repeat_of = None
        if headers:
            my = _header_signature(row_values, span, width)
            for prev in headers:
                shared = total = 0
                for c in range(1, width + 1):
                    a, b = my.get(c), prev["_signature"].get(c)
                    if a is None or b is None:
                        continue
                    total += 1
                    if a == b:
                        shared += 1
                if total and shared / total >= 0.6:
                    repeat_of = prev
                    break
        if repeat_of is not None:
            score += 15
            evidence.append(f"matches earlier header at row "
                            f"{repeat_of['rows'][0]}")

        # a real container fragment may still fail the score; if we have
        # enough evidence of textiness + fill, keep it but LOW (never drop)
        if score < HEADER_SCORE_MIN:
            result.diagnose(
                f"{sheet_name}: band rows {r0}-{r1} scored {score} "
                "(below header threshold); treated as data.",
                sheet=sheet_name, band=[r0, r1], score=score)
            continue

        if blank_cells:
            shown = ", ".join(blank_cells[:5])
            if len(blank_cells) > 5:
                shown += f", ... (+{len(blank_cells) - 5})"
            result.warn("HEADER_BLANK_CELLS",
                        f"{sheet_name}: header band at row {span[0]} has "
                        f"blank cells ({shown}); inference required.",
                        sheet=sheet_name, row=span[0])

        # merged group rows only for the wide-merge multi-row case
        merged_spans = []
        groups = []
        for m in wide_merge:
            merged_spans.append(m["range"])
            top_values = row_values(m["min_row"])
            text = top_values.get(m["min_col"])
            if text is not None:
                groups.append({"range": m["range"], "text": str(text),
                               "normalized": normalize_header(text)})

        labels = []
        for c in range(1, width + 1):
            raw_parts = []
            group = None
            for r in span:
                v = row_values(r).get(c)
                if v is not None:
                    raw_parts.append(str(v))
            for g in groups:
                gm = [m for m in merged0 if m["range"] == g["range"]][0]
                if gm["min_col"] <= c <= gm["max_col"]:
                    group = g["text"]
            raw = " | ".join(raw_parts) if raw_parts else None
            labels.append({
                "column": get_column_letter(c),
                "raw": raw,
                "normalized": normalize_header(raw),
                "group": group,
            })

        score_final, level = _conf(score)
        entry = {
            "range": range_for_rows(profiles_by_row, span[0], span[-1]),
            "rows": list(span),
            "multi_row": len(span) > 1,
            "merged": bool(merged_spans),
            "merged_spans": merged_spans,
            "groups": groups,
            "blank_cells": blank_cells,
            "needs_inference": need_inference,
            "table": table_rows.get(r0),
            "labels": labels,
            "repeats_at_rows": [],
            "confidence_score": score_final,
            "confidence_level": level,
            "evidence": evidence,
            "_signature": _header_signature(row_values, span, width),
        }
        if repeat_of is not None:
            repeat_of["repeats_at_rows"].append(r0)
            entry["repeat_of_row"] = repeat_of["rows"][0]
            entry["evidence"].append(f"repeated header (labels match row "
                                     f"{repeat_of['rows'][0]})")
        headers.append(entry)

    for entry in headers:
        entry.pop("_signature", None)
    return headers


def _header_signature(row_values, span, width):
    """Normalized per-column label map used for repeated-header matching."""
    sig = {}
    for c in range(1, width + 1):
        parts = []
        for r in span:
            v = row_values(r).get(c)
            if v is not None:
                parts.append(str(v))
        if parts:
            sig[c] = normalize_header(" | ".join(parts))
    return sig

# ---------------------------------------------------------------------------
# classification of individual rows (regions)
# ---------------------------------------------------------------------------

def _is_section_row(prof):
    """A single-cell text row inside a band (chapter/heading-like)."""
    if prof["nonempty"] == 0 or prof["nonempty"] > 2:
        return False
    if prof["formula"] or prof["num"] or prof["datish"] or prof["bool"]:
        return False
    if prof["text"] != prof["nonempty"]:
        return False
    value = prof.get("first_value")
    if value is None:
        return False
    text = str(value)
    if not (1 <= len(text) <= 60):
        return False
    has_letter = any(ch.isalpha() for ch in text)
    upper = has_letter and text == text.upper()
    if prof["bold"] >= 1:
        return True
    if not upper:
        return False
    has_digit = any(ch.isdigit() for ch in text)
    has_space = " " in text
    if has_digit and not has_space:
        return False
    if has_space:
        return len(text) >= 5
    return len(text) >= 4 and text.isalpha()


def _summary_kind(prof):
    """'total'/'subtotal' only with an exact keyword or formula proof."""
    norm = normalize_header(prof.get("first_value"))
    if not norm:
        return None
    kind = classify_summary_label(norm)
    if not kind:
        return None
    exact = norm in TOTAL_WORDS or norm in SUBTOTAL_WORDS
    if not exact and prof["formula"] == 0:
        return None
    return kind


# ---------------------------------------------------------------------------
# region detection
# ---------------------------------------------------------------------------

def detect_regions(profiles_by_row, bands, first_title, headers, result,
                   sheet_name):
    """Deterministic heuristic regions: title/header/subheader/data/...
    Every band is fully covered; nothing is silently dropped."""
    regions = []
    headers_by_start = {h["rows"][0]: h for h in headers}
    last_band_index = len(bands) - 1

    for index, band in enumerate(bands):
        r0, r1 = band["start"], band["end"]
        if index == 0 and first_title:
            regions.append(_region(
                "title", r0, r1, profiles_by_row, 65,
                ["sparse block before the first data band (title-like shape)"]))
            continue

        header = headers_by_start.get(r0)
        if header is None:
            for offset in (1, 2):
                candidate = headers_by_start.get(r0 + offset)
                if candidate is None:
                    continue
                prefix = range(r0, r0 + offset)
                if all(profiles_by_row[r]["formula"] == 0 and
                       profiles_by_row[r]["nonempty"] <= 2
                       for r in prefix):
                    regions.append(_region(
                        "title", r0, r0 + offset - 1, profiles_by_row, 65,
                        ["sparse block immediately before the header row"]))
                    header = candidate
                break
        body_start = r0
        if header:
            h_rows = header["rows"]
            regions.append(_region(
                "header", h_rows[0], h_rows[-1], profiles_by_row,
                header["confidence_score"],
                [f"header confidence {header['confidence_level']}"] +
                header["evidence"][:3]))
            if header["multi_row"]:
                regions.append(_region(
                    "subheader", h_rows[1], h_rows[-1], profiles_by_row,
                    max(60, header["confidence_score"] - 10),
                    ["continuation row of the multi-row header"]))
            body_start = h_rows[-1] + 1

        if body_start > r1:
            continue

        # last band, purely textual, no header -> footer candidate
        if (index == last_band_index and not header and
                len(bands) >= 2 and (r1 - r0 + 1) <= 3):
            body_profs = [profiles_by_row[r] for r in range(r0, r1 + 1)]
            if all(p["text"] == p["nonempty"] and p["formula"] == 0 and
                   p["nonempty"] > 0 for p in body_profs):
                regions.append(_region(
                    "footer", r0, r1, profiles_by_row, 65,
                    ["text-only block after the last data band"]))
                continue

        data_run, formula_run = [], []

        def flush():
            if formula_run:
                if len(formula_run) >= 2:
                    regions.append(_region(
                        "formula_region", formula_run[0], formula_run[-1],
                        profiles_by_row, 70,
                        [f"{len(formula_run)} contiguous rows where "
                         "formulas dominate the non-empty cells"]))
                else:
                    data_run.extend(formula_run)
                formula_run.clear()
            if data_run:
                regions.append(_region(
                    "data", data_run[0], data_run[-1], profiles_by_row, 70,
                    ["contiguous value rows inside the band"]))
                data_run.clear()

        r = body_start
        while r <= r1:
            prof = profiles_by_row[r]
            kind = _summary_kind(prof)
            if kind:
                flush()
                norm = normalize_header(prof.get("first_value"))
                exact = norm in TOTAL_WORDS or norm in SUBTOTAL_WORDS
                is_total = norm in TOTAL_WORDS
                score = 75 if exact else 65
                regions.append(_region(
                    "total" if is_total else "subtotal",
                    r, r, profiles_by_row, score,
                    [f"summary row labelled '{prof.get('first_value')}'" +
                     (" (exact keyword)" if exact else
                      " (keyword + formula evidence)")]))
                r += 1
                continue
            if _is_section_row(prof):
                flush()
                label = str(prof.get("first_value"))
                score = 65 if prof["bold"] >= 1 else 45
                regions.append(_region(
                    "section_header", r, r, profiles_by_row, score,
                    [f"single text row ('{truncate_text(label, 30)}')" +
                     (" (bold)" if prof["bold"] >= 1 else
                      " (upper-case text)")]))
                r += 1
                continue
            if looks_like_notes(normalize_header(prof.get("first_value"))):
                flush()
                regions.append(_region(
                    "notes", r, r, profiles_by_row, 65,
                    ["left-most cell starts with a notes keyword"]))
                r += 1
                continue
            if (prof["formula"] >= 1 and prof["nonempty"] >= 2 and
                    prof["formula"] * 3 >= prof["nonempty"]):
                formula_run.append(r)
                r += 1
                continue
            if formula_run:
                flush()
            data_run.append(r)
            r += 1
        flush()

    # styled-but-empty row runs -> input_region (declared into the map)
    run_start = None
    for prof in profiles_by_row.values():
        if prof["nonempty"] == 0 and prof["styled"] > 0:
            if run_start is None:
                run_start = prof["row"]
            run_end = prof["row"]
        else:
            if run_start is not None:
                regions.append(_styled_input_region(profiles_by_row,
                                                    run_start, run_end))
                run_start = None
    if run_start is not None:
        regions.append(_styled_input_region(profiles_by_row, run_start,
                                            run_end))

    regions.sort(key=lambda e: (e["rows"][0], e["rows"][1], e["type"]))
    return regions


def _styled_input_region(profiles_by_row, r_start, r_end):
    styled_cells = 0
    min_col = max_col = None
    for r in range(r_start, r_end + 1):
        prof = profiles_by_row[r]
        styled_cells += prof["styled"]
        if prof["styled_min"] is not None:
            min_col = (prof["styled_min"] if min_col is None
                       else min(min_col, prof["styled_min"]))
            max_col = (prof["styled_max"] if max_col is None
                       else max(max_col, prof["styled_max"]))
    entry = _region("input_region", r_start, r_end, profiles_by_row, 70,
                    [f"{styled_cells} styled empty cells (fill/border) "
                     "without values"])
    if min_col is not None:
        entry["range"] = (f"{get_column_letter(min_col)}{r_start}:"
                          f"{get_column_letter(max_col)}{r_end}")
    return entry


# ---------------------------------------------------------------------------
# repeated structure detection  (detection only -- no template matching)
# ---------------------------------------------------------------------------

def _sig_sparse(sig):
    if not sig:
        return True
    dots = sig.count(".")
    return dots * 5 > len(sig) * 3


def detect_repeated_structures(profiles_by_row, bands, headers, sheet_name):
    """Repeated header groups, repeated row blocks and periodic patterns.

    The confidence here is DETECTION confidence (structure was observed
    repeating), never a template-matching confidence.
    """
    out = []

    for header in headers:
        if header["repeats_at_rows"]:
            rows = [header["rows"][0]] + list(header["repeats_at_rows"])
            signature = " | ".join(
                str(lab["normalized"]) for lab in header["labels"]
                if lab["normalized"])
            out.append({
                "type": "repeated_header",
                "signature": truncate_text(signature, 120),
                "rows": rows,
                "occurrences": len(rows),
                "evidence": [f"header labels match across rows {rows}"],
                "confidence_score": 80,
                "confidence_level": "HIGH",
            })

    reported_periodic_rows = set()
    for band in bands:
        body = [r for r in range(band["start"], band["end"] + 1)
                if profiles_by_row[r]["nonempty"] > 0]
        if len(body) < 3:
            continue
        width = max((profiles_by_row[r]["last_col"] or 0) for r in body)
        sigs = {r: _sig_pad(profiles_by_row[r]["sig"], width) for r in body}

        # 1) identical adjacent runs, occurring more than once
        runs = []
        i = 0
        while i < len(body):
            j = i
            while (j + 1 < len(body) and
                   sigs[body[j + 1]] == sigs[body[i]]):
                j += 1
            runs.append((body[i], body[j]))
            i = j + 1
        by_sig = {}
        for a, b in runs:
            by_sig.setdefault(sigs[a], []).append((a, b))
        big_run_rows = set()
        for sig, occurrences in sorted(by_sig.items()):
            if len(occurrences) < 2:
                continue
            for a, b in occurrences:
                big_run_rows.update(range(a, b + 1))
            sparse = _sig_sparse(sig)
            score, level = ((45, "LOW") if sparse else
                            (80, "HIGH") if len(occurrences) >= 3 else
                            (65, "MEDIUM"))
            out.append({
                "type": "repeated_row_block",
                "signature": sig,
                "runs": [{"rows": [a, b],
                          "range": range_for_rows(profiles_by_row, a, b)}
                         for a, b in occurrences],
                "occurrences": len(occurrences),
                "row_count_per_run": occurrences[0][1] - occurrences[0][0] + 1,
                "evidence": [f"rows {a}-{b}" for a, b in occurrences],
                "confidence_score": score,
                "confidence_level": level,
            })

        # 2) periodic patterns (smallest period wins; no duplicate rows)
        covered_periods = set()
        for period in range(2, 7):
            spans = []
            i = 0
            while i < len(body):
                j = i + period
                while (j < len(body) and j - period >= i and
                       sigs[body[j]] == sigs[body[j - period]]):
                    j += 1
                span_len = j - i
                if span_len >= 2 * period and span_len >= 4:
                    spans.append((i, j - 1))
                    i = j
                else:
                    i += 1
            if not spans:
                continue
            a, b = max(spans, key=lambda s: s[1] - s[0])
            span_rows = set(body[a:b + 1])
            if len({sigs[r] for r in span_rows}) == 1:
                continue
            if span_rows <= reported_periodic_rows:
                continue
            if len(span_rows & big_run_rows) >= 0.8 * len(span_rows):
                continue
            reported_periodic_rows |= span_rows
            occurrences = (b - a + 1) // period
            score = 80 if occurrences >= 3 else 65
            out.append({
                "type": "repeated_periodic_pattern",
                "period": period,
                "rows": [body[a], body[b]],
                "unit_signatures": [sigs[body[a + k]]
                                    for k in range(period)],
                "occurrences": occurrences,
                "evidence": [f"row signatures repeat every {period} rows "
                             f"from row {body[a]} to row {body[b]}"],
                "confidence_score": score,
                "confidence_level": "HIGH" if score >= 80 else "MEDIUM",
            })

    def sort_key(entry):
        if "rows" in entry:
            return entry["rows"][0]
        return entry["runs"][0]["rows"][0]

    out.sort(key=lambda e: (sort_key(e), e["type"]))
    return out

# ---------------------------------------------------------------------------
# Pass A  --  read_only, data_only=False: values, formula text, profiles
# ---------------------------------------------------------------------------

def _new_profile(row):
    return {
        "row": row,
        "nonempty": 0, "formula": 0, "num": 0, "text": 0, "datish": 0,
        "bool": 0,
        "styled": 0, "styled_min": None, "styled_max": None,
        "first_col": None, "last_col": None, "first_value": None,
        "values": {},
        "bold": 0, "filled": 0,
        "sig": "",
    }


def scan_pass_a(ws_ro, sample_cap, result, sheet_name):
    """Stream one read-only worksheet exactly once.

    Returns (profiles_by_row, column_stats, formulas, failures, caps):
      profiles_by_row: row -> profile dict (gap rows included)
      column_stats:    col index -> ColumnStats
      formulas:        [(row, col, text)] in scan order
      failures:        {(property, exc_type)} style read failures
      caps:            {"declared_rows","declared_cols","scanned_rows",
                        "capped": bool}
    """
    declared_rows = ws_ro.max_row or 0
    declared_cols = ws_ro.max_column or 0
    profiles_by_row = {}
    column_stats = {}
    formulas = []
    failures = set()

    capped = declared_rows > SCAN_ROW_CAP
    limit = min(declared_rows, SCAN_ROW_CAP) if declared_rows else 0
    expected_row = 1
    scanned_rows = 0
    values_stored = 0
    values_partial = False

    for row_cells in ws_ro.iter_rows():
        actual = None
        for cell in row_cells:
            coord = cell_coordinate(cell)
            if coord:
                actual = coordinate_to_tuple(coord)[0]
                break
        r = actual if actual is not None else expected_row
        expected_row = r + 1
        if capped and r > limit:
            break
        scanned_rows += 1

        prof = _new_profile(r)
        if values_stored >= ROW_VALUE_BUDGET:
            prof["values"] = None
            values_partial = True
        sig_letters = {}
        max_col_seen = 0
        for cell in row_cells:
            coord = cell_coordinate(cell)
            if coord is None:
                continue
            _, c_idx = coordinate_to_tuple(coord)
            max_col_seen = max(max_col_seen, c_idx)
            value = cell_value(cell)
            numfmt, bold, italic, filled = read_cell_style(cell, failures)
            kind = classify_scalar(value)
            sig_letters[c_idx] = ROW_SIG_TYPE_CODES.get(kind, "S")

            if kind != "blank":
                prof["nonempty"] += 1
                if kind == "formula":
                    prof["formula"] += 1
                    formulas.append((r, c_idx, value))
                elif kind in ("integer", "decimal"):
                    prof["num"] += 1
                elif kind == "string":
                    prof["text"] += 1
                elif kind in ("date", "datetime"):
                    prof["datish"] += 1
                elif kind == "boolean":
                    prof["bool"] += 1
                if prof["first_col"] is None:
                    prof["first_col"] = c_idx
                    prof["first_value"] = value
                prof["last_col"] = c_idx
                if prof["values"] is not None:
                    if values_stored < ROW_VALUE_BUDGET:
                        prof["values"][c_idx] = value
                        values_stored += 1
                    else:
                        prof["values"] = None
                        values_partial = True
            else:
                styled = bool((numfmt not in (None, "General")) or bold or
                              italic or filled)
                if styled:
                    prof["styled"] += 1
                    prof["styled_min"] = (c_idx if prof["styled_min"] is None
                                          else min(prof["styled_min"], c_idx))
                    prof["styled_max"] = (c_idx if prof["styled_max"] is None
                                          else max(prof["styled_max"], c_idx))

            if bold:
                prof["bold"] += 1
            if filled:
                prof["filled"] += 1
            stats = column_stats.get(c_idx)
            if stats is None:
                stats = ColumnStats(c_idx, sample_cap)
                column_stats[c_idx] = stats
            stats.feed(value, (numfmt, bold, italic, filled))

        if max_col_seen:
            prof["sig"] = "".join(sig_letters.get(c, ".")
                                  for c in range(1, max_col_seen + 1))
        profiles_by_row[r] = prof

    if values_partial:
        result.warn(
            "ROW_VALUES_PARTIAL",
            f"{sheet_name}: per-row label values stored for the "
            f"first {ROW_VALUE_BUDGET} non-empty cells; later rows "
            "fall back to first-value-only label evidence (never "
            "silent).",
            sheet=sheet_name, budget=ROW_VALUE_BUDGET)

    caps = {
        "declared_rows": declared_rows,
        "declared_cols": declared_cols,
        "scanned_rows": scanned_rows,
        "capped": capped,
    }
    if capped:
        result.warn(
            "SCAN_CAPPED",
            f"{sheet_name}: sheet declares {declared_rows} rows; the "
            f"analysis was capped at {SCAN_ROW_CAP} rows (never silent).",
            sheet=sheet_name, declared_rows=declared_rows, cap=SCAN_ROW_CAP)
    return profiles_by_row, column_stats, formulas, failures, caps


# ---------------------------------------------------------------------------
# Pass B  --  read_only, data_only=True: cached values (availability only)
# ---------------------------------------------------------------------------

def scan_pass_b(handle_b, formulas_by_sheet):
    """Cached values for the formula cells found in Pass A, nothing else.

    Cached non-None values mark availability; they are prior-save
    artifacts and never reported as a recalculation.
    """
    cached = {}
    for sheet_name, wanted in formulas_by_sheet.items():
        wanted_set = {(r, c) for r, c, _ in wanted}
        if not wanted_set:
            continue
        wanted_rows = {r for r, _ in wanted_set}
        for row_index, row in enumerate(handle_b.iter_values(sheet_name),
                                       start=1):
            if row_index not in wanted_rows:
                continue
            for col_index, value in enumerate(row, start=1):
                if value is not None and (row_index, col_index) in wanted_set:
                    cached[(sheet_name, row_index, col_index)] = value
    return cached


# ---------------------------------------------------------------------------
# formula discovery  (Faz 2A: discovery ONLY -- no pattern engine)
# ---------------------------------------------------------------------------

OPERATOR_ORDER = ("+", "-", "*", "/", "^", "&", "<", ">", "=",
                  "<=", ">=", "<>")


def formula_kind(text, refs):
    """Coarse, deterministic kind label for one formula string."""
    body = text[1:] if text.startswith("=") else text
    masked = STRING_RE.sub('""', body)
    has_func = bool(FUNC_RE.search(masked))
    has_cmp = bool(CMP_RE.search(masked))
    has_concat = "&" in masked
    has_arith = any(ch in ARITH_CHARS for ch in masked)
    if not refs and not has_func:
        return "constant"
    if has_func:
        return "function"
    if has_cmp:
        return "comparison"
    if has_concat:
        return "concat"
    if has_arith:
        if any(e["sheet"] for e in refs):
            return "cross_sheet_arithmetic"
        if any(e["start"]["abs_column"] or e["start"]["abs_row"] or
               (e["end"] and (e["end"]["abs_column"] or e["end"]["abs_row"]))
               for e in refs):
            return "absolute_arithmetic"
        return "relative_arithmetic"
    if refs:
        return "reference"
    return "mixed"


def build_formula_entries(formulas, cached, sheet_name, result,
                          cap=FORMULA_LIST_CAP):
    """Listed formula entries (capped) + full-coverage aggregates."""
    entries = []
    kind_counts = {}
    function_counts = {}
    cross_count = 0
    absolute_count = 0
    cached_available = 0
    for r, c, text in formulas:
        refs = extract_refs(text)
        sheet_refs = sorted({e["sheet"] for e in refs if e["sheet"]})
        has_abs = any(
            e["start"]["abs_column"] or e["start"]["abs_row"] or
            (e["end"] and (e["end"]["abs_column"] or e["end"]["abs_row"]))
            for e in refs)
        masked = STRING_RE.sub('""', text[1:] if text.startswith("=")
                               else text)
        functions = sorted({m.group(1).upper()
                            for m in FUNC_RE.finditer(masked)})
        operators = [op for op in OPERATOR_ORDER if op in masked]
        kind = formula_kind(text, refs)
        kind_counts[kind] = kind_counts.get(kind, 0) + 1
        for name in functions:
            function_counts[name] = function_counts.get(name, 0) + 1
        if sheet_refs:
            cross_count += 1
        if has_abs:
            absolute_count += 1
        key = (sheet_name, r, c)
        available = key in cached
        if available:
            cached_available += 1
        if len(entries) < cap:
            entries.append({
                "cell": f"{get_column_letter(c)}{r}",
                "text": text,
                "kind": kind,
                "functions": functions,
                "operators": operators,
                "ref_count": len(refs),
                "refs": refs,
                "sheet_refs": sheet_refs,
                "has_absolute": has_abs,
                "has_cross_sheet": bool(sheet_refs),
                "cached_value_available": available,
                "cached_value": sample_value(cached[key]) if available
                else None,
            })
    summary = {
        "total": len(formulas),
        "listed": len(entries),
        "truncated": len(formulas) > len(entries),
        "kinds": dict(sorted(kind_counts.items())),
        "functions_top": [
            {"name": name, "count": count}
            for name, count in sorted(function_counts.items(),
                                      key=lambda kv: (-kv[1], kv[0]))[:15]],
        "cross_sheet_count": cross_count,
        "absolute_count": absolute_count,
        "cached_available": cached_available,
    }
    if summary["truncated"]:
        result.warn(
            "FORMULA_LIST_TRUNCATED",
            f"{sheet_name}: formula list capped at {cap} listed entries "
            f"of {summary['total']} formulas; formula_summary aggregates "
            "cover ALL formulas (never silent).",
            sheet=sheet_name, listed=len(entries),
            total=summary["total"], cap=cap)
    return entries, summary


# ---------------------------------------------------------------------------
# Pass S  --  normal load, strictly read-only USAGE: structure extraction
# ---------------------------------------------------------------------------

KNOWN_CF_KINDS = {
    "cellIs", "colorScale", "dataBar", "iconSet", "expression",
    "containsText", "notContainsText", "beginsWith", "endsWith",
    "timePeriod", "top10", "aboveAverage", "duplicateValues",
    "uniqueValues", "containsBlanks", "notContainsBlanks",
    "containsErrors", "notContainsErrors",
}

STRUCTURE_CAP = 1000  # row-height / column-width maps cap (warned)


NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _sheet_target(base, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base, target))


def _read_xml(zipf, part, result, sheet_name=None):
    try:
        blob = zipf.read(part)
    except KeyError:
        result.unsupported_item(
            f"{part}: package part is missing; its information cannot "
            "be reported.", sheet=sheet_name)
        return None
    try:
        return _ET.fromstring(blob)
    except Exception as exc:  # noqa: BLE001
        result.unsupported_item(
            f"{part}: XML could not be parsed ({type(exc).__name__}); "
            "its information cannot be reported.", sheet=sheet_name)
        return None


def read_package_relationships(zipf, part, result, sheet_name=None):
    """rId -> normalized target, from a .rels part (type kept too)."""
    root = _read_xml(zipf, part, result, sheet_name)
    out = {}
    if root is None:
        return out
    base = posixpath.dirname(part)
    base = posixpath.dirname(base)  # 'xl/_rels/x.rels' -> 'xl'
    for rel in root:
        rid = rel.get("Id")
        if not rid:
            continue
        rel_type = rel.get("Type") or ""
        target = rel.get("Target") or ""
        if rel.get("TargetMode") == "External":
            out[rid] = {"type": rel_type, "target": target, "external": True}
            continue
        out[rid] = {"type": rel_type,
                    "target": _sheet_target(base, target),
                    "external": False}
    return out


def read_workbook_raw(zipf, result):
    """Sheets (order + state + part path), names, calcPr, epoch, rels."""
    root = _read_xml(zipf, "xl/workbook.xml", result)
    if root is None:
        raise common.XlsxError(
            "VALIDATION_FAILED", "xl/workbook.xml could not be read.",
            recovery="The file may be damaged; open and re-save it in "
                     "Excel (or LibreOffice) and try again.")
    rels = read_package_relationships(zipf, "xl/_rels/workbook.xml.rels",
                                      result)
    sheets = []
    sheets_el = root.find(NS_MAIN + "sheets")
    if sheets_el is not None:
        for sheet in sheets_el:
            rid = sheet.get(NS_REL_ATTR + "id")
            target = None
            if rid and rid in rels and not rels[rid]["external"]:
                target = rels[rid]["target"]
            sheets.append({
                "name": sheet.get("name"),
                "state": sheet.get("state") or "visible",
                "sheet_id": sheet.get("sheetId"),
                "rid": rid,
                "part": target,
            })
    defined_names = []
    dn_root = root.find(NS_MAIN + "definedNames")
    if dn_root is not None:
        for dn in dn_root:
            local = dn.get("localSheetId")
            defined_names.append({
                "name": dn.get("name"),
                "local_sheet_id": int(local) if local is not None else None,
                "attr_text": (dn.text or "").strip(),
            })
    defined_names.sort(key=lambda e: (str(e["name"]),
                                      e["local_sheet_id"]
                                      if e["local_sheet_id"] is not None
                                      else -1))
    calc_pr = {}
    calc_el = root.find(NS_MAIN + "calcPr")
    if calc_el is not None:
        calc_pr = dict(calc_el.attrib)
    workbook_pr = root.find(NS_MAIN + "workbookPr")
    date1904 = False
    if workbook_pr is not None:
        date1904 = str(workbook_pr.get("date1904", "")).lower() in ("1",
                                                                    "true")
    ext_refs = []
    ext_el = root.find(NS_MAIN + "externalReferences")
    if ext_el is not None:
        for ref in ext_el:
            rid = ref.get(NS_REL_ATTR + "id")
            if rid:
                ext_refs.append(rid)
    return {
        "sheets": sheets,
        "defined_names": defined_names,
        "calc_pr": calc_pr,
        "date1904": date1904,
        "external_refs": ext_refs,
        "rels": rels,
    }


def _range_dict(ref):
    min_col, min_row, max_col, max_row = range_boundaries(ref)
    return {
        "range": f"{get_column_letter(min_col)}{min_row}:"
                 f"{get_column_letter(max_col)}{max_row}",
        "top_left_cell": f"{get_column_letter(min_col)}{min_row}",
        "row_span": max_row - min_row + 1,
        "column_span": max_col - min_col + 1,
        "min_row": min_row, "max_row": max_row,
        "min_col": min_col, "max_col": max_col,
    }


def extract_structure_raw(zipf, sheet_meta, result):
    """Same shape as the former openpyxl structure dict, from raw XML."""
    sheet_name = sheet_meta["name"]
    part = sheet_meta["part"]
    structure = {
        "sheet_state": sheet_meta["state"],
        "merged_cells": [], "hidden": {"rows": [], "columns": []},
        "row_heights": {}, "column_widths": {},
        "freeze_panes": None, "autofilter": None, "tables": [],
        "validations": [], "conditional_formats": [],
        "charts": 0, "images": 0,
        "_dimension_ref": None,
    }
    if not part:
        result.unsupported_item(
            f"{sheet_name}: worksheet part could not be located via the "
            "workbook relationships.", sheet=sheet_name)
        return structure
    root = _read_xml(zipf, part, result, sheet_name)
    if root is None:
        return structure

    dim_el = root.find(NS_MAIN + "dimension")
    if dim_el is not None and dim_el.get("ref"):
        structure["_dimension_ref"] = dim_el.get("ref")

    for merge in root.iter(NS_MAIN + "mergeCell"):
        ref = merge.get("ref")
        if ref:
            try:
                structure["merged_cells"].append(_range_dict(ref))
            except Exception as exc:  # noqa: BLE001
                result.unsupported_item(
                    f"{sheet_name}: merge range '{ref}' could not be "
                    f"parsed ({type(exc).__name__}).", sheet=sheet_name)
    structure["merged_cells"].sort(key=lambda m: (m["min_row"],
                                                  m["min_col"]))

    heights_truncated = False
    widths_truncated = False
    outline_levels = set()
    for row in root.iter(NS_MAIN + "row"):
        index = row.get("r")
        if index is None:
            continue
        outline = row.get("outlineLevel")
        if outline:
            outline_levels.add(int(outline))
        if row.get("hidden") in ("1", "true"):
            structure["hidden"]["rows"].append({
                "index": int(index),
                "outline_level": int(outline or 0),
                "collapsed": row.get("collapsed") in ("1", "true"),
            })
        if row.get("ht") is not None:
            if len(structure["row_heights"]) < STRUCTURE_CAP:
                structure["row_heights"][str(index)] = float(row.get("ht"))
            else:
                heights_truncated = True
    structure["hidden"]["rows"].sort(key=lambda r: r["index"])

    for col in root.iter(NS_MAIN + "col"):
        try:
            c_min = int(col.get("min"))
        except (TypeError, ValueError):
            continue
        outline = col.get("outlineLevel")
        if outline:
            outline_levels.add(int(outline))
        letter = get_column_letter(c_min)
        if col.get("hidden") in ("1", "true"):
            structure["hidden"]["columns"].append({
                "index": letter,
                "outline_level": int(outline or 0),
                "collapsed": col.get("collapsed") in ("1", "true"),
            })
        if col.get("width") is not None:
            if len(structure["column_widths"]) < STRUCTURE_CAP:
                structure["column_widths"][letter] = float(col.get("width"))
            else:
                widths_truncated = True
    structure["hidden"]["columns"].sort(key=lambda c: c["index"])

    if heights_truncated:
        result.warn("STRUCTURE_CAP_REACHED",
                    f"{sheet_name}: row_heights truncated at "
                    f"{STRUCTURE_CAP} entries (never silent).",
                    sheet=sheet_name, kind="row_heights")
    if widths_truncated:
        result.warn("STRUCTURE_CAP_REACHED",
                    f"{sheet_name}: column_widths truncated at "
                    f"{STRUCTURE_CAP} entries (never silent).",
                    sheet=sheet_name, kind="column_widths")

    for pane in root.iter(NS_MAIN + "pane"):
        state = pane.get("state")
        if state in ("frozen", "frozenSplit") and pane.get("topLeftCell"):
            structure["freeze_panes"] = pane.get("topLeftCell")
            break

    af = root.find(NS_MAIN + "autoFilter")
    if af is not None:
        structure["autofilter"] = af.get("ref")

    for dv in root.iter(NS_MAIN + "dataValidation"):
        f1 = f2 = None
        for child in dv:
            if child.tag == NS_MAIN + "formula1":
                f1 = child.text
            elif child.tag == NS_MAIN + "formula2":
                f2 = child.text
        structure["validations"].append({
            "range": dv.get("sqref"),
            "type": dv.get("type"),
            "operator": dv.get("operator"),
            "formula1": common.jsonable(f1),
            "formula2": common.jsonable(f2),
            "allow_blank": str(dv.get("allowBlank", "")).lower()
                           in ("1", "true"),
        })
    structure["validations"].sort(key=lambda v: (v["range"] or "",
                                                 str(v["type"])))

    for cf in root.iter(NS_MAIN + "conditionalFormatting"):
        sqref = cf.get("sqref")
        for rule in cf:
            if rule.tag != NS_MAIN + "cfRule":
                continue
            kind = rule.get("type") or ""
            entry = {"range": sqref, "kind": kind}
            if rule.get("operator"):
                entry["operator"] = rule.get("operator")
            structure["conditional_formats"].append(entry)
            if kind not in KNOWN_CF_KINDS:
                result.unsupported_item(
                    f"{sheet_name}: conditional-format rule type '{kind}' "
                    "is not recognised by Phase 2A and is reported, not "
                    "silently skipped.", sheet=sheet_name, range=sqref)
    structure["conditional_formats"].sort(key=lambda e: (e["range"] or "",
                                                         str(e["kind"])))

    if outline_levels:
        result.diagnose(
            f"{sheet_name}: row/column grouping present (outline levels "
            f"{sorted(outline_levels)}); hidden flags and outline levels "
            "are reported separately and never equated.",
            sheet=sheet_name)

    # tables + chart/image counts via the sheet's own relationships
    sheet_dir = posixpath.dirname(part)
    rels_part = posixpath.join(sheet_dir, "_rels",
                               posixpath.basename(part) + ".rels")
    rels = read_package_relationships(zipf, rels_part, result, sheet_name)
    for rel in rels.values():
        rtype = rel["type"]
        if rtype.endswith("/chart"):
            structure["charts"] += 1
        elif rtype.endswith("/image"):
            structure["images"] += 1
        elif rtype.endswith("/drawing") and not rel["external"]:
            dpart = rel["target"]
            drels = posixpath.join(posixpath.dirname(dpart), "_rels",
                                   posixpath.basename(dpart) + ".rels")
            if drels in zipf.namelist():
                drawing_rels = read_package_relationships(
                    zipf, drels, result, sheet_name)
                for drel in drawing_rels.values():
                    dtype = drel["type"]
                    if dtype.endswith("/chart"):
                        structure["charts"] += 1
                    elif dtype.endswith("/image"):
                        structure["images"] += 1
        elif rtype.endswith("/table") and not rel["external"]:
            t_root = _read_xml(zipf, rel["target"], result, sheet_name)
            if t_root is None:
                continue
            ref = t_root.get("ref")
            if not ref:
                continue
            try:
                bounds = range_boundaries(ref)
            except Exception as exc:  # noqa: BLE001
                result.unsupported_item(
                    f"{sheet_name}: table '{t_root.get('displayName')}' "
                    f"has an unreadable ref ({type(exc).__name__}).",
                    sheet=sheet_name)
                continue
            min_col, min_row, max_col, max_row = bounds
            header_count = t_root.get("headerRowCount")
            header_count = 1 if header_count is None else int(header_count)
            totals_count = int(t_root.get("totalsRowCount") or 0)
            columns = []
            cols_el = t_root.find(NS_MAIN + "tableColumns")
            if cols_el is not None:
                for idx, tc in enumerate(cols_el):
                    columns.append({
                        "column": get_column_letter(
                            min_col + idx),
                        "name": common.jsonable(tc.get("name")),
                    })
            style = None
            style_el = t_root.find(NS_MAIN + "tableStyleInfo")
            if style_el is not None:
                style = style_el.get("name")
            structure["tables"].append({
                "name": t_root.get("displayName"),
                "ref": f"{get_column_letter(min_col)}{min_row}:"
                       f"{get_column_letter(max_col)}{max_row}",
                "sheet": sheet_name,
                "header_row": min_row,
                "min_col": min_col,
                "max_col": max_col,
                "columns": columns,
                "row_count": (max_row - min_row + 1 - header_count -
                              totals_count),
                "totals_row": bool(totals_count),
                "totals_row_count": totals_count,
                "style": style,
            })
    structure["tables"].sort(key=lambda t: str(t["name"]))
    return structure


# ---------------------------------------------------------------------------
# sheet assembly
# ---------------------------------------------------------------------------

def make_row_values(profiles_by_row):
    """Row -> {column: raw value} from Pass A profiles (no re-reads)."""
    cache = {}

    def row_values(row):
        if row in cache:
            return cache[row]
        prof = profiles_by_row.get(row)
        if prof is None:
            out = {}
        elif prof.get("values") is not None:
            out = prof["values"]
        else:
            out = {}
            if prof["first_col"] is not None:
                out = {prof["first_col"]: prof["first_value"]}
        cache[row] = out
        return out

    return row_values


def _column_region(profiles_by_row, regions, col_index):
    counts = {}
    order = {}
    for pos, region in enumerate(regions):
        a, b = region["rows"]
        count = 0
        for r in range(a, b + 1):
            prof = profiles_by_row.get(r)
            if (prof and prof["first_col"] is not None and
                    prof["first_col"] <= col_index <= prof["last_col"]):
                count += 1
        if count:
            counts[region["type"]] = counts.get(region["type"], 0) + count
            order.setdefault(region["type"], pos)
    if not counts:
        return None
    best = sorted(counts.items(), key=lambda kv: (-kv[1], order[kv[0]]))
    return best[0][0]


def _flag_low_confidence(sheet_name, headers, regions, repeated, result):
    for entry in headers:
        if entry["confidence_level"] == "LOW":
            result.warn(
                "LOW_CONFIDENCE_DETECTION",
                f"{sheet_name}: header at rows {entry['rows']} flagged LOW "
                "confidence; review the evidence before relying on it.",
                sheet=sheet_name, kind="header", rows=entry["rows"])
    for entry in regions:
        if entry["confidence_level"] == "LOW":
            result.warn(
                "LOW_CONFIDENCE_DETECTION",
                f"{sheet_name}: region '{entry['type']}' at {entry['range']}"
                " flagged LOW confidence; review the evidence.",
                sheet=sheet_name, kind=entry["type"], range=entry["range"])
    for entry in repeated:
        if entry["confidence_level"] == "LOW":
            result.warn(
                "LOW_CONFIDENCE_DETECTION",
                f"{sheet_name}: repeated structure '{entry['type']}' flagged"
                " LOW confidence; review the evidence.",
                sheet=sheet_name, kind=entry["type"])


def assemble_sheet(sheet_name, index, data, structure, cached, result,
                   samples, formula_cap):
    profiles_by_row = data["profiles"]
    column_stats = data["columns"]
    caps = data["caps"]

    merged_rows = {}
    for m in structure["merged_cells"]:
        for r in range(m["min_row"], m["max_row"] + 1):
            merged_rows.setdefault(r, []).append(m)

    ordered = [profiles_by_row[r] for r in sorted(profiles_by_row)]
    bands = compute_bands(ordered)
    first_title, title_evidence = _first_band_is_title(
        profiles_by_row, bands, merged_rows)
    row_values = make_row_values(profiles_by_row)

    headers = detect_headers(profiles_by_row, bands, first_title, row_values,
                             merged_rows, structure["tables"], result,
                             sheet_name)
    regions = detect_regions(profiles_by_row, bands, first_title, headers,
                             result, sheet_name)
    repeated = detect_repeated_structures(profiles_by_row, bands, headers,
                                          sheet_name)
    _flag_low_confidence(sheet_name, headers, regions, repeated, result)

    # dimensions / used range / inflation
    true_rows = [r for r, p in profiles_by_row.items() if p["nonempty"] > 0]
    true_max_row = max(true_rows) if true_rows else None
    true_min_row = min(true_rows) if true_rows else None
    true_cols = [p["last_col"] for p in profiles_by_row.values()
                 if p["last_col"] is not None]
    true_max_col = max(true_cols) if true_cols else None
    true_min_cols = [p["first_col"] for p in profiles_by_row.values()
                     if p["first_col"] is not None]
    true_min_col = min(true_min_cols) if true_min_cols else None
    used_range = None
    if true_max_row is not None:
        used_range = (f"{get_column_letter(true_min_col)}"
                      f"{true_min_row}:"
                      f"{get_column_letter(true_max_col)}{true_max_row}")
    declared_rows = caps["declared_rows"]
    declared_cols = caps["declared_cols"]
    inflation_rows = max(0, declared_rows - (true_max_row or 0))
    inflation_cols = max(0, declared_cols - (true_max_col or 0))
    styled_empty_rows = sum(1 for p in profiles_by_row.values()
                            if p["nonempty"] == 0 and p["styled"] > 0)
    dimensions = {
        "declared_max_row": declared_rows,
        "declared_max_column": declared_cols,
        "true_max_row": true_max_row,
        "true_max_column": true_max_col,
        "true_min_row": true_min_row,
        "true_min_column": true_min_col,
        "used_range": used_range,
        "inflation_rows": inflation_rows,
        "inflation_columns": inflation_cols,
        "styled_empty_rows": styled_empty_rows,
    }
    if inflation_rows or inflation_cols:
        hint = "styling-only cells"
        if structure["merged_cells"]:
            hint += " and/or merged regions"
        result.diagnose(
            f"{sheet_name}: declared dimension ({declared_rows}r x "
            f"{declared_cols}c) exceeds the true used range "
            f"({used_range}); inflation rows +{inflation_rows}, "
            f"columns +{inflation_cols} ({hint}). The declared dimension "
            "and the computed used range are NOT the same thing.",
            sheet=sheet_name)

    # columns  (header / normalized_header / region attached)
    columns = []
    primary_header = headers[0] if headers else None
    header_labels = {}
    if primary_header:
        for lab in primary_header["labels"]:
            header_labels[lab["column"]] = lab
    header_row_set = set()
    for h in headers:
        header_row_set.update(range(h["rows"][0], h["rows"][-1] + 1))
    letter_kind = {"t": "string", "i": "integer", "n": "decimal",
                   "d": "datetime", "b": "boolean", "q": "formula",
                   "x": "string"}
    for c_idx in sorted(column_stats):
        col = column_stats[c_idx].to_map()
        col["data_type_full_column"] = col["data_type"]
        rebuilt = {}
        for r, prof in profiles_by_row.items():
            if r in header_row_set:
                continue
            sig = prof["sig"]
            letter = sig[c_idx - 1] if c_idx - 1 < len(sig) else "."
            if letter == ".":
                continue
            kind = letter_kind.get(letter)
            if kind:
                rebuilt[kind] = rebuilt.get(kind, 0) + 1
        if rebuilt:
            col["data_type"] = rollup_column_type(rebuilt)
        label = header_labels.get(col["letter"])
        col["header"] = label["raw"] if label else None
        col["normalized_header"] = label["normalized"] if label else None
        col["region"] = _column_region(profiles_by_row, regions, c_idx)
        columns.append(col)

    formulas, formula_summary = build_formula_entries(
        data["formulas"], cached, sheet_name, result, cap=formula_cap)

    return {
        "name": sheet_name,
        "index": index,
        "visibility": structure["sheet_state"],
        "dimensions": dimensions,
        "regions": regions,
        "headers": headers,
        "columns": columns,
        "tables": structure["tables"],
        "merged_cells": structure["merged_cells"],
        "hidden": structure["hidden"],
        "row_heights": structure["row_heights"],
        "column_widths": structure["column_widths"],
        "freeze_panes": structure["freeze_panes"],
        "autofilter": structure["autofilter"],
        "validations": structure["validations"],
        "conditional_formats": structure["conditional_formats"],
        "formulas": formulas,
        "formula_summary": formula_summary,
        "repeated_structures": repeated,
        "charts": structure["charts"],
        "images": structure["images"],
    }


# ---------------------------------------------------------------------------
# workbook map + locale
# ---------------------------------------------------------------------------

def _norm_dims(ref):
    """Match openpyxl's ws.dimensions display ('A1' -> 'A1:A1')."""
    if not ref:
        return "A1:A1"
    if ":" in ref:
        return ref
    try:
        col = "".join(ch for ch in ref if ch.isalpha())
        row = "".join(ch for ch in ref if ch.isdigit())
        return f"{col}{row}:{col}{row}"
    except Exception:  # noqa: BLE001
        return ref


def build_workbook_map_raw(path, zipf, wb_raw, sheets, structures, result):
    names = [
        {"name": dn["name"], "attr_text": dn["attr_text"],
         "local_sheet_id": dn["local_sheet_id"]}
        for dn in wb_raw["defined_names"]
    ]

    entries = []
    available = True
    for rid in wb_raw["external_refs"]:
        rel = wb_raw["rels"].get(rid)
        target = None
        if rel is None:
            available = False
            result.diagnose(
                "external reference rId without a relationship entry; "
                "target unavailable.")
            entries.append({"target": None})
            continue
        if rel["external"]:
            target = rel["target"]
        else:
            ext_root = _read_xml(zipf, rel["target"], result)
            if ext_root is not None:
                for el in ext_root.iter():
                    if el.tag.endswith("fileLink"):
                        target = el.get("target")
                        break
        entries.append({"target": target})
    external = {"available": available, "count": len(entries),
                "links": entries}

    BOOL_KEYS = {"fullCalcOnLoad", "iterate", "fullPrecision",
                 "calcCompleted", "calcOnSave", "concurrentCalc",
                 "forceFullCalc"}
    INT_KEYS = {"calcId", "iterateCount"}
    FLOAT_KEYS = {"iterateDelta"}
    settings = {}
    for key, raw_value in wb_raw["calc_pr"].items():
        text = str(raw_value)
        if key in BOOL_KEYS:
            settings[key] = text.lower() in ("1", "true")
        elif key in INT_KEYS:
            try:
                settings[key] = int(text)
            except ValueError:
                settings[key] = common.jsonable(raw_value)
        elif key in FLOAT_KEYS:
            try:
                settings[key] = float(text)
            except ValueError:
                settings[key] = common.jsonable(raw_value)
        else:
            settings[key] = common.jsonable(raw_value)
    settings["date_system"] = "1904" if wb_raw["date1904"] else "1900"

    parts = []
    for sheet, structure in zip(sheets, structures):
        tables = ",".join(
            f"{t['name']}={t['ref']}" for t in sorted(
                structure["tables"], key=lambda t: str(t["name"])))
        parts.append(f"{sheet['name']}|"
                     f"{_norm_dims(structure['_dimension_ref'])}|{tables}")

    formulas_total = sum(s["formula_summary"]["total"] for s in sheets)
    listed_total = sum(s["formula_summary"]["listed"] for s in sheets)
    cached_available = sum(s["formula_summary"]["cached_available"]
                           for s in sheets)
    truncated = any(s["formula_summary"]["truncated"] for s in sheets)

    return {
        "path": str(Path(path).resolve()),
        "name": Path(path).name,
        "file_fingerprint": common.fingerprint_file(path),
        "structure_fingerprint":
            common.fingerprint_structure_parts(parts),
        "sheet_count": len(sheets),
        "sheet_names": [s["name"] for s in sheets],
        "defined_names": names,
        "external_links_status": external,
        "calculation_settings": settings,
        "cached_values": {
            "formulas_total": formulas_total,
            "formula_entries_listed": listed_total,
            "formula_list_truncated": truncated,
            "cached_available": cached_available,
            "values_recalculated": False,
            "note": "cached values are prior-save artifacts; openpyxl never "
                    "recomputes formulas, so no recalculation is claimed. "
                    "formula_summary per sheet aggregates ALL formulas.",
        },
    }


def _strip_quoted_sections(fmt):
    fmt = re.sub(r'\\.', "", fmt)
    fmt = re.sub(r'"[^"]*"', "", fmt)
    fmt = re.sub(r"\[[^\]]*\]", "", fmt)
    return fmt


def _date_format_like(fmt):
    stripped = _strip_quoted_sections(fmt).lower()
    if "#" in stripped:
        return False
    return bool(re.search(r"y{1,4}|d{1,4}|h{1,2}|s{1,2}", stripped))


def build_locale(sheets):
    counter = {}
    for sheet in sheets:
        for col in sheet["columns"]:
            fmt = col["number_format"]
            if fmt:
                counter[fmt] = counter.get(fmt, 0) + 1
    ordered = sorted(counter.items(), key=lambda kv: (-kv[1], kv[0]))
    number_patterns = [{"format": f, "count": c} for f, c in ordered[:5]]
    date_patterns = [f for f, _ in ordered if _date_format_like(f)]
    comma = sum(1 for f, c in ordered
                if re.search(r"0,0", _strip_quoted_sections(f)))
    point = sum(1 for f, c in ordered
                if re.search(r"0\.0", _strip_quoted_sections(f)))
    guess = "comma" if comma > point else "point" if point > comma else None
    evidence = [f"decimal-comma formats: {comma}",
                f"decimal-point formats: {point}"]
    if guess == "comma":
        evidence.append("formats suggest comma decimal separator")
    elif guess == "point":
        evidence.append("formats suggest point decimal separator")
    else:
        evidence.append("no clear decimal separator signal")
    return {
        "formula_locale": "excel-formula/en",
        "formula_locale_note": "stored formulas always carry English "
                               "function names; no locale conversion is "
                               "applied to cell values.",
        "number_format_patterns": number_patterns,
        "date_format_patterns": date_patterns,
        "possible_csv_locale": {
            "decimal_separator_guess": guess,
            "confidence_level": "LOW",
            "evidence": evidence,
        },
    }


# ---------------------------------------------------------------------------
# orchestration + CLI
# ---------------------------------------------------------------------------

def build_document_map(path, samples, result, formula_cap=FORMULA_LIST_CAP):
    handle_a = common.load_workbook_safe(path, read_only=True)
    handle_b = common.load_workbook_safe(path, read_only=True, data_only=True)
    try:
        try:
            zipf = zipfile.ZipFile(path)
        except zipfile.BadZipFile as exc:
            raise common.XlsxError(
                "VALIDATION_FAILED",
                f"Workbook is not a readable zip package: {exc}",
                recovery="The file may be damaged or not a real .xlsx; "
                         "open and re-save it in Excel (or LibreOffice).",
                context={"path": str(Path(path).resolve())}) from exc
        try:
            wb_raw = read_workbook_raw(zipf, result)
            meta_by_name = {s["name"]: s for s in wb_raw["sheets"]}

            per_sheet = {}
            formulas_by_sheet = {}
            for sheet_name in handle_a.sheetnames:
                ws_ro = handle_a.wb[sheet_name]
                profiles, columns, formulas, failures, caps = scan_pass_a(
                    ws_ro, samples, result, sheet_name)
                per_sheet[sheet_name] = {
                    "profiles": profiles, "columns": columns, "caps": caps,
                    "formulas": formulas,
                }
                formulas_by_sheet[sheet_name] = formulas
                for prop, exc in sorted(failures):
                    result.warn(
                        "STYLE_READ_PARTIAL",
                        f"{sheet_name}: cell.{prop} could not be read "
                        f"({exc}); the style summary for this sheet is "
                        "partial.", sheet=sheet_name, property=prop)

            cached = scan_pass_b(handle_b, formulas_by_sheet)

            sheets = []
            structures = []
            for index, sheet_name in enumerate(handle_a.sheetnames):
                data = per_sheet[sheet_name]
                meta = meta_by_name.get(sheet_name)
                if meta is None:
                    meta = {"name": sheet_name, "state": "visible",
                            "part": None}
                structure = extract_structure_raw(zipf, meta, result)
                structures.append(structure)
                sheets.append(assemble_sheet(
                    sheet_name, index, data, structure, cached, result,
                    samples, formula_cap))

            workbook = build_workbook_map_raw(path, zipf, wb_raw, sheets,
                                              structures, result)
            locale = build_locale(sheets)
            result.set(workbook=workbook, sheets=sheets, locale=locale)
            return result
        finally:
            zipf.close()
    finally:
        handle_a.close()
        handle_b.close()
def _jd(value):
    from datetime import time  # noqa: PLC0415
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    return str(value)


def emit_map(result, pretty):
    payload = dict(result.payload)
    payload["warnings"] = result.warnings
    payload["unsupported"] = result.unsupported
    payload["diagnostics"] = result.diagnostics
    if pretty:
        text = json.dumps(payload, ensure_ascii=False, sort_keys=True,
                          indent=2, default=_jd)
    else:
        text = json.dumps(payload, ensure_ascii=False, sort_keys=True,
                          separators=(",", ":"), default=_jd)
    print(text)
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Build a deterministic, read-only Document Map of an "
                    ".xlsx workbook (Phase 2A: document understanding; no "
                    "writing, no template intelligence).")
    ap.add_argument("file", help="path to .xlsx file")
    ap.add_argument("--pretty", action="store_true",
                    help="indented JSON instead of compact canonical JSON")
    ap.add_argument("--samples", type=int, default=SAMPLE_DEFAULT,
                    metavar="N",
                    help=f"sample values per column (default {SAMPLE_DEFAULT},"
                         f" max {SAMPLE_MAX})")
    ap.add_argument("--formula-list-cap", type=int, default=FORMULA_LIST_CAP,
                    metavar="N",
                    help="max listed formula entries per sheet (aggregates "
                         "always cover all formulas; default "
                         f"{FORMULA_LIST_CAP}; 0 = summary only)")
    args = ap.parse_args(argv)

    if args.samples < 0:
        raise common.XlsxError(
            "SPEC_INVALID", "--samples must be >= 0.",
            recovery="Pass a non-negative integer, e.g. --samples 5.",
            context={"samples": args.samples})
    if args.formula_list_cap < 0:
        raise common.XlsxError(
            "SPEC_INVALID", "--formula-list-cap must be >= 0.",
            recovery="Pass a non-negative integer, e.g. "
                     "--formula-list-cap 5000.",
            context={"formula_list_cap": args.formula_list_cap})

    result = common.Result(mode="understand", file=str(Path(args.file)))
    samples = args.samples
    if samples > SAMPLE_MAX:
        result.warn("SAMPLES_CLAMPED",
                    f"--samples {samples} clamped to {SAMPLE_MAX} "
                    "(never silent).",
                    requested=samples, applied=SAMPLE_MAX)
        samples = SAMPLE_MAX

    build_document_map(args.file, samples, result,
                       formula_cap=args.formula_list_cap)
    return emit_map(result, args.pretty)


if __name__ == "__main__":
    sys.exit(common.guard(main)())
