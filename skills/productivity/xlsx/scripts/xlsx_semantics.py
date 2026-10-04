#!/usr/bin/env python3
"""FAZ 2B -- semantic understanding & mapping (read-only, plan only).

Consumes the FAZ 2A Document Map (JSON) and nothing else. It never
opens a workbook itself, never imports openpyxl directly and never
writes anything: every output is a deterministic, JSON-serialisable
data model (semantics / template profile / match report). Execution
(writing values, expanding rows, propagating formulas) belongs to a
later phase and is deliberately absent here.

Commands:
    xlsx_semantics.py BOOK.xlsx --emit semantics
    xlsx_semantics.py BOOK.xlsx --emit profile --out book.profile.json
    xlsx_semantics.py NEW.xlsx --match book.profile.json
    xlsx_semantics.py --doc-map map.json --emit semantics
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_common as common          # noqa: E402
import xlsx_understand as understand  # noqa: E402

SEMANTICS_VERSION = "2b.1"
PROFILE_VERSION = "2b.1"
DOC_MAP_SCHEMA_VERSION = "2a.1"

ROLE_LEVEL_HIGH = 80
ROLE_LEVEL_MEDIUM = 65
AMBIGUITY_MARGIN = 10

BLOCK_LIST_CAP = 2000
SLOT_LIST_CAP = 2000
PROFILE_COLUMN_CAP = 2000   # listed profile columns; blanks are summarised
MATCH_COLUMN_CAP = 2000     # instance columns scored during matching
MATCH_LIST_CAP = 2000

VALUE_KIND_BY_DATA_TYPE = {
    "string": "string", "integer": "integer", "decimal": "decimal",
    "datetime": "datetime", "date": "date", "boolean": "boolean",
    "formula": "formula", "mixed": "mixed", "blank": "blank",
}
SUPPORTED_VALUE_KINDS = ("string", "integer", "decimal", "date",
                         "datetime", "boolean", "formula", "blank", "mixed")

DICTIONARY_PATH = (Path(__file__).resolve().parent.parent / "references"
                   / "semantic-dictionary.tr-en.json")

_DICT_CACHE = {}


# ---------------------------------------------------------------------------
# normalisation + dictionary
# ---------------------------------------------------------------------------

def normalize_term(text):
    """Reuse the 2A header normaliser -- never a second implementation."""
    return understand.normalize_header(text)


def load_dictionary(path=None):
    """Load + index the versioned TR/EN dictionary (deterministic)."""
    key = str(path or DICTIONARY_PATH)
    cached = _DICT_CACHE.get(key)
    if cached is not None:
        return cached
    target = Path(path) if path else DICTIONARY_PATH
    if not target.exists():
        raise common.XlsxError(
            "FILE_NOT_FOUND",
            f"semantic dictionary not found: {target}",
            recovery="Restore references/semantic-dictionary.tr-en.json "
                     "or pass --dictionary PATH.",
            context={"dictionary": str(target)})
    try:
        data = json.loads(target.read_text(encoding="utf-8"))
    except ValueError as exc:
        raise common.XlsxError(
            "VALIDATION_FAILED",
            f"semantic dictionary is not valid JSON: {exc}",
            recovery="Fix the dictionary file; it must be UTF-8 JSON.",
            context={"dictionary": str(target)}) from exc
    version = data.get("dictionary_version")
    roles = data.get("roles")
    if not version or not isinstance(roles, dict) or not roles:
        raise common.XlsxError(
            "VALIDATION_FAILED",
            "semantic dictionary needs 'dictionary_version' and a "
            "non-empty 'roles' object.",
            recovery="See SKILL.md (Semantic Layer, Profiles & Mapping).",
            context={"dictionary": str(target), "keys": sorted(data)})
    terms = {}
    conflicts = []
    for role, items in roles.items():
        for term in items:
            norm = normalize_term(term)
            if norm in terms and terms[norm] != role:
                conflicts.append({"term": norm, "roles": [terms[norm], role]})
                continue
            terms[norm] = role
    units = {}
    for unit, syn in (data.get("units") or {}).items():
        for item in syn:
            units.setdefault(normalize_term(item), unit)
    out = {
        "version": version,
        "terms": terms,
        "units": units,
        "keywords": data.get("keywords") or {},
        "conflicts": conflicts,
    }
    _DICT_CACHE[key] = out
    return out


# ---------------------------------------------------------------------------
# confidence / ambiguity vocabulary (2A thresholds preserved)
# ---------------------------------------------------------------------------

def level_for(score):
    if score >= ROLE_LEVEL_HIGH:
        return "HIGH"
    if score >= ROLE_LEVEL_MEDIUM:
        return "MEDIUM"
    return "LOW"


def ambiguity_entry(best, second, kind, subject):
    """Return (ambiguity_list, requires_review, margin).

    margin = best - second; below AMBIGUITY_MARGIN the decision is
    ambiguous and must be reviewed (never accepted silently).
    """
    if second is None:
        return [], False, None
    margin = best - second
    if margin >= AMBIGUITY_MARGIN:
        return [], False, margin
    return ([{"kind": kind, "subject": subject, "best": best,
              "second": second, "margin": margin,
              "policy": f"margin < {AMBIGUITY_MARGIN}"}], True, margin)


def decision(score, evidence, second=None, kind="decision", subject=None):
    """Build the shared decision block (score + level + evidence)."""
    ambiguity, review, margin = ambiguity_entry(score, second, kind, subject)
    block = {
        "confidence_score": score,
        "confidence_level": level_for(score),
        "evidence": list(evidence),
        "ambiguity": ambiguity,
        "requires_review": review,
    }
    if margin is not None:
        block["ambiguity_margin"] = margin
    return block


def cap_list(items, cap, result, code, what, **context):
    """Cap a list; capped output is always warned about (never silent)."""
    total = len(items)
    if total > cap:
        result.warn(code, f"{what}: {total} items found; the first {cap} "
                          "are listed (never silent).",
                    total=total, returned=cap, **context)
        listed = items[:cap]
    else:
        listed = items
    return listed, {"count_total": total, "returned_count": len(listed),
                    "truncated": total > len(listed)}


# ---------------------------------------------------------------------------
# identity: semantic content identity -> profile_id  (mtime-free)
# ---------------------------------------------------------------------------

def _hidden_indexes(items):
    out = []
    for item in items or []:
        if isinstance(item, dict):
            idx = item.get("index")
            if idx is not None:
                out.append(idx)
        else:
            out.append(item)
    return sorted(out)


def _project_sheet(sheet):
    """Semantic content projection of one sheet (no paths, no mtime)."""
    dims = sheet.get("dimensions") or {}
    return {
        "name": sheet.get("name"),
        "visibility": sheet.get("visibility"),
        "used_range": dims.get("used_range"),
        "regions": [[r.get("type"), r.get("range")]
                    for r in sheet.get("regions") or []],
        "headers": [[[lab.get("column"), lab.get("normalized")]
                     for lab in header.get("labels") or []]
                    for header in sheet.get("headers") or []],
        "columns": [[c.get("letter"), c.get("normalized_header"),
                     c.get("data_type"), c.get("number_format")]
                    for c in sheet.get("columns") or []],
        "tables": sorted([t.get("name"), t.get("ref")]
                         for t in sheet.get("tables") or []),
        "merged_cells": sorted(m.get("range")
                               for m in sheet.get("merged_cells") or []),
        "hidden_rows": _hidden_indexes((sheet.get("hidden") or {}).get("rows")),
        "hidden_columns": _hidden_indexes(
            (sheet.get("hidden") or {}).get("columns")),
        "freeze_panes": sheet.get("freeze_panes"),
        "autofilter": sheet.get("autofilter"),
        "validations": sorted([[v.get("range"), v.get("type")]
                               for v in sheet.get("validations") or []],
                              key=lambda pair: [str(p) for p in pair]),
        "conditional_formats": sorted([[c.get("range"), c.get("kind")]
                                       for c in sheet.get("conditional_formats")
                                       or []],
                                      key=lambda pair: [str(p) for p in pair]),
        "formula_summary": sheet.get("formula_summary"),
        "repeated_structures": [[r.get("type"), r.get("signature"),
                                 r.get("occurrences"), r.get("row_count_per_run")]
                                for r in sheet.get("repeated_structures") or []],
    }


def content_identity(doc_map, dictionary_version):
    """Hash of the workbook's semantic content -- deliberately mtime-free."""
    workbook = doc_map.get("workbook") or {}
    projection = {
        "doc_map_schema": DOC_MAP_SCHEMA_VERSION,
        "dictionary_version": dictionary_version,
        "sheets": [_project_sheet(s) for s in doc_map.get("sheets") or []],
        "defined_names": sorted(
            [n.get("name") if isinstance(n, dict) else str(n)
             for n in workbook.get("defined_names") or []]),
        "calculation_settings": workbook.get("calculation_settings"),
        "locale": {k: (doc_map.get("locale") or {}).get(k)
                   for k in ("formula_locale", "number_format_patterns",
                             "date_format_patterns", "possible_csv_locale")},
    }
    return common.sha256_text(common.canonical_json(projection))


def profile_identity(doc_map, dictionary_version):
    """content_identity + structure fingerprint + schema + dictionary."""
    workbook = doc_map.get("workbook") or {}
    identity = content_identity(doc_map, dictionary_version)
    structure = workbook.get("structure_fingerprint")
    profile_id = common.sha256_text(common.canonical_json({
        "content_identity": identity,
        "structure_fingerprint": structure,
        "profile_version": PROFILE_VERSION,
        "dictionary_version": dictionary_version,
    }))
    return {"content_identity": identity,
            "structure_fingerprint": structure,
            "profile_id": profile_id}


# ---------------------------------------------------------------------------
# document map input (workbook -> 2A, or a saved 2A JSON)
# ---------------------------------------------------------------------------

def build_map_from_workbook(path, result, samples, formula_cap=None):
    """Run 2A (the single Document Map producer) and forward its findings."""
    upstream = common.Result()
    cap = understand.FORMULA_LIST_CAP if formula_cap is None else formula_cap
    understand.build_document_map(path, samples, upstream, formula_cap=cap)
    for item in upstream.warnings:
        result.warn(item.get("code", "UPSTREAM_WARNING"),
                    item.get("message", ""), source="2a",
                    **(item.get("context") or {}))
    for item in upstream.unsupported:
        result.unsupported_item(item.get("message", ""), source="2a",
                                **(item.get("context") or {}))
    for item in upstream.diagnostics:
        result.diagnose(item.get("message", ""), source="2a",
                        **(item.get("context") or {}))
    return dict(upstream.payload)


def load_doc_map(doc_map_path, result):
    """Load a saved 2A Document Map JSON file."""
    target = Path(doc_map_path)
    if not target.exists():
        raise common.XlsxError(
            "FILE_NOT_FOUND", f"document map not found: {target}",
            recovery="Generate one with `xlsx_understand.py BOOK.xlsx > map.json`.",
            context={"doc_map": str(target)})
    try:
        doc = json.loads(target.read_text(encoding="utf-8"))
    except ValueError as exc:
        raise common.XlsxError(
            "VALIDATION_FAILED", f"document map is not valid JSON: {exc}",
            recovery="Re-generate the map with xlsx_understand.py.",
            context={"doc_map": str(target)}) from exc
    if not isinstance(doc, dict) or "sheets" not in doc:
        raise common.XlsxError(
            "VALIDATION_FAILED",
            "document map JSON needs a 'sheets' array.",
            recovery="Pass a file produced by xlsx_understand.py.",
            context={"doc_map": str(target), "keys": sorted(doc)
                     if isinstance(doc, dict) else None})
    return doc


# ---------------------------------------------------------------------------
# emission
# ---------------------------------------------------------------------------

def emit(payload, result, pretty):
    out = dict(payload)
    out["warnings"] = result.warnings
    out["unsupported"] = result.unsupported
    out["diagnostics"] = result.diagnostics
    if pretty:
        text = json.dumps(out, ensure_ascii=False, sort_keys=True, indent=2,
                          default=str)
    else:
        text = json.dumps(out, ensure_ascii=False, sort_keys=True,
                          separators=(",", ":"), default=str)
    print(text)
    return 0

# ---------------------------------------------------------------------------
# B1 -- semantic labeling (roles, value_kind, unit_guess)
# ---------------------------------------------------------------------------

RANGE_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$")
DATE_FORMAT_HINT = re.compile(r"(y{2,4}|d{1,2}[./-]m{1,2}|m{1,2}[./-]y{2,4}|"
                              r"m{1,4}[./-]d{1,2})", re.IGNORECASE)
CURRENCY_HINTS = ("₺", "tl", "try", "$", "usd", "€", "eur", "£", "gbp")


def _col_letter_to_index(letter):
    out = 0
    for ch in letter.upper():
        out = out * 26 + (ord(ch) - 64)
    return out


def range_columns(ref):
    """Column letters covered by a 2A range string, or None if unparsable."""
    match = RANGE_RE.match((ref or "").strip())
    if not match:
        return None
    start, end = match.group(1), match.group(3) or match.group(1)
    first, last = _col_letter_to_index(start), _col_letter_to_index(end)
    if last < first:
        first, last = last, first
    return first, last


def _validated_columns(sheet):
    """letter -> number of validations covering that column."""
    counts = Counter()
    for entry in sheet.get("validations") or []:
        bounds = range_columns(entry.get("range"))
        if not bounds:
            continue
        for idx in range(bounds[0], bounds[1] + 1):
            counts[idx] += 1
    return counts


def _is_date_format(fmt):
    return bool(DATE_FORMAT_HINT.search(str(fmt or "")))


def _format_hints(fmt):
    text = str(fmt or "")
    low = text.casefold()
    return ("%" in text,
            any(hint in low for hint in CURRENCY_HINTS),
            _is_date_format(text))


def _header_evidence(col, dictionary):
    """Family A -- header text evidence."""
    raw = col.get("header")
    norm = col.get("normalized_header") or (normalize_term(raw) if raw else None)
    hits = []
    if not norm:
        return norm, hits
    exact = dictionary["terms"].get(norm)
    if exact:
        hits.append((exact, 45, f"header_term:{norm}"))
    for term, role in dictionary["terms"].items():
        if term == norm or len(term) < 3:
            continue
        if term in norm:
            hits.append((role, 25, f"header_partial:{term}"))
        elif len(norm) >= 5 and norm in term:
            hits.append((role, 12, f"header_partial_rev:{term}"))
    keywords = dictionary["keywords"]
    families = (
        ("identifier", "id_words", 20),
        ("date", "date_words", 25),
        ("percent", "percent_symbols", 20),
        ("free_text", "note_words", 20),
    )
    for role, key, weight in families:
        for word in keywords.get(key) or []:
            if word and word in norm:
                hits.append((role, weight, f"keyword:{key}:{word}"))
                break
    for symbol in keywords.get("currency_symbols") or []:
        if symbol and symbol in norm:
            hits.append(("currency", 25, f"keyword:currency_symbol:{symbol}"))
            break
    return norm, hits


AFFINITY = {
    "string": {"label": 20, "identifier": 10, "code": 10, "address": 10,
               "person_name": 10, "free_text": 10, "period": 8, "flag": 10},
    "integer": {"quantity": 20, "currency": 15, "percent": 15,
                "identifier": 12, "code": 12, "period": 10},
    "decimal": {"quantity": 20, "currency": 18, "percent": 15},
    "datetime": {"date": 20},
    "date": {"date": 20},
    "boolean": {"flag": 20, "code": 8},
    "formula": {"formula_derived": 40},
    "mixed": {},
    "blank": {"empty_slot": 15},
}


def _value_evidence(col):
    """Family B -- data_type / format / samples evidence."""
    dtype = col.get("data_type")
    fmt = col.get("number_format")
    share = col.get("formula_share") or 0.0
    cell_count = col.get("cell_count") or 0
    distinct = col.get("distinct_count") or 0
    hits = []
    percent_fmt, currency_fmt, date_fmt = _format_hints(fmt)
    if percent_fmt:
        hits.append(("percent", 25, f"number_format:percent({fmt})"))
    if currency_fmt:
        hits.append(("currency", 25, f"number_format:currency({fmt})"))
    if date_fmt:
        hits.append(("date", 25, f"number_format:date({fmt})"))
    for role, weight in AFFINITY.get(dtype or "blank", {}).items():
        hits.append((role, weight, f"value_affinity:{role}:{dtype}"))
    if dtype == "formula" or share >= 0.5:
        hits.append(("formula_derived", max(30, int(60 * share)),
                     f"formula_share:{share:.2f}"))
    if dtype == "string" and cell_count >= 3 and cell_count:
        ratio = distinct / cell_count
        if ratio >= 0.9:
            hits.append(("identifier", 15,
                         f"distinct_ratio:{distinct}/{cell_count}"))
    return hits


def _context_evidence(col, sheet, validated):
    """Family C -- structural / contextual evidence."""
    hits = []
    index = col.get("index")
    inside_input_region = False
    for region in sheet.get("regions") or []:
        if region.get("type") != "input_region":
            continue
        bounds = range_columns(region.get("range"))
        if bounds and index and bounds[0] <= index <= bounds[1]:
            hits.append(("empty_slot", 25,
                         f"input_region:{region.get('range')}"))
            inside_input_region = True
            break
    if not inside_input_region and col.get("region") == "input_region":
        hits.append(("empty_slot", 25, "region:input_region"))
    if validated.get(index) and \
            col.get("data_type") in ("string", "integer", "decimal"):
        hits.append(("identifier", 10, "validation:column"))
    norm = col.get("normalized_header") or ""
    if "|" in norm and col.get("data_type") in ("integer", "decimal"):
        hits.append(("quantity", 15, "grouped_header:numeric"))
    if col.get("region") in (None, ""):
        hits.append((None, 0, "region:unknown"))
    return hits


def _score_roles(hits):
    """Aggregate evidence into ranked role candidates (deterministic)."""
    scores = {}
    evidence = {}
    for role, weight, reason in hits:
        if not role:
            continue
        scores[role] = min(100, scores.get(role, 0) + weight)
        evidence.setdefault(role, []).append(reason)
    ranked = sorted(scores.items(), key=lambda item: (-item[1], item[0]))
    return ranked, evidence


def value_kind_of(col, result, sheet_name):
    """value_kind is NOT the semantic role; unknown input is reported."""
    dtype = col.get("data_type")
    share = col.get("formula_share") or 0.0
    if share >= 0.5 and dtype != "formula":
        return "formula"
    kind = VALUE_KIND_BY_DATA_TYPE.get(dtype)
    if kind is None:
        result.warn(
            "VALUE_KIND_UNKNOWN",
            f"{sheet_name}!{col.get('letter')}: 2A data_type "
            f"{dtype!r} is not in the known value_kind map; reported "
            "as 'mixed' (never silent).",
            sheet=sheet_name, column=col.get("letter"), data_type=dtype)
        return "mixed"
    return kind


def unit_guess_of(norm_header, fmt, dictionary, dtype=None):
    """unit_guess only when evidenced (header token or number format)."""
    numeric = dtype in ("integer", "decimal")
    if norm_header and numeric:
        for token in norm_header.replace("/", " ").replace("-", " ").split():
            unit = dictionary["units"].get(token)
            if unit:
                return unit, f"header_unit_token:{token}"
    text = str(fmt or "")
    low = text.casefold()
    for symbol, unit in (("₺", "₺"), ("$", "$"), ("€", "€"), ("%", "%")):
        if symbol in text:
            return unit, f"number_format_symbol:{symbol}"
    for word, unit in (("tl", "₺"), ("try", "₺"), ("usd", "$"), ("eur", "€")):
        if word in low:
            return unit, f"number_format_word:{word}"
    return None, None


def label_column(col, sheet, dictionary, validated, result, sheet_name):
    """One column -> semantic annotation (evidence mandatory)."""
    norm, hits = _header_evidence(col, dictionary)
    hits += _value_evidence(col)
    hits += _context_evidence(col, sheet, validated)
    ranked, evidence = _score_roles(hits)
    best_role, best_score = ranked[0] if ranked else ("unknown", 0)
    second = ranked[1][1] if len(ranked) > 1 else None
    if not ranked:
        hits.append((None, 0, "no header/format/value/context evidence"))
        evidence.setdefault("unknown", []).append(
            "no header/format/value/context evidence")
    block = decision(best_score, evidence.get(best_role, []),
                     second=second, kind="role",
                     subject=f"{sheet_name}!{col.get('letter')}")
    if block["confidence_level"] == "LOW":
        result.warn(
            "LOW_CONFIDENCE_DETECTION",
            f"{sheet_name}!{col.get('letter')}: semantic role "
            f"{best_role!r} scored {best_score} (LOW); evidence: "
            f"{', '.join(block['evidence']) or 'none'}.",
            sheet=sheet_name, column=col.get("letter"), role=best_role,
            score=best_score)
    unit, unit_reason = unit_guess_of(norm, col.get("number_format"),
                                      dictionary, col.get("data_type"))
    value_kind = value_kind_of(col, result, sheet_name)
    out = {
        "letter": col.get("letter"),
        "index": col.get("index"),
        "header": col.get("header"),
        "normalized_header": norm,
        "region": col.get("region"),
        "data_type": col.get("data_type"),
        "number_format": col.get("number_format"),
        "cell_count": col.get("cell_count"),
        "distinct_count": col.get("distinct_count"),
        "null_count": col.get("null_count"),
        "formula_count": col.get("formula_count"),
        "formula_share": col.get("formula_share"),
        "value_kind": value_kind,
        "semantic_role": best_role,
        "role_score": best_score,
        "role_level": block["confidence_level"],
        "confidence_score": block["confidence_score"],
        "confidence_level": block["confidence_level"],
        "unit_guess": unit,
        "unit_evidence": [unit_reason] if unit_reason else [],
        "evidence": block["evidence"],
        "candidates": [{"role": role, "score": score}
                       for role, score in ranked[:3]],
        "ambiguity": block["ambiguity"],
        "requires_review": block["requires_review"],
    }
    if unit is None and any(h for h in hits if h[0] in ("quantity",
                                                        "currency", "percent")):
        result.diagnose(
            f"{sheet_name}!{col.get('letter')}: numeric/quantity evidence "
            "but no unit could be evidenced; unit_guess stays null "
            "(no conversion is attempted in this phase).",
            sheet=sheet_name, column=col.get("letter"))
    return out


def label_sheet(sheet, dictionary, result):
    validated = _validated_columns(sheet)
    columns = [label_column(col, sheet, dictionary, validated, result,
                            sheet.get("name") or "?")
               for col in sheet.get("columns") or []]
    role_counts = Counter(c["semantic_role"] for c in columns)
    return {
        "name": sheet.get("name"),
        "columns": columns,
        "role_counts": dict(sorted(role_counts.items())),
        "coverage": {
            "columns": len(columns),
            "labeled": sum(1 for c in columns if c["semantic_role"] != "unknown"),
            "unknown": role_counts.get("unknown", 0),
        },
    }


def label_workbook(doc_map, dictionary, result):
    return {
        "semantics_version": SEMANTICS_VERSION,
        "dictionary_version": dictionary["version"],
        "sheets": [label_sheet(sheet, dictionary, result)
                   for sheet in doc_map.get("sheets") or []],
    }

# ---------------------------------------------------------------------------
# B2 -- record archetypes / logical record blocks
# ---------------------------------------------------------------------------

BLOCK_ARCHETYPE = {
    "title": "title_block",
    "header": "header_block",
    "subheader": "header_block",
    "data": "record_block",
    "subtotal": "subtotal_block",
    "total": "total_block",
    "notes": "note_block",
    "section_header": "section_block",
    "input_region": "input_slot_block",
}
ROW_ARCHETYPE = {
    "header": "header_row",
    "subheader": "header_row",
    "data": "record_row",
    "subtotal": "subtotal_row",
    "total": "total_row",
    "notes": "note_row",
    "input_region": "input_slot_row",
    "section_header": "section_row",
    "title": "title_row",
}


def _region_bounds(region):
    rows = region.get("rows") or []
    if len(rows) == 2:
        return rows[0], rows[1]
    cols = range_columns(region.get("range"))
    return None, None


def _column_span(region):
    bounds = range_columns(region.get("range"))
    if not bounds:
        return []
    return list(range(bounds[0], bounds[1] + 1))


def _key_candidate(column, validated, block_span):
    """identifier-role column -> candidate key score (never assumed)."""
    cell_count = column.get("cell_count") or 0
    distinct = column.get("distinct_count") or 0
    nulls = column.get("null_count") or 0
    evidence = []
    score = 0
    if column.get("semantic_role") == "identifier":
        score += 40
        evidence.append("semantic_role:identifier")
    if cell_count and distinct:
        ratio = distinct / cell_count
        if ratio >= 1.0:
            score += 25
            evidence.append(f"distinct_ratio:{distinct}/{cell_count} (unique)")
        else:
            score += max(0, int(15 * ratio))
            evidence.append(f"distinct_ratio:{distinct}/{cell_count}")
    if nulls == 0 and cell_count:
        score += 15
        evidence.append("null_count:0")
    if validated.get(column.get("index")):
        score += 10
        evidence.append("validation:column")
    if column.get("letter") and column.get("index") in block_span:
        score += 5
        evidence.append("inside:record_block")
    duplicates = max(0, cell_count - distinct)
    if duplicates:
        evidence.append(f"duplicates:{duplicates}")
    return min(100, score), evidence, duplicates


def detect_blocks(sheet, semantics, result):
    """Regions + semantics -> logical record blocks (archetype + key)."""
    sheet_name = sheet.get("name") or "?"
    validated = _validated_columns(sheet)
    columns_by_letter = {c.get("letter"): c for c in semantics["columns"]}
    blocks = []
    index = 0
    for region in sheet.get("regions") or []:
        start, end = _region_bounds(region)
        if start is None:
            continue
        index += 1
        kind = region.get("type") or "unknown"
        archetype = BLOCK_ARCHETYPE.get(kind, "unclassified_block")
        span = _column_span(region)
        evidence = [f"region_type:{kind}",
                    f"region_confidence:{region.get('confidence_level')}"]
        entries = [columns_by_letter[chr(64 + i)] for i in span
                   if chr(64 + i) in columns_by_letter]
        key_block = None
        scored = []
        for column in entries:
            score, key_evidence, duplicates = _key_candidate(
                column, validated, span)
            scored.append((column, score, key_evidence, duplicates))
        if archetype == "record_block" and scored:
            scored.sort(key=lambda item: (-item[1], item[0].get("letter")))
            column, score, key_evidence, duplicates = scored[0]
            second = scored[1][1] if len(scored) > 1 else None
            block_dec = decision(score, key_evidence, second=second,
                                 kind="record_key",
                                 subject=f"{sheet_name}!{column.get('letter')}")
            key_block = {
                "column": column.get("letter"),
                "semantic_role": column.get("semantic_role"),
                "candidate_key_score": score,
                "null_count": column.get("null_count"),
                "distinct_count": column.get("distinct_count"),
                "duplicate_count": duplicates,
                "evidence": block_dec["evidence"],
                "confidence_score": block_dec["confidence_score"],
                "confidence_level": block_dec["confidence_level"],
                "ambiguity": block_dec["ambiguity"],
                "requires_review": block_dec["requires_review"],
            }
            evidence.append(f"record_key_candidate:{column.get('letter')}"
                            f"@{score}")
            if duplicates:
                result.warn(
                    "KEY_DUPLICATES",
                    f"{sheet_name}!{column.get('letter')}: record key "
                    f"candidate has {duplicates} duplicate value(s); it is "
                    "NOT a proven unique key (never assumed).",
                    sheet=sheet_name, column=column.get("letter"),
                    duplicates=duplicates)
        block_dec = decision(region.get("confidence_score") or 0, evidence,
                             kind="block",
                             subject=f"{sheet_name}!{region.get('range')}")
        blocks.append({
            "block_id": f"b{index}",
            "range": region.get("range"),
            "rows": [start, end],
            "row_archetype": ROW_ARCHETYPE.get(kind, "unknown_row"),
            "archetype": archetype,
            "source_region_type": kind,
            "record_key_candidate": key_block,
            "confidence_score": block_dec["confidence_score"],
            "confidence_level": block_dec["confidence_level"],
            "evidence": block_dec["evidence"],
            "ambiguity": block_dec["ambiguity"],
            "requires_review": block_dec["requires_review"],
        })
    for repeated in sheet.get("repeated_structures") or []:
        evidence = [f"repeated_type:{repeated.get('type')}",
                    f"occurrences:{repeated.get('occurrences')}",
                    f"signature:{repeated.get('signature')}"]
        block_dec = decision(repeated.get("confidence_score") or 0, evidence,
                             kind="pattern",
                             subject=f"{sheet_name}!{repeated.get('signature')}")
        blocks.append({
            "block_id": f"p{repeated.get('signature')}",
            "range": None,
            "rows": None,
            "row_archetype": "record_row",
            "archetype": "repeated_pattern",
            "source_region_type": repeated.get("type"),
            "runs": repeated.get("runs") or [],
            "occurrences": repeated.get("occurrences"),
            "row_count_per_run": repeated.get("row_count_per_run"),
            "record_key_candidate": None,
            "confidence_score": block_dec["confidence_score"],
            "confidence_level": block_dec["confidence_level"],
            "evidence": block_dec["evidence"],
            "ambiguity": block_dec["ambiguity"],
            "requires_review": block_dec["requires_review"],
        })
    if not blocks:
        result.diagnose(f"{sheet_name}: no regions were detected by 2A, so "
                        "no logical blocks could be derived (nothing is "
                        "invented).", sheet=sheet_name)
    return blocks


# ---------------------------------------------------------------------------
# B3 -- template profile (deterministic identity, no timestamps)
# ---------------------------------------------------------------------------

SNIPPET_MAX = 60


def _required_of(column, block_columns):
    """required + required_source: heuristic proposal, never user truth."""
    role = column.get("semantic_role")
    reasons = []
    if column.get("region") == "input_region":
        reasons.append("region:input_region")
    if role in ("formula_derived", "summary"):
        return False, "heuristic", [f"role:{role} is derived, not filled"]
    if column.get("index") in block_columns:
        reasons.append("inside:record_block")
    if column.get("data_type") in ("string", "integer", "decimal", "datetime"):
        reasons.append(f"data_type:{column.get('data_type')}")
    if not reasons:
        reasons.append("no positive evidence")
    return bool(reasons), "heuristic", reasons


def extract_slots(sheet, semantics, result):
    """Input slots: 2A input_region ranges + empty_slot columns."""
    sheet_name = sheet.get("name") or "?"
    slots = []
    for region in sheet.get("regions") or []:
        if region.get("type") != "input_region":
            continue
        slots.append({
            "range": region.get("range"),
            "rows": region.get("rows"),
            "kind": "input_slot",
            "required": True,
            "required_source": "heuristic",
            "evidence": ["region_type:input_region",
                         f"region_confidence:{region.get('confidence_level')}"],
            "confidence_score": region.get("confidence_score"),
            "confidence_level": region.get("confidence_level"),
        })
    for column in semantics["columns"]:
        if column.get("semantic_role") != "empty_slot":
            continue
        slots.append({
            "range": None,
            "rows": None,
            "kind": "input_slot_column",
            "column": column.get("letter"),
            "required": True,
            "required_source": "heuristic",
            "evidence": column.get("evidence") or [],
            "confidence_score": column.get("confidence_score"),
            "confidence_level": column.get("confidence_level"),
        })
    if not slots:
        result.diagnose(f"{sheet_name}: no input slots were evidenced "
                        "(nothing is assumed to be fillable).",
                        sheet=sheet_name)
    return slots


def is_material_column(col):
    """True when a column carries information worth listing.

    A column with no header, no cells, no region, no format and no
    semantic role is blank/unused; listing 16k of them would inflate the
    profile without carrying meaning. They are always COUNTERED (see
    blank_columns_meta) - never silently dropped.
    """
    if col.get("header") or col.get("normalized_header"):
        return True
    if (col.get("cell_count") or 0) > 0:
        return True
    if col.get("region"):
        return True
    role = col.get("semantic_role")
    if role and role not in ("unknown", "empty_slot"):
        return True
    fmt = col.get("number_format")
    if fmt and fmt != "General":
        return True
    dtype = col.get("data_type")
    if dtype and dtype != "blank":
        return True
    return False


def blank_columns_meta(columns, sheet_name, result, *, where):
    """Split material vs blank columns and report the blanks explicitly."""
    material = [c for c in columns if is_material_column(c)]
    blanks = [c for c in columns if not is_material_column(c)]
    meta = {"count": len(blanks), "listed": False,
            "material_count": len(material),
            "note": ("blank/unused columns (no header, no cells, no "
                     "region, no format) are counted but not listed")}
    if blanks:
        letters = [c.get("letter") for c in blanks if c.get("letter")]
        meta["first_letter"] = letters[0] if letters else None
        meta["last_letter"] = letters[-1] if letters else None
        result.warn(
            "BLANK_COLUMNS_SUMMARISED",
            f"{sheet_name}: {len(blanks)} blank/unused columns summarised "
            f"(first {meta['first_letter']}, last {meta['last_letter']}) in "
            f"the {where}; they are counted and reported, not listed "
            "(never silent).",
            sheet=sheet_name, count=len(blanks), where=where)
    return material, meta


def build_profile(doc_map, semantics, dictionary, result, source_meta=None):
    """Deterministic template profile: same content -> same bytes."""
    identity = profile_identity(doc_map, dictionary["version"])
    workbook = doc_map.get("workbook") or {}
    sheets = []
    for sheet, sem in zip(doc_map.get("sheets") or [],
                          semantics.get("sheets") or []):
        blocks = detect_blocks(sheet, sem, result)
        slots = extract_slots(sheet, sem, result)
        block_columns = set()
        for block in blocks:
            if block.get("rows") and block.get("archetype") == "record_block":
                block_columns.update(_column_span({"range": block["range"]}))
        material_columns, blank_meta = blank_columns_meta(
            sem["columns"], sheet.get("name"), result, where="profile")
        columns = []
        for column in material_columns:
            required, required_source, reasons = _required_of(column,
                                                              block_columns)
            columns.append({
                "letter": column["letter"],
                "index": column["index"],
                "header": column["header"],
                "normalized_header": column["normalized_header"],
                "region": column["region"],
                "semantic_role": column["semantic_role"],
                "role_score": column["role_score"],
                "role_level": column["role_level"],
                "value_kind": column["value_kind"],
                "data_type": column["data_type"],
                "number_format": column["number_format"],
                "cell_count": column.get("cell_count"),
                "distinct_count": column.get("distinct_count"),
                "null_count": column.get("null_count"),
                "unit_guess": column["unit_guess"],
                "required": required,
                "required_source": required_source,
                "required_evidence": reasons,
                "evidence": column["evidence"],
                "ambiguity": column["ambiguity"],
                "requires_review": column["requires_review"],
            })
        columns_listed, columns_meta = cap_list(
            columns, PROFILE_COLUMN_CAP, result, "COLUMNS_TRUNCATED",
            f"{sheet.get('name')}: profile columns",
            sheet=sheet.get("name"))
        slots_listed, slots_meta = cap_list(
            slots, SLOT_LIST_CAP, result, "SLOTS_TRUNCATED",
            f"{sheet.get('name')}: input slots", sheet=sheet.get("name"))
        blocks_listed, blocks_meta = cap_list(
            blocks, BLOCK_LIST_CAP, result, "BLOCKS_TRUNCATED",
            f"{sheet.get('name')}: blocks", sheet=sheet.get("name"))
        sheets.append({
            "name": sheet.get("name"),
            "visibility": sheet.get("visibility"),
            "used_range": (sheet.get("dimensions") or {}).get("used_range"),
            "regions": sheet.get("regions") or [],
            "columns": columns_listed,
            "columns_cap": columns_meta,
            "blank_columns": blank_meta,
            "blocks": blocks_listed,
            "blocks_cap": blocks_meta,
            "slots": slots_listed,
            "slots_cap": slots_meta,
            "formula_summary": sheet.get("formula_summary"),
            "tables": sheet.get("tables") or [],
            "merged_cells": [m.get("range")
                             for m in sheet.get("merged_cells") or []],
            "hidden": sheet.get("hidden") or {},
            "freeze_panes": sheet.get("freeze_panes"),
            "autofilter": sheet.get("autofilter"),
            "validations": sheet.get("validations") or [],
            "conditional_formats": sheet.get("conditional_formats") or [],
        })
    profile = {
        "profile_version": PROFILE_VERSION,
        "generator": f"xlsx_semantics.py {SEMANTICS_VERSION}",
        "profile_id": identity["profile_id"],
        "content_identity": identity["content_identity"],
        "structure_fingerprint": identity["structure_fingerprint"],
        "dictionary_version": dictionary["version"],
        "doc_map_schema": DOC_MAP_SCHEMA_VERSION,
        "workbook_name": workbook.get("name"),
        "sheet_count": workbook.get("sheet_count"),
        "source_bytes_sha256": (source_meta or {}).get("sha256"),
        "source_size": (source_meta or {}).get("size"),
        "sheets": sheets,
    }
    profile["profile_hash"] = common.sha256_text(common.canonical_json(profile))
    return profile

# ---------------------------------------------------------------------------
# B4 -- instance <-> template matching (read-only, never executes)
# ---------------------------------------------------------------------------

WEIGHTS = {"column_score": 0.40, "region_score": 0.25,
           "structural_score": 0.20, "formula_score": 0.15}

MIN_COLUMN_MATCH = 30         # below this a pair is UNMATCHED, not a match


def _interval_similarity(a, b):
    if not a or not b:
        return 0
    span = max(max(a) - min(a), max(b) - min(b), 1)
    gap = abs(min(a) - min(b))
    return max(0, 100 - int(100 * gap / span))


def _kinds_similarity(left, right):
    left = {str(k): v for k, v in (left or {}).items()}
    right = {str(k): v for k, v in (right or {}).items()}
    if not left and not right:
        return 100, "both sides declare no formulas"
    keys = set(left) | set(right)
    total = sum(left.values()) + sum(right.values()) or 1
    distance = sum(abs(left.get(k, 0) - right.get(k, 0)) for k in keys)
    return max(0, 100 - int(100 * distance / total)), (
        f"formula kind mix: {sorted(keys)}")


def _region_match(template_regions, instance_regions):
    """Match regions by type + relative row position (evidence kept)."""
    remaining = list(enumerate(instance_regions))
    matched, unmatched, ambiguity = [], [], []
    for t_region in template_regions:
        best = None
        for pos, i_region in remaining:
            if i_region.get("type") != t_region.get("type"):
                continue
            score = _interval_similarity(t_region.get("rows"),
                                         i_region.get("rows"))
            if best is None or score > best[0]:
                best = (score, pos, i_region)
        if best is None:
            unmatched.append({"range": t_region.get("range"),
                              "type": t_region.get("type"),
                              "reason": "no instance region of this type"})
            continue
        score, pos, i_region = best
        second = None
        for pos2, i2 in remaining:
            if pos2 == pos or i2.get("type") != t_region.get("type"):
                continue
            other = _interval_similarity(t_region.get("rows"), i2.get("rows"))
            second = other if second is None else max(second, other)
        block = decision(score, [f"type:{t_region.get('type')}",
                                 f"rows:{t_region.get('rows')} -> "
                                 f"{i_region.get('rows')}"],
                         second=second, kind="region",
                         subject=str(t_region.get("range")))
        matched.append({"template_range": t_region.get("range"),
                        "instance_range": i_region.get("range"),
                        "type": t_region.get("type"),
                        **{k: block[k] for k in
                           ("confidence_score", "confidence_level",
                            "evidence", "ambiguity", "requires_review")}})
        if block["requires_review"]:
            ambiguity.append(block["ambiguity"])
        remaining = [item for item in remaining if item[0] != pos]
    extras = [i_region.get("range") for _, i_region in remaining]
    return matched, unmatched, extras, ambiguity


def _prep_column(col):
    """Pre-extract the compared fields once (matching is O(n*m))."""
    return (col.get("normalized_header"), col.get("semantic_role"),
            col.get("value_kind"), col.get("number_format"),
            col.get("letter"))


def _pair_score(t_col, i_col, want_reasons=False):
    """THE column-pair score (single implementation).

    ``t_col``/``i_col`` may be a raw column dict or a ``_prep_column``
    tuple. ``want_reasons`` only controls evidence building, never the
    number, so a fast scan and the recorded evidence agree exactly.
    """
    if isinstance(t_col, dict):
        t_col = _prep_column(t_col)
    if isinstance(i_col, dict):
        i_col = _prep_column(i_col)
    t_header, t_role, t_kind, t_fmt, t_letter = t_col
    i_header, i_role, i_kind, i_fmt, i_letter = i_col
    score, reasons = 0, []
    if t_header and t_header == i_header:
        score += 55
        if want_reasons:
            reasons.append(f"header_exact:{t_header}")
    if t_role and t_role == i_role:
        score += 25
        if want_reasons:
            reasons.append(f"role:{t_role}")
    if t_kind and t_kind == i_kind:
        score += 12
        if want_reasons:
            reasons.append(f"value_kind:{t_kind}")
    if t_fmt == i_fmt:
        score += 8
        if want_reasons:
            reasons.append(f"number_format:{t_fmt}")
    if t_letter == i_letter:
        score += 5
        if want_reasons:
            reasons.append("same_letter")
    return min(100, score), reasons


def _score_column_pair(t_col, i_col):
    """Full-score comparison of one template column against one instance
    column. Kept as one function so the best/second-best margin is exact."""
    return _pair_score(t_col, i_col, want_reasons=True)


def _competition_ambiguity(template_cols, instance_cols, matched):
    """Symmetric candidate ranking (spec 9.1 / 9.2).

    The greedy pass ranks candidates per TEMPLATE column. The opposite
    direction matters just as much: an instance column is also contested
    when a DIFFERENT template column scores within AMBIGUITY_MARGIN of the
    template it was matched to. Both directions are reported; neither is
    ever silently accepted.
    """
    by_letter = {c.get("letter"): c for c in template_cols}
    instance_by_letter = {c.get("letter"): c for c in instance_cols}
    entries = []
    for entry in matched:
        t_letter = entry.get("template_letter")
        i_letter = entry.get("instance_letter")
        own = entry.get("confidence_score")
        i_col = instance_by_letter.get(i_letter)
        if own is None or i_col is None:
            continue
        rivals = []
        for other_letter, other_col in by_letter.items():
            if other_letter == t_letter:
                continue
            score, _reasons = _score_column_pair(other_col, i_col)
            if score >= MIN_COLUMN_MATCH:
                rivals.append((score, other_letter))
        if not rivals:
            continue
        rivals.sort(key=lambda item: (-item[0], item[1]))
        rival_score, rival_letter = rivals[0]
        ambiguity, review, _margin = ambiguity_entry(
            own, rival_score, "column_competition",
            f"instance:{i_letter}, matched template:{t_letter}, "
            f"rival template:{rival_letter}")
        if review:
            entry["ambiguity"] = list(entry.get("ambiguity") or []) + ambiguity
            entry["requires_review"] = True
            entry["evidence"] = list(entry.get("evidence") or []) + [
                f"competition:template:{rival_letter}@{rival_score}"]
            entries.append(ambiguity)
    return entries


def _column_match(template_cols, instance_cols):
    """Columns: normalized header, then role + value_kind + format.

    best/second-best come from the SAME full score function, so the
    ambiguity margin is exact (never approximated).
    """
    remaining = [(pos, col, _prep_column(col))
                 for pos, col in enumerate(instance_cols)]
    matched, unmatched, conflicts, ambiguity = [], [], [], []
    for t_col in template_cols:
        t_prep = _prep_column(t_col)
        # fast scan: scores only (same arithmetic as the evidence path)
        scored = sorted(
            ((_pair_score(t_prep, prep)[0], pos) for pos, _col, prep in
             remaining), key=lambda item: (-item[0], item[1]))
        best = None
        if scored:
            best_score, best_pos = scored[0]
            best = (best_score, best_pos, None, None)
        second = scored[1][0] if len(scored) > 1 else None
        if best is None or best[0] < MIN_COLUMN_MATCH:
            unmatched.append({
                "letter": t_col.get("letter"),
                "header": t_col.get("header"),
                "role": t_col.get("semantic_role"),
                "reason": ("no instance column reached the minimum match "
                           f"score {MIN_COLUMN_MATCH}"),
                "best_score": best[0] if best else 0,
                "minimum_score": MIN_COLUMN_MATCH})
            continue
        score, pos, _none, _none2 = best
        i_col = next(col for p2, col, _prep in remaining if p2 == pos)
        _scored, reasons = _pair_score(t_prep, _prep_column(i_col),
                                       want_reasons=True)
        assert _scored == score, "fast scan and evidence score diverged"
        block = decision(score, reasons or ["weak structural similarity"],
                         second=second, kind="column",
                         subject=f"{t_col.get('letter')}->"
                                 f"{i_col.get('letter')}")
        if block["requires_review"]:
            ambiguity.append(block["ambiguity"])
        entry = {"template_letter": t_col.get("letter"),
                 "instance_letter": i_col.get("letter"),
                 "template_header": t_col.get("header"),
                 "instance_header": i_col.get("header"),
                 "semantic_role": t_col.get("semantic_role"),
                 **{k: block[k] for k in
                    ("confidence_score", "confidence_level", "evidence",
                     "ambiguity", "requires_review")}}
        matched.append(entry)
        if t_col.get("value_kind") and i_col.get("value_kind") and \
                t_col["value_kind"] != i_col["value_kind"]:
            conflicts.append({
                "kind": "value_kind",
                "template": {"letter": t_col.get("letter"),
                             "value_kind": t_col.get("value_kind")},
                "instance": {"letter": i_col.get("letter"),
                             "value_kind": i_col.get("value_kind")},
                "message": "value kinds differ; no automatic conversion in "
                           "this phase"})
        if t_col.get("number_format") != i_col.get("number_format"):
            conflicts.append({
                "kind": "number_format",
                "template": {"letter": t_col.get("letter"),
                             "number_format": t_col.get("number_format")},
                "instance": {"letter": i_col.get("letter"),
                             "number_format": i_col.get("number_format")},
                "message": "number formats differ"})
        remaining = [item for item in remaining if item[0] != pos]
    extras = [{"letter": c.get("letter"), "header": c.get("header"),
               "role": c.get("semantic_role")}
              for _pos, c, _prep in remaining]
    ambiguity = ambiguity + _competition_ambiguity(template_cols,
                                                   instance_cols, matched)
    return matched, unmatched, extras, conflicts, ambiguity
def _merge_ranges(value):
    """Merge-cell ranges as a set of strings.

    2A's Document Map reports merges as dicts ({"range": ...}), the saved
    profile as plain range strings; both must compare equal.
    """
    out = set()
    for item in value or []:
        if isinstance(item, str):
            out.add(item)
        elif isinstance(item, dict) and item.get("range"):
            out.add(item["range"])
    return out


def _structural_score(template_sheet, instance_sheet):
    scores, evidence = [], []
    t_merges = _merge_ranges(template_sheet.get("merged_cells"))
    i_merges = _merge_ranges(instance_sheet.get("merged_cells"))
    if t_merges or i_merges:
        common_merges = len(t_merges & i_merges)
        total = len(t_merges | i_merges)
        scores.append(int(100 * common_merges / total) if total else 100)
        evidence.append(f"merged ranges shared {common_merges}/{total}")
    else:
        scores.append(100)
        evidence.append("no merged cells on either side")
    t_tables = sorted((t.get("name"), t.get("ref"))
                      for t in template_sheet.get("tables") or [])
    i_tables = sorted((t.get("name"), t.get("ref"))
                      for t in instance_sheet.get("tables") or [])
    scores.append(100 if t_tables == i_tables else 50)
    evidence.append(f"tables {t_tables} vs {i_tables}")
    t_valid = sorted(v.get("range") for v in template_sheet.get("validations") or [])
    i_valid = sorted(v.get("range") for v in instance_sheet.get("validations") or [])
    scores.append(100 if t_valid == i_valid else 50)
    evidence.append(f"validation ranges {len(t_valid)} vs {len(i_valid)}")
    for key, label in (("freeze_panes", "freeze panes"),
                       ("autofilter", "autofilter")):
        same = template_sheet.get(key) == instance_sheet.get(key)
        scores.append(100 if same else 60)
        evidence.append(f"{label}: {template_sheet.get(key)} vs "
                        f"{instance_sheet.get(key)}")
    t_hidden = (template_sheet.get("hidden") or {})
    i_hidden = (instance_sheet.get("hidden") or {})
    same_hidden = (_hidden_indexes(t_hidden.get("rows")) ==
                   _hidden_indexes(i_hidden.get("rows")) and
                   _hidden_indexes(t_hidden.get("columns")) ==
                   _hidden_indexes(i_hidden.get("columns")))
    scores.append(100 if same_hidden else 70)
    evidence.append("hidden rows/columns "
                    + ("match" if same_hidden else "differ"))
    return (sum(scores) // len(scores) if scores else 100), evidence


def _pair_sheets(template_profile, doc_map):
    template_sheets = template_profile.get("sheets") or []
    instance_sheets = doc_map.get("sheets") or []
    pairs, used = [], set()
    for t_sheet in template_sheets:
        for pos, i_sheet in enumerate(instance_sheets):
            if pos in used:
                continue
            if t_sheet.get("name") == i_sheet.get("name"):
                pairs.append((pos, t_sheet))
                used.add(pos)
                break
    for t_sheet in template_sheets:
        if any(pair[1] is t_sheet for pair in pairs):
            continue
        best = None
        for pos, i_sheet in enumerate(instance_sheets):
            if pos in used:
                continue
            t_headers = {c.get("normalized_header")
                         for c in t_sheet.get("columns") or []}
            i_headers = {c.get("normalized_header")
                         for c in i_sheet.get("columns") or []}
            overlap = len(t_headers & i_headers) if (t_headers or i_headers) else 0
            if best is None or overlap > best[0]:
                best = (overlap, pos)
        if best and best[0] > 0:
            pairs.append((best[1], t_sheet))
            used.add(best[1])
        else:
            pairs.append((None, t_sheet))
    return pairs, [s for pos, s in enumerate(instance_sheets)
                   if pos not in used]


def match_profile(doc_map, template_profile, dictionary, result,
                  instance_meta=None):
    """Compare an instance workbook's map against a saved profile."""
    if template_profile.get("profile_version") != PROFILE_VERSION:
        result.warn("PROFILE_VERSION_DIFF",
                    f"profile_version {template_profile.get('profile_version')!r}"
                    f" differs from {PROFILE_VERSION!r}; matching continues "
                    "but field additions may be reported as unmatched "
                    "(never silent).", profile_version=template_profile.get(
                        "profile_version"))
    pairs, extra_sheets = _pair_sheets(template_profile, doc_map)
    sheet_reports = []
    columns_by_name = {s.get("name"): s for s in doc_map.get("sheets") or []}
    for pos, t_sheet in pairs:
        i_sheet = columns_by_name.get((doc_map.get("sheets") or [])[pos].get("name")) \
            if pos is not None else None
        if i_sheet is None:
            sheet_reports.append({
                "template_sheet": t_sheet.get("name"),
                "instance_sheet": None,
                "pair_score": 0, "pair_level": "LOW",
                "components": {}, "column_matches": [],
                "region_matches": [], "unmatched_template": [
                    {"kind": "sheet", "name": t_sheet.get("name"),
                     "reason": "no instance sheet matched"}],
                "extra_instance": [], "conflicts": [],
                "blank_template_columns": t_sheet.get("blank_columns"),
                "blank_instance_columns": None,
                "instance_columns_cap": None,
                "evidence": ["template sheet has no instance counterpart"],
                "ambiguity": [], "requires_review": True})
            continue
        region_matches, unmatched_regions, extra_regions, reg_amb = \
            _region_match(t_sheet.get("regions") or [],
                          i_sheet.get("regions") or [])
        i_columns_all, i_blank_meta = blank_columns_meta(
            i_sheet.get("columns") or [], i_sheet.get("name"), result,
            where="match")
        i_columns, i_columns_cap = cap_list(
            i_columns_all, MATCH_COLUMN_CAP, result,
            "INSTANCE_COLUMNS_TRUNCATED",
            f"{i_sheet.get('name')}: instance columns scored during matching",
            sheet=i_sheet.get("name"))
        column_matches, unmatched_columns, extra_columns, conflicts, col_amb = \
            _column_match(t_sheet.get("columns") or [], i_columns)
        region_score = 100 if not (t_sheet.get("regions") or []) else \
            int(100 * len(region_matches) /
                max(1, len(t_sheet.get("regions") or [])))
        column_score = 100 if not (t_sheet.get("columns") or []) else \
            int(100 * len(column_matches) /
                max(1, len(t_sheet.get("columns") or [])))  # material only
        formula_score, formula_evidence = _kinds_similarity(
            (t_sheet.get("formula_summary") or {}).get("kinds"),
            (i_sheet.get("formula_summary") or {}).get("kinds"))
        structural_score, structural_evidence = _structural_score(t_sheet,
                                                                  i_sheet)
        components = {"region_score": region_score,
                      "column_score": column_score,
                      "formula_score": formula_score,
                      "structural_score": structural_score}
        overall = int(round(sum(components[k] * w for k, w in WEIGHTS.items())))
        evidence = [f"region_score:{region_score}", f"column_score:{column_score}",
                    f"formula_score:{formula_score}",
                    f"structural_score:{structural_score}",
                    formula_evidence]
        block = decision(overall, evidence, kind="match",
                         subject=f"{t_sheet.get('name')} -> "
                                 f"{i_sheet.get('name')}")
        ambiguity = list(col_amb)
        if block["requires_review"]:
            result.warn("LOW_CONFIDENCE_DETECTION",
                        f"match {t_sheet.get('name')!r}: overall score "
                        f"{overall} ({block['confidence_level']}); review "
                        "required.", sheet=t_sheet.get("name"), score=overall)
        for entry in region_matches:
            if entry.get("requires_review"):
                ambiguity.append(entry["ambiguity"])
        sheet_reports.append({
            "template_sheet": t_sheet.get("name"),
            "instance_sheet": i_sheet.get("name"),
            "blank_template_columns": t_sheet.get("blank_columns"),
            "blank_instance_columns": i_blank_meta,
            "instance_columns_cap": i_columns_cap,
            "components": components,
            "pair_score": overall,
            "pair_level": block["confidence_level"],
            "column_matches": column_matches,
            "region_matches": region_matches,
            "unmatched_template": unmatched_columns + unmatched_regions,
            "extra_instance": extra_columns + extra_regions,
            "conflicts": conflicts,
            "evidence": block["evidence"] + structural_evidence,
            "ambiguity": ambiguity,
            "requires_review": block["requires_review"] or bool(ambiguity),
        })
    all_columns_unmatched = [dict(item, sheet=rep["template_sheet"])
                             for rep in sheet_reports
                             for item in rep["unmatched_template"]
                             if item.get("kind") != "sheet"]
    columns_listed, columns_meta = cap_list(
        all_columns_unmatched, MATCH_LIST_CAP, result, "UNMATCHED_TRUNCATED",
        "unmatched template columns")
    unmatched_sheets = [dict(item, sheet=rep["template_sheet"])
                        for rep in sheet_reports
                        for item in rep["unmatched_template"]
                        if item.get("kind") == "sheet"]
    unmatched_sheets += [{"kind": "sheet", "name": sheet,
                          "reason": "template sheet missing on instance side"}
                         for sheet in extra_sheets]
    extras = [dict(item, sheet=rep["instance_sheet"])
              for rep in sheet_reports for item in rep["extra_instance"]]
    conflicts = [dict(item, sheet=rep["instance_sheet"])
                 for rep in sheet_reports for item in rep["conflicts"]]
    if sheet_reports:
        overall = int(round(sum(rep["pair_score"] for rep in sheet_reports) /
                            len(sheet_reports)))
    else:
        overall = 0
        result.warn("MATCH_NO_SHEETS",
                    "no template sheet could be paired with an instance "
                    "sheet (never silent).")
    sheet_listed, sheet_meta = cap_list(
        sheet_reports, MATCH_LIST_CAP, result, "MATCH_SHEETS_TRUNCATED",
        "sheet match reports")
    report = {
        "match_report_version": SEMANTICS_VERSION,
        "generator": f"xlsx_semantics.py {SEMANTICS_VERSION}",
        "profile_id": template_profile.get("profile_id"),
        "profile_hash": template_profile.get("profile_hash"),
        "dictionary_version": dictionary["version"],
        "instance": {
            "structure_fingerprint": (doc_map.get("workbook") or {}).get(
                "structure_fingerprint"),
            "sheet_count": (doc_map.get("workbook") or {}).get("sheet_count"),
            "bytes_sha256": (instance_meta or {}).get("sha256"),
            "size": (instance_meta or {}).get("size"),
        },
        "template_content_identity": template_profile.get("content_identity"),
        "same_content_identity": (
            template_profile.get("content_identity") ==
            content_identity(doc_map, dictionary["version"])),
        "overall_score": overall,
        "overall_level": level_for(overall),
        "components": {
            "region_score": (sum(r["components"].get("region_score", 0)
                                 for r in sheet_reports) //
                             len(sheet_reports)) if sheet_reports else 0,
            "column_score": (sum(r["components"].get("column_score", 0)
                                 for r in sheet_reports) //
                             len(sheet_reports)) if sheet_reports else 0,
            "formula_score": (sum(r["components"].get("formula_score", 0)
                                  for r in sheet_reports) //
                              len(sheet_reports)) if sheet_reports else 0,
            "structural_score": (sum(r["components"].get("structural_score", 0)
                                     for r in sheet_reports) //
                                 len(sheet_reports)) if sheet_reports else 0,
        },
        "sheets": sheet_listed,
        "sheets_cap": sheet_meta,
        "unmatched_template": unmatched_sheets + columns_listed,
        "unmatched_template_cap": columns_meta,
        "extra_instance": extras,
        "conflicts": conflicts,
        "requires_review": any(rep["requires_review"] for rep in sheet_reports),
    }
    return report

# ---------------------------------------------------------------------------
# instance augmentation + CLI
# ---------------------------------------------------------------------------

def augment_instance(doc_map, semantics):
    """Attach B1 annotations to a Document Map's columns (no mutation)."""
    by_name = {s.get("name"): s for s in semantics.get("sheets") or []}
    sheets = []
    for sheet in doc_map.get("sheets") or []:
        sem = by_name.get(sheet.get("name")) or {"columns": []}
        sem_by_letter = {c.get("letter"): c for c in sem.get("columns") or []}
        columns = []
        for col in sheet.get("columns") or []:
            extra = sem_by_letter.get(col.get("letter")) or {}
            merged = dict(col)
            merged["semantic_role"] = extra.get("semantic_role")
            merged["role_level"] = extra.get("role_level")
            merged["value_kind"] = extra.get("value_kind")
            merged["unit_guess"] = extra.get("unit_guess")
            columns.append(merged)
        sheets.append(dict(sheet, columns=columns))
    out = dict(doc_map)
    out["sheets"] = sheets
    return out


REVIEW_STATES = ("accepted", "rejected", "pending")


def _review_subjects(report):
    """Every ambiguity subject in a match report, as stable sortable keys."""
    keys = []
    for rep in report.get("sheets") or []:
        for group in rep.get("ambiguity") or []:
            entries = group if isinstance(group, list) else [group]
            for entry in entries:
                if isinstance(entry, dict):
                    keys.append(f"{entry.get('kind')}|{entry.get('subject')}")
    for entry in report.get("ambiguity") or []:
        if isinstance(entry, dict):
            keys.append(f"{entry.get('kind')}|{entry.get('subject')}")
    return sorted(set(keys))


def load_review(path, result):
    """Load human review decisions (data in; detection is never changed)."""
    target = Path(path)
    if not target.exists():
        raise common.XlsxError(
            "FILE_NOT_FOUND", f"review file not found: {target}",
            recovery="Create one from the match report's "
                     "review_state.subjects list.",
            context={"review": str(target)})
    try:
        doc = json.loads(target.read_text(encoding="utf-8"))
    except ValueError as exc:
        raise common.XlsxError(
            "VALIDATION_FAILED", f"review file is not valid JSON: {exc}",
            recovery='Use {"decisions": {"<kind>|<subject>": {"state": '
                     '"accepted", "note": "..."}}}.',
            context={"review": str(target)}) from exc
    decisions = doc.get("decisions") if isinstance(doc, dict) else None
    if not isinstance(decisions, dict):
        raise common.XlsxError(
            "VALIDATION_FAILED",
            "review file needs a 'decisions' object keyed by subject.",
            recovery='{"decisions": {"column|A->B": {"state": "accepted"}}}',
            context={"review": str(target)})
    out = {}
    for key, value in decisions.items():
        state = value.get("state") if isinstance(value, dict) else value
        if state not in REVIEW_STATES:
            raise common.XlsxError(
                "SPEC_INVALID",
                f"review state {state!r} for {key!r} is not one of "
                f"{REVIEW_STATES}.",
                recovery="Use accepted / rejected / pending.",
                context={"subject": key, "state": state})
        out[str(key)] = {
            "state": state,
            "note": (value.get("note") if isinstance(value, dict) else None)}
    return out


def attach_review(report, review, result):
    """Record the human decision next to the machine decision.

    The machine result (scores, levels, ambiguity) is never rewritten; the
    review only states what a human decided about it.
    """
    subjects = _review_subjects(report)
    review = review or {}
    unknown = sorted(set(review) - set(subjects))
    for key in unknown:
        result.warn(
            "UNKNOWN_REVIEW_SUBJECT",
            f"review decision {key!r} does not match any ambiguity subject "
            "in this report (never silent).", subject=key)
    entries = []
    for key in subjects:
        decided = review.get(key)
        entries.append({"subject": key,
                        "state": decided["state"] if decided else "unreviewed",
                        "note": decided.get("note") if decided else None})
    states = {entry["state"] for entry in entries}
    if not entries:
        state = "nothing_to_review"
    elif states == {"unreviewed"}:
        state = "unreviewed"
    elif "unreviewed" in states or "pending" in states:
        state = "partially_reviewed"
    else:
        state = "reviewed"
    report["review_state"] = {"state": state, "subjects": entries,
                              "unknown_review_subjects": unknown}
    if review and state != "reviewed":
        pending = sum(1 for entry in entries
                      if entry["state"] == "unreviewed")
        result.warn("REVIEW_INCOMPLETE",
                    f"review state is {state}: {pending} subject(s) still "
                    "unreviewed (never silent).")
    return report


def load_profile(path, result):
    """Load a saved 2B template profile."""
    target = Path(path)
    if not target.exists():
        raise common.XlsxError(
            "FILE_NOT_FOUND", f"template profile not found: {target}",
            recovery="Create one with `xlsx_semantics.py BOOK.xlsx --emit "
                     "profile --out book.profile.json`.",
            context={"profile": str(target)})
    try:
        profile = json.loads(target.read_text(encoding="utf-8"))
    except ValueError as exc:
        raise common.XlsxError(
            "VALIDATION_FAILED", f"profile is not valid JSON: {exc}",
            recovery="Re-generate the profile.",
            context={"profile": str(target)}) from exc
    if not isinstance(profile, dict) or "sheets" not in profile:
        raise common.XlsxError(
            "VALIDATION_FAILED",
            "profile JSON needs 'profile_version' and a 'sheets' array.",
            recovery="Re-generate the profile with --emit profile.",
            context={"profile": str(target),
                     "keys": sorted(profile) if isinstance(profile, dict)
                     else None})
    return profile


def write_out(path, payload):
    """Write a deterministic JSON artifact (never a workbook)."""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(common.canonical_json(payload) + "\n", encoding="utf-8")
    return target


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="FAZ 2B: semantic labeling, template profile and "
                    "instance<->template matching. Read-only: consumes the "
                    "FAZ 2A Document Map, never opens or writes a workbook.")
    parser.add_argument("file", nargs="?",
                        help="workbook (.xlsx/.xlsm) or a saved Document "
                             "Map (.json)")
    parser.add_argument("--doc-map", metavar="PATH",
                        help="use a saved 2A Document Map JSON instead of "
                             "running 2A on a workbook")
    parser.add_argument("--emit", choices=("semantics", "profile"),
                        default="semantics",
                        help="what to emit (default: semantics)")
    parser.add_argument("--match", metavar="PROFILE.json",
                        help="match the instance against a saved profile")
    parser.add_argument("--out", metavar="PATH",
                        help="also write the emitted model to this path")
    parser.add_argument("--pretty", action="store_true",
                        help="indented JSON instead of compact canonical JSON")
    parser.add_argument("--samples", type=int,
                        default=understand.SAMPLE_DEFAULT, metavar="N",
                        help="passed through to 2A (default "
                             f"{understand.SAMPLE_DEFAULT}, max "
                             f"{understand.SAMPLE_MAX})")
    parser.add_argument("--formula-list-cap", type=int,
                        default=understand.FORMULA_LIST_CAP, metavar="N",
                        help="passed through to 2A (default "
                             f"{understand.FORMULA_LIST_CAP}; 0 = summary "
                             "only)")
    parser.add_argument("--dictionary", metavar="PATH",
                        help="override references/semantic-dictionary."
                             "tr-en.json")
    parser.add_argument("--review-file", metavar="PATH",
                        help="human review decisions for --match "
                             "(recorded in match_report.review_state; "
                             "detection is never changed)")
    args = parser.parse_args(argv)

    if not args.file and not args.doc_map:
        raise common.XlsxError(
            "SPEC_INVALID", "no input given.",
            recovery="Pass a workbook path, a Document Map JSON, or "
                     "--doc-map PATH.",
            context={"file": args.file, "doc_map": args.doc_map})
    if args.match and args.emit == "profile":
        raise common.XlsxError(
            "SPEC_INVALID", "--match and --emit profile cannot be combined.",
            recovery="Use `--match PROFILE.json` alone, or `--emit profile`.",
            context={"emit": args.emit, "match": args.match})
    if args.samples < 0:
        raise common.XlsxError("SPEC_INVALID", "--samples must be >= 0.",
                               recovery="Pass a non-negative integer.",
                               context={"samples": args.samples})
    if args.formula_list_cap < 0:
        raise common.XlsxError("SPEC_INVALID",
                               "--formula-list-cap must be >= 0.",
                               recovery="Pass a non-negative integer.",
                               context={"formula_list_cap":
                                        args.formula_list_cap})

    result = common.Result()
    dictionary = load_dictionary(args.dictionary)
    for conflict in dictionary["conflicts"]:
        result.warn("DICTIONARY_TERM_CONFLICT",
                    f"dictionary term {conflict['term']!r} maps to several "
                    f"roles {conflict['roles']}; the first wins and the "
                    "dictionary must be fixed (never silent).",
                    **conflict)

    profile = load_profile(args.match, result) if args.match else None
    if args.review_file and not args.match:
        raise common.XlsxError(
            "SPEC_INVALID",
            "--review-file is only meaningful together with --match.",
            recovery="Pass `--match PROFILE.json --review-file review.json`.",
            context={"review_file": args.review_file})
    review = load_review(args.review_file, result) if args.review_file else None
    source_meta = None
    if args.doc_map:
        doc_map = load_doc_map(args.doc_map, result)
        result.set(doc_map_source=str(Path(args.doc_map)))
    elif args.file and Path(args.file).suffix.casefold() == ".json":
        doc_map = load_doc_map(args.file, result)
        result.set(doc_map_source=str(Path(args.file)))
    else:
        samples = args.samples
        if samples > understand.SAMPLE_MAX:
            result.warn("SAMPLES_CLAMPED",
                        f"--samples {samples} clamped to "
                        f"{understand.SAMPLE_MAX} (never silent).",
                        requested=samples, applied=understand.SAMPLE_MAX)
            samples = understand.SAMPLE_MAX
        doc_map = build_map_from_workbook(args.file, result, samples,
                                          args.formula_list_cap)
        source_meta = common.sha256_file(args.file)

    result.set(mode="semantics", file=doc_map.get("file") or args.doc_map
               or args.file)
    semantics = label_workbook(doc_map, dictionary, result)

    if args.match:
        augmented = augment_instance(doc_map, semantics)
        report = match_profile(augmented, profile, dictionary, result,
                               instance_meta=source_meta)
        report = attach_review(report, review, result)
        payload = {"ok": True, "mode": "semantics", "command": "match",
                   "dictionary_version": dictionary["version"],
                   "match_report": report}
        if args.out:
            write_out(args.out, report)
            payload["out"] = str(Path(args.out))
        return emit(payload, result, args.pretty)

    if args.emit == "profile":
        profile = build_profile(doc_map, semantics, dictionary, result,
                                source_meta)
        payload = {"ok": True, "mode": "semantics", "command": "profile",
                   "profile_version": profile["profile_version"],
                   "profile_id": profile["profile_id"],
                   "profile_hash": profile["profile_hash"],
                   "out": None, "profile": profile}
        if args.out:
            write_out(args.out, profile)
            payload["out"] = str(Path(args.out))
        return emit(payload, result, args.pretty)

    payload = {"ok": True, "mode": "semantics", "command": "semantics",
               "semantics_version": SEMANTICS_VERSION,
               "identity": profile_identity(doc_map, dictionary["version"]),
               "semantics": semantics}
    if args.out:
        write_out(args.out, semantics)
        payload["out"] = str(Path(args.out))
    return emit(payload, result, args.pretty)


if __name__ == "__main__":
    sys.exit(common.guard(main)())
