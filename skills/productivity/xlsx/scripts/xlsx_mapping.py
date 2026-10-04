#!/usr/bin/env python3
"""FAZ 2B -- mapping plan + dry-run fill plan (read-only, plan only).

Turns a saved FAZ 2B template profile plus a source table (CSV/JSON)
into a deterministic MAPPING PLAN, and can then simulate -- without
writing anything -- what a fill would touch on the target workbook
(B6 dry-run). Reading the target workbook goes through FAZ 2A's
Document Map; this module never opens a workbook itself and never
writes one.

Commands:
    xlsx_mapping.py --profile p.json --source data.csv --emit plan
    xlsx_mapping.py --profile p.json --plan plan.json --target B.xlsx \
                    --emit dry-run
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_common as common          # noqa: E402
import xlsx_semantics as semantics    # noqa: E402
import xlsx_understand as understand  # noqa: E402

PLAN_VERSION = "2b.1"
DRY_RUN_VERSION = "2b.1"

PLAN_LIST_CAP = 2000
DRY_RUN_LIST_CAP = 5000

MIN_MATCH_SCORE = 30          # below this a source key stays unresolved
DEFAULT_POLICY = {
    "on_low_confidence": "report_only",
    "on_conflict": "block",
    "max_rows": None,
    "missing_required": "report_only",
    # FAZ 3B opt-in: authorizes the executor to perform the plan's
    # evidenced row expansion (adds rows). Default off -- plans stay
    # planned-only until expansion is asked for explicitly.
    "allow_expansion": False,
    # FAZ 3B opt-in: authorizes lookup execution. Default off.
    "allow_lookup": False,
    # FAZ 3C opt-in (decision D3): when idempotency evidence is missing
    # (manifest absent but the workspace has history, unreadable, corrupt),
    # writing into a POPULATED target region requires this explicit
    # acknowledgement. Default off -- the executor fails closed instead.
    "allow_unverified_region_overwrite": False,
}

SOURCE_DATE_PATTERNS = (
    re.compile(r"^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2})?)?$"),
    re.compile(r"^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$"),
)
SOURCE_BOOL_WORDS = {"true": True, "false": False, "evet": True,
                     "hayir": False, "hayır": False, "yes": True,
                     "no": False}


# ---------------------------------------------------------------------------
# source table loading (CSV / JSON) -- value inference, no invention
# ---------------------------------------------------------------------------

def _infer_scalar(text):
    """Very small, deterministic scalar inference for source tables."""
    if text is None:
        return "blank", None
    if isinstance(text, bool):
        return "boolean", text
    if isinstance(text, int):
        return "integer", text
    if isinstance(text, float):
        return "decimal", text
    stripped = str(text).strip()
    if not stripped:
        return "blank", None
    low = stripped.casefold()
    if low in SOURCE_BOOL_WORDS:
        return "boolean", SOURCE_BOOL_WORDS[low]
    body = stripped.replace(" ", "")
    if re.fullmatch(r"[+-]?\d+", body):
        try:
            return "integer", int(body)
        except ValueError:
            pass
    number = body
    if number.count(",") and number.count("."):
        if number.rfind(",") > number.rfind("."):
            number = number.replace(".", "").replace(",", ".")
        else:
            number = number.replace(",", "")
    elif number.count(","):
        number = number.replace(",", ".")
    if re.fullmatch(r"[+-]?\d+(\.\d+)?", number):
        try:
            return "decimal", float(number)
        except ValueError:
            pass
    for pattern in SOURCE_DATE_PATTERNS:
        if pattern.match(stripped):
            return "date", stripped
    return "string", stripped


def _sniff_delimiter(sample):
    try:
        return csv.Sniffer().sniff(sample, delimiters=";,\t|").delimiter
    except csv.Error:
        counts = {d: sample.count(d) for d in (";", ",", "\t", "|")}
        best = max(counts, key=lambda key: counts[key])
        return best if counts[best] else ","


def load_source(path, result):
    """Load a CSV/JSON source table; every fallback is reported."""
    target = Path(path)
    if not target.exists():
        raise common.XlsxError(
            "FILE_NOT_FOUND", f"source data not found: {target}",
            recovery="Pass a CSV or JSON table path as --source.",
            context={"source": str(target)})
    suffix = target.suffix.casefold()
    headers, rows, meta = [], [], {"kind": suffix.lstrip(".")}
    if suffix == ".json":
        try:
            payload = json.loads(target.read_text(encoding="utf-8"))
        except (ValueError, UnicodeDecodeError) as exc:
            raise common.XlsxError(
                "VALIDATION_FAILED", f"source JSON is not readable: {exc}",
                recovery="Provide UTF-8 JSON: a list of objects, or "
                         "{\"rows\": [...]}.",
                context={"source": str(target)}) from exc
        if isinstance(payload, dict) and isinstance(payload.get("rows"), list):
            payload = payload["rows"]
        if not isinstance(payload, list) or not payload:
            raise common.XlsxError(
                "VALIDATION_FAILED",
                "source JSON must be a non-empty list of objects (or "
                "{\"rows\": [...]}).",
                recovery="See --help for accepted shapes.",
                context={"source": str(target)})
        if all(isinstance(row, dict) for row in payload):
            for row in payload:
                for key in row:
                    if key not in headers:
                        headers.append(key)
            rows = [[row.get(key) for key in headers] for row in payload]
        elif all(isinstance(row, list) for row in payload):
            headers = [str(cell) for cell in payload[0]]
            rows = [list(row) for row in payload[1:]]
        else:
            raise common.XlsxError(
                "VALIDATION_FAILED",
                "source JSON rows must be all objects or all arrays.",
                recovery="Normalise the file; nothing is guessed.",
                context={"source": str(target)})
    else:
        raw = target.read_bytes()
        text, encoding = None, None
        for candidate in ("utf-8-sig", "utf-8", "cp1254", "latin-1"):
            try:
                text = raw.decode(candidate)
                encoding = candidate
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise common.XlsxError(
                "VALIDATION_FAILED", "source file could not be decoded.",
                recovery="Re-save the CSV as UTF-8.",
                context={"source": str(target)})
        if encoding not in ("utf-8-sig", "utf-8"):
            result.warn("SOURCE_ENCODING_FALLBACK",
                        f"source decoded as {encoding}; UTF-8 failed "
                        "(never silent).",
                        source=str(target), encoding=encoding)
        delimiter = _sniff_delimiter(text[:4096])
        meta["delimiter"] = delimiter
        reader = csv.reader(io.StringIO(text), delimiter=delimiter)
        parsed = [row for row in reader if any(str(cell).strip() for cell in row)]
        if not parsed:
            raise common.XlsxError(
                "VALIDATION_FAILED", "source CSV has no data rows.",
                recovery="Check the delimiter/encoding of the file.",
                context={"source": str(target), "delimiter": delimiter})
        headers = [str(cell).strip() for cell in parsed[0]]
        rows = parsed[1:]
    meta["encoding"] = meta.get("encoding") or "utf-8"
    meta["row_count"] = len(rows)
    meta["headers"] = headers
    columns = []
    width = len(headers)
    for index in range(width):
        values = [row[index] if index < len(row) else None for row in rows]
        kinds = []
        samples = []
        nulls = 0
        for value in values:
            kind, converted = _infer_scalar(value)
            if kind == "blank":
                nulls += 1
                continue
            kinds.append(kind)
            if len(samples) < understand.SAMPLE_DEFAULT and \
                    converted not in samples:
                samples.append(converted)
        distinct = len({str(value) for value in values if str(value) != ""})
        if not kinds:
            kind = "blank"
        elif len(set(kinds)) == 1:
            kind = kinds[0]
        elif set(kinds) <= {"integer", "decimal"}:
            kind = "decimal"
        else:
            kind = "mixed"
        columns.append({
            "key": headers[index],
            "index": index,
            "value_kind": kind,
            "null_count": nulls,
            "distinct_count": distinct,
            "cell_count": len(values),
            "sample_values": samples,
        })
    meta["columns"] = columns
    identity = common.sha256_text(common.canonical_json(
        {"headers": headers, "rows": rows}))
    meta["content_fingerprint"] = identity
    return {"headers": headers, "rows": rows, "meta": meta}


def source_semantics(source, dictionary, result):
    """Reuse B1 labeling on the source columns (no second implementation)."""
    annotated = []
    sheet_stub = {"name": "(source)", "validations": []}
    for column in source["meta"]["columns"]:
        synth = {
            "index": column["index"] + 1,
            "letter": None,
            "header": column["key"],
            "normalized_header": semantics.normalize_term(column["key"]),
            "data_type": semantics.VALUE_KIND_BY_DATA_TYPE.get(
                column["value_kind"], column["value_kind"]),
            "number_format": None,
            "cell_count": column["cell_count"],
            "distinct_count": column["distinct_count"],
            "null_count": column["null_count"],
            "formula_count": 0,
            "formula_share": 0.0,
            "region": None,
        }
        labeled = semantics.label_column(synth, sheet_stub, dictionary,
                                        {}, result, "(source)")
        labeled["key"] = column["key"]
        labeled["source_index"] = column["index"]
        annotated.append(labeled)
    return annotated


# ---------------------------------------------------------------------------
# policy + plan identity
# ---------------------------------------------------------------------------

def read_policy(path, result):
    policy = dict(DEFAULT_POLICY)
    if not path:
        return policy
    target = Path(path)
    if not target.exists():
        raise common.XlsxError(
            "FILE_NOT_FOUND", f"policy file not found: {target}",
            recovery="Pass --policy-file JSON or omit it for defaults.",
            context={"policy": str(target)})
    try:
        loaded = json.loads(target.read_text(encoding="utf-8"))
    except ValueError as exc:
        raise common.XlsxError(
            "VALIDATION_FAILED", f"policy is not valid JSON: {exc}",
            recovery="See SKILL.md (Semantic Layer, Profiles & Mapping) for the keys.",
            context={"policy": str(target)}) from exc
    unknown = sorted(set(loaded) - set(DEFAULT_POLICY))
    if unknown:
        raise common.XlsxError(
            "SPEC_UNKNOWN_KEY",
            f"unknown policy key(s): {unknown}",
            recovery=f"Allowed keys: {sorted(DEFAULT_POLICY)}.",
            context={"policy": str(target), "unknown": unknown})
    policy.update(loaded)
    if policy["on_low_confidence"] not in ("report_only", "block"):
        raise common.XlsxError(
            "SPEC_INVALID",
            "on_low_confidence must be 'report_only' or 'block'.",
            recovery="Fix the policy file.",
            context={"policy": str(target)})
    if policy["on_conflict"] not in ("report", "block"):
        raise common.XlsxError(
            "SPEC_INVALID", "on_conflict must be 'report' or 'block'.",
            recovery="Fix the policy file.",
            context={"policy": str(target)})
    if policy["max_rows"] is not None and not isinstance(policy["max_rows"],
                                                         int):
        raise common.XlsxError(
            "SPEC_INVALID", "max_rows must be an integer or null.",
            recovery="Fix the policy file.",
            context={"policy": str(target)})
    return policy


def plan_identity(profile, source, target_structure, policy):
    payload = {
        "plan_version": PLAN_VERSION,
        "profile_id": profile.get("profile_id"),
        "source_content_fingerprint": source["meta"]["content_fingerprint"],
        "target_structure_fingerprint": target_structure,
        "policy": policy,
    }
    return common.sha256_text(common.canonical_json(payload))

# ---------------------------------------------------------------------------
# B5 -- mapping plan (source table -> template profile targets)
# ---------------------------------------------------------------------------

NUMERIC_KINDS = {"integer", "decimal"}
DATE_KINDS = {"date", "datetime"}


def _column_span_of_ref(ref):
    bounds = semantics.range_columns(ref)
    if not bounds:
        return []
    return list(range(bounds[0], bounds[1] + 1))


def collect_targets(profile):
    """Flatten a profile into addressable target columns (deterministic)."""
    targets = []
    for sheet in profile.get("sheets") or []:
        in_block = {}
        anchor = {}
        for block in sheet.get("blocks") or []:
            if block.get("archetype") != "record_block" or not block.get("rows"):
                continue
            for index in _column_span_of_ref(block.get("range")):
                in_block[index] = block.get("block_id")
                anchor[index] = block["rows"][0]
        slots = {}
        for slot in sheet.get("slots") or []:
            if not slot.get("range"):
                continue
            bounds = semantics.range_columns(slot.get("range"))
            if not bounds:
                continue
            match = semantics.RANGE_RE.match(slot["range"])
            row = int(match.group(2)) if match else None
            last = int(match.group(4) or match.group(2)) if match else row
            span = (abs(last - row) + 1) if (row is not None and
                                             last is not None) else 1
            for index in range(bounds[0], bounds[1] + 1):
                slots[index] = (slot.get("range"), row, span)
        for column in sheet.get("columns") or []:
            index = column.get("index")
            slot_range, slot_row, slot_span = slots.get(index,
                                                        (None, None, None))
            if index in in_block:
                mode_hint, anchor_row = "row_block", anchor.get(index)
            elif slot_range and (slot_span or 1) > 1:
                mode_hint, anchor_row = "row_block", slot_row
            elif slot_range:
                mode_hint, anchor_row = "scalar", slot_row
            else:
                mode_hint, anchor_row = "scalar", None
            targets.append({
                "sheet": sheet.get("name"),
                "letter": column.get("letter"),
                "index": index,
                "header": column.get("header"),
                "normalized_header": column.get("normalized_header"),
                "semantic_role": column.get("semantic_role"),
                "value_kind": column.get("value_kind"),
                "number_format": column.get("number_format"),
                "unit_guess": column.get("unit_guess"),
                "required": column.get("required"),
                "required_source": column.get("required_source"),
                "region": column.get("region"),
                "mode_hint": mode_hint,
                "anchor_row": anchor_row,
                "block_id": in_block.get(index),
                "slot_range": slot_range,
                "slot_rows": slot_span,
                "number_formats": [v.get("range")
                                   for v in sheet.get("validations") or []],
            })
    return targets


def _compat_of(source_kind, target_kind):
    if source_kind == target_kind:
        return "exact", None
    if {source_kind, target_kind} <= NUMERIC_KINDS:
        return "numeric_family", "numeric_family_conversion_not_performed"
    if {source_kind, target_kind} <= DATE_KINDS:
        return "date_family", "date_family_conversion_not_performed"
    if "blank" in (source_kind, target_kind):
        return "blank_tolerant", None
    return "mismatch", f"{source_kind} -> {target_kind}"


def _score_pair(source_col, target):
    score, reasons, notes = 0, [], []
    source_norm = source_col.get("normalized_header")
    target_norm = target.get("normalized_header")
    if source_norm and source_norm == target_norm:
        score += 55
        reasons.append(f"header_exact:{source_norm}")
    elif source_norm and target_norm:
        source_tokens = {t for t in source_norm.replace("/", " ").split() if t}
        target_tokens = {t for t in target_norm.replace("/", " ").split() if t}
        shared = source_tokens & target_tokens
        if shared:
            score += int(30 * len(shared) /
                         max(len(source_tokens), len(target_tokens)))
            reasons.append(f"header_tokens:{','.join(sorted(shared))}")
    role = source_col.get("semantic_role")
    if role and role != "unknown" and role == target.get("semantic_role"):
        score += 25
        reasons.append(f"role:{role}")
    compat, note = _compat_of(source_col.get("value_kind"),
                              target.get("value_kind"))
    if compat == "exact":
        score += 12
        reasons.append(f"value_kind:{source_col.get('value_kind')}")
    elif note:
        notes.append(note)
    if source_col.get("unit_guess") and \
            source_col.get("unit_guess") == target.get("unit_guess"):
        score += 8
        reasons.append(f"unit:{source_col.get('unit_guess')}")
    if target.get("required"):
        score += 5
        reasons.append("target:required(heuristic)")
    return min(100, score), reasons, notes, compat


def build_plan(profile, source, source_cols, policy, result,
               target_structure=None):
    """Deterministic source->target mapping plan (writes nothing)."""
    targets = collect_targets(profile)
    mappings, unresolved, conflicts, notes = [], [], [], []
    claimed = {}
    used_targets = set()
    for source_col in source_cols:
        scored = []
        for target in targets:
            score, reasons, pair_notes, compat = _score_pair(source_col, target)
            scored.append((score, target, reasons, pair_notes, compat))
        scored.sort(key=lambda item: (-item[0], str(item[1].get("sheet")),
                                      str(item[1].get("letter"))))
        best = scored[0] if scored else None
        second = scored[1][0] if len(scored) > 1 else None
        if best is None or best[0] < MIN_MATCH_SCORE:
            unresolved.append({
                "source_key": source_col.get("key"),
                "reason": "no target column reached the minimum score",
                "best_score": best[0] if best else 0,
                "minimum_score": MIN_MATCH_SCORE,
            })
            continue
        chosen = None
        skipped = []
        for candidate in scored:
            candidate_key = (f"{candidate[1].get('sheet')}!"
                             f"{candidate[1].get('letter')}")
            if candidate_key in claimed:
                skipped.append({"target": candidate_key,
                                "claimed_by": claimed[candidate_key]})
                continue
            if candidate[0] < MIN_MATCH_SCORE:
                break
            chosen = candidate
            break
        if chosen is None:
            unresolved.append({
                "source_key": source_col.get("key"),
                "reason": ("every candidate target is already claimed by "
                           "another source key" if skipped else
                           "no unclaimed target column reached the minimum "
                           "score"),
                "best_score": best[0],
                "minimum_score": MIN_MATCH_SCORE,
                "skipped_candidates": skipped,
            })
            continue
        score, target, reasons, pair_notes, compat = chosen
        decision = semantics.decision(score, reasons, second=second,
                                      kind="mapping",
                                      subject=f"{source_col.get('key')} -> "
                                              f"{target.get('sheet')}!"
                                              f"{target.get('letter')}")
        if decision["confidence_level"] == "LOW":
            result.warn("LOW_CONFIDENCE_DETECTION",
                        f"mapping {source_col.get('key')!r} -> "
                        f"{target.get('sheet')}!{target.get('letter')} scored "
                        f"{score} (LOW); review required.",
                        source_key=source_col.get("key"), score=score)
        target_key = f"{target.get('sheet')}!{target.get('letter')}"
        if skipped:
            conflicts.append({
                "kind": "target_reassigned",
                "target": target_key,
                "source_key": source_col.get("key"),
                "skipped": skipped,
                "message": "a higher-scoring target was already claimed by "
                           "another source key; this mapping used an "
                           "alternate target instead (reported, never "
                           "silent)",
            })
        claimed[target_key] = source_col.get("key")
        used_targets.add(target_key)
        if compat == "mismatch":
            conflicts.append({
                "kind": "value_kind_mismatch",
                "target": target_key,
                "source_key": source_col.get("key"),
                "source_value_kind": source_col.get("value_kind"),
                "target_value_kind": target.get("value_kind"),
                "message": "value kinds are incompatible; no automatic "
                           "conversion exists in this phase",
            })
        precondition = []
        if target.get("required"):
            precondition.append("required target: empty source values stay "
                                "unresolved (never invented)")
        if compat == "numeric_family":
            precondition.append("numeric family conversion is NOT performed")
        if compat == "date_family":
            precondition.append("date family conversion is NOT performed")
        if target.get("mode_hint") == "scalar" and target.get("anchor_row") is None:
            precondition.append("no evidenced slot/anchor for this target")
        mappings.append({
            "source_key": source_col.get("key"),
            "target": {
                "sheet": target.get("sheet"),
                "letter": target.get("letter"),
                "anchor": (f"{target.get('letter')}{target.get('anchor_row')}"
                           if target.get("anchor_row") else None),
                "mode": target.get("mode_hint"),
                "block_id": target.get("block_id"),
                "slot_range": target.get("slot_range"),
            },
            "mode": target.get("mode_hint"),
            "confidence": score,
            "level": decision["confidence_level"],
            "confidence_score": decision["confidence_score"],
            "confidence_level": decision["confidence_level"],
            "evidence": decision["evidence"],
            "ambiguity": decision["ambiguity"],
            "requires_review": decision["requires_review"],
            "preconditions": precondition,
            "source_value_kind": source_col.get("value_kind"),
            "target_value_kind": target.get("value_kind"),
            "source_role": source_col.get("semantic_role"),
            "target_role": target.get("semantic_role"),
            "compatibility": compat,
            "notes": pair_notes,
        })
        notes.extend(pair_notes)
    for target in targets:
        if not target.get("required"):
            continue
        target_key = f"{target.get('sheet')}!{target.get('letter')}"
        if target_key not in used_targets:
            unresolved.append({
                "target": target_key,
                "reason": "required target column has no source key",
                "required_source": target.get("required_source"),
            })
    listed, cap_meta = semantics.cap_list(
        mappings, PLAN_LIST_CAP, result, "MAPPINGS_TRUNCATED",
        "mapping entries")
    complete = not unresolved and not conflicts
    plan = {
        "plan_version": PLAN_VERSION,
        "generator": f"xlsx_mapping.py {PLAN_VERSION}",
        "plan_id": plan_identity(profile, source, target_structure, policy),
        "profile_id": profile.get("profile_id"),
        "profile_hash": profile.get("profile_hash"),
        "profile_version": profile.get("profile_version"),
        "target_structure_fingerprint": target_structure,
        "source_content_fingerprint": source["meta"]["content_fingerprint"],
        "source": {
            "kind": source["meta"]["kind"],
            "encoding": source["meta"].get("encoding"),
            "delimiter": source["meta"].get("delimiter"),
            "row_count": source["meta"]["row_count"],
            "content_fingerprint": source["meta"]["content_fingerprint"],
            "columns": [
                {"key": column.get("key"),
                 "value_kind": column.get("value_kind"),
                 "semantic_role": column.get("semantic_role"),
                 "role_score": column.get("role_score"),
                 "role_level": column.get("role_level"),
                 "unit_guess": column.get("unit_guess"),
                 "evidence": column.get("evidence"),
                 "null_count": column.get("null_count"),
                 "distinct_count": column.get("distinct_count"),
                 "cell_count": column.get("cell_count"),
                 "sample_values": column.get("sample_values")}
                for column in source_cols],
            "raw_columns": source["meta"]["columns"],
        },
        "policy": policy,
        "mappings": listed,
        "mappings_cap": cap_meta,
        "unresolved": unresolved,
        "conflicts": conflicts,
        "notes": sorted(set(notes)),
        "complete": complete,
        "summary": {
            "mappings": len(mappings),
            "row_block": sum(1 for m in mappings if m["mode"] == "row_block"),
            "scalar": sum(1 for m in mappings if m["mode"] == "scalar"),
            "unresolved": len(unresolved),
            "conflicts": len(conflicts),
            "requires_review": sum(1 for m in mappings
                                   if m["requires_review"]),
        },
    }
    if not complete:
        result.warn("PLAN_INCOMPLETE",
                    f"mapping plan is incomplete: {len(unresolved)} "
                    f"unresolved, {len(conflicts)} conflict(s) "
                    "(never silent).",
                    unresolved=len(unresolved), conflicts=len(conflicts))
    return plan

# ---------------------------------------------------------------------------
# B6 -- dry-run fill plan (simulation only; writes nothing)
# ---------------------------------------------------------------------------

def _bounds_of_ref(ref):
    match = semantics.RANGE_RE.match(str(ref or "").strip())
    if not match:
        return None
    start_col = semantics._col_letter_to_index(match.group(1))
    start_row = int(match.group(2))
    end_col = semantics._col_letter_to_index(match.group(3) or match.group(1))
    end_row = int(match.group(4) or match.group(2))
    return (min(start_col, end_col), min(start_row, end_row),
            max(start_col, end_col), max(start_row, end_row))


def _sheet_index(profile, sheet_name):
    for sheet in profile.get("sheets") or []:
        if sheet.get("name") == sheet_name:
            return sheet
    return None


def _slot_rows_for_column(sheet, column_index):
    """Evidenced blank slot rows for a column (None when unknown)."""
    rows = []
    for slot in sheet.get("slots") or []:
        bounds = _bounds_of_ref(slot.get("range"))
        if not bounds:
            continue
        if bounds[0] <= column_index <= bounds[2]:
            rows.extend(range(bounds[1], bounds[3] + 1))
    return sorted(set(rows)) if rows else None


def _looks_like_reference(formula):
    text = str(formula or "").strip().strip('"')
    if not text or text.startswith("="):
        return False
    return bool(semantics.RANGE_RE.match(text)) or "!" in text


def build_dry_run(plan, profile, source, source_cols, dictionary, result,
                  target_structure=None):
    """Answer 'what would a fill touch?' without touching anything."""
    source_by_key = {c["key"]: c for c in source_cols}
    index_by_key = {c["key"]: c["source_index"] for c in source_cols}
    rows = source["rows"]
    policy = plan.get("policy") or {}
    max_rows = policy.get("max_rows")
    would_write, blocked, unsupported = [], [], []
    preserved = []
    stats = {"cells": 0, "rows_touched": 0, "sheets": set(), "blocked": 0,
             "preserved": 0}
    expansion_records = []          # (planned_rows, available_rows | None)
    expansion_unknown = set()
    blocks = {}                     # (sheet, block_id) -> expansion block
    for mapping in plan.get("mappings") or []:
        target = mapping.get("target") or {}
        sheet_name = target.get("sheet")
        letter = target.get("letter")
        sheet = _sheet_index(profile, sheet_name)
        if sheet is None or not letter:
            blocked.append({"target": f"{sheet_name}!{letter}",
                            "source_key": mapping.get("source_key"),
                            "reason": "target column is not in the profile"})
            stats["blocked"] += 1
            continue
        column = next((c for c in sheet.get("columns") or []
                       if c.get("letter") == letter), None)
        column_index = column.get("index") if column else None
        role = mapping.get("target_role") or (column or {}).get("semantic_role")
        target_kind = (mapping.get("target_value_kind") or
                       (column or {}).get("value_kind"))
        preserve = role == "formula_derived" or target_kind == "formula"
        source_key = mapping.get("source_key")
        source_index = index_by_key.get(source_key)
        if source_index is None:
            blocked.append({"target": f"{sheet_name}!{letter}",
                            "source_key": source_key,
                            "reason": "source key not present in the source "
                                      "table"})
            stats["blocked"] += 1
            continue
        anchor = target.get("anchor")
        bounds = _bounds_of_ref(anchor)
        mode = mapping.get("mode")
        planned_rows = len(rows) if mode == "row_block" else 1
        if max_rows is not None:
            planned_rows = min(planned_rows, max_rows)
        if bounds is None:
            blocked.append({"target": f"{sheet_name}!{letter}",
                            "source_key": source_key,
                            "reason": "no evidenced anchor/slot row for this "
                                      "target (nothing is guessed)"})
            stats["blocked"] += 1
            continue
        anchor_row, column_index_of_anchor = bounds[1], bounds[0]
        column_letter = semantics._col_letter_to_index(letter)
        validations = []
        for entry in sheet.get("validations") or []:
            entry_bounds = _bounds_of_ref(entry.get("range"))
            if entry_bounds:
                validations.append((entry, entry_bounds))
        merges = [_bounds_of_ref(ref) for ref in
                  (sheet.get("merged_cells") or [])]
        required = bool((column or {}).get("required"))
        for offset in range(planned_rows):
            row = anchor_row + offset
            cell = f"{letter}{row}"
            value = rows[offset][source_index] if offset < len(rows) and \
                source_index < len(rows[offset]) else None
            empty = value is None or str(value).strip() == ""
            checks = {}
            matching_validation = None
            for entry, entry_bounds in validations:
                if (entry_bounds[0] <= column_letter <= entry_bounds[2] and
                        entry_bounds[1] <= row <= entry_bounds[3]):
                    matching_validation = entry
                    break
            checks["validation"] = ("none" if matching_validation is None
                                    else f"{matching_validation.get('type')}")
            checks["merged"] = ("inside_merged_range"
                                if any(bounds2 and bounds2[0] <= column_letter
                                       <= bounds2[2] and
                                       bounds2[1] <= row <= bounds2[3]
                                       for bounds2 in merges)
                                else "not_merged")
            checks["format_compatible"] = mapping.get("compatibility") in (
                "exact", "numeric_family", "date_family", "blank_tolerant")
            checks["role_compatible"] = mapping.get("target_role") == \
                mapping.get("source_role") or mapping.get("compatibility") == \
                "exact"
            entry = {
                "sheet": sheet_name,
                "cell": cell,
                "source": f"{source_key}[{offset}]",
                "preview": (None if empty
                            else str(value)[:semantics.SNIPPET_MAX]),
                "type": mapping.get("source_value_kind"),
                "checks": checks,
                "write_policy": ("preserve_formula" if preserve
                                 else "write_value"),
            }
            if preserve:
                entry["note"] = ("formula-derived target: preserved, never "
                                 "overwritten")
                preserved.append(entry)
                stats["preserved"] += 1
                continue
            if empty and required:
                blocked.append(dict(entry, reason="source value is empty for a "
                                                  "required target; nothing "
                                                  "is invented"))
                stats["blocked"] += 1
                continue
            failing = []
            if matching_validation is not None and empty:
                failing.append("validation target would stay empty")
            if checks["merged"] == "inside_merged_range":
                failing.append("cell lies inside a merged range")
            if not checks["format_compatible"]:
                failing.append("value kinds are incompatible")
            if failing:
                blocked.append(dict(entry, reason="; ".join(failing)))
                stats["blocked"] += 1
                continue
            if matching_validation is not None and \
                    _looks_like_reference(
                        (matching_validation or {}).get("formula1")):
                unsupported.append({
                    "cell": cell,
                    "message": "target has a reference-style validation "
                               "(lookup/list): validating against another "
                               "range is NOT executed in this phase.",
                    "context": {"validation": matching_validation.get("type")},
                })
                result.unsupported_item(
                    "reference-style validation (lookup/list) is reported, "
                    "never executed in this phase",
                    sheet=sheet_name, cell=cell,
                    validation=matching_validation.get("type"))
            would_write.append(entry)
            stats["cells"] += 1
            stats["sheets"].add(sheet_name)
            if offset + 1 > stats["rows_touched"]:
                stats["rows_touched"] = offset + 1
        if mode == "row_block":
            available = _slot_rows_for_column(sheet, column_index)
            if available is None:
                expansion_unknown.add(sheet_name)
            expansion_records.append((planned_rows, available))
            # FAZ 3B evidence: which block expands, from which anchor, with
            # which columns and how much slot capacity per column. The 3B
            # boundary itself is derived live at execution time (D1); these
            # fields are the deterministic plan-side locator evidence.
            record = blocks.setdefault(
                (sheet_name, target.get("block_id")),
                {"sheet": sheet_name, "block_id": target.get("block_id"),
                 "anchor_row": anchor_row, "planned_rows": 0,
                 "columns": [], "slot_range": target.get("slot_range"),
                 "available": []})
            record["planned_rows"] = max(record["planned_rows"], planned_rows)
            if letter not in record["columns"]:
                record["columns"].append(letter)
            if available is not None:
                record["available"].append(len(available))
    for sheet_name in sorted(expansion_unknown):
        result.warn("ROW_SLOTS_UNKNOWN",
                    f"{sheet_name}: the Document Map carries no evidenced "
                    "blank input slots for this target, so whether row "
                    "expansion is needed cannot be determined (it is NOT "
                    "assumed).",
                    sheet=sheet_name)
    for entry in blocked:
        result.diagnose(f"blocked: {entry.get('cell') or entry.get('target')} "
                        f"({entry.get('reason')})")
    listed, cap_meta = semantics.cap_list(
        would_write, DRY_RUN_LIST_CAP, result, "DRY_RUN_TRUNCATED",
        "would_write entries")
    known = [(planned, available) for planned, available in expansion_records
             if available is not None]
    if not expansion_records or not known:
        needed = None
    else:
        needed = any(planned > len(available) for planned, available in known)
    expansion_report = {
        "needed": needed,
        "requested_rows": max((planned for planned, _ in expansion_records),
                              default=0),
        "available_rows": (None if not known
                           else sum(len(available)
                                    for _, available in known)),
        "unknown_sheets": sorted(expansion_unknown),
        "actual_rows_added": 0,
        "status": "planned_only",
    }
    expansion_blocks = []
    for key in sorted(blocks, key=lambda k: (str(k[0]), str(k[1]))):
        record = blocks[key]
        avail = record["available"]
        entry = {
            "sheet": record["sheet"],
            "block_id": record["block_id"],
            "anchor_row": record["anchor_row"],
            "planned_rows": record["planned_rows"],
            "columns": sorted(set(record["columns"])),
            "slot_range": record["slot_range"],
            "available_rows": (max(avail) if avail else None),
            "needed": (None if not avail
                       else record["planned_rows"] > max(avail)),
        }
        if len(set(avail)) > 1:
            entry["available_mixed"] = True
        expansion_blocks.append(entry)
    execution = {"expansion": {
        "mode": ("execute" if policy.get("allow_expansion") is True
                 else "not_authorized")},
        "lookup": {
            "mode": ("execute" if policy.get("allow_lookup") is True
                     else "not_authorized"),
            "table": {"kind": "source_json_tables",
                      "name_or_range": "main",
                      "key_column": ""},
            "key": {"from": "source", "column": ""},
            "result_column": "",
            "target_column": "",
            "on_missing": "LOOKUP_NOT_FOUND",
            "on_duplicate": "LOOKUP_AMBIGUOUS",
            "result_kind": "literal",
            "allow_formula_overwrite": False}}
    fill_plan = {
        "dry_run_version": DRY_RUN_VERSION,
        "generator": f"xlsx_mapping.py {DRY_RUN_VERSION}",
        "plan_id": plan.get("plan_id"),
        "profile_id": profile.get("profile_id"),
        "target_structure_fingerprint": target_structure,
        "plan_target_structure_fingerprint": plan.get(
            "target_structure_fingerprint"),
        "target_drift": (None if target_structure is None else
                         target_structure != plan.get(
                             "target_structure_fingerprint")),
        "would_write": listed,
        "would_write_cap": cap_meta,
        "blocked": blocked,
        "preserved": preserved,
        "unresolved": plan.get("unresolved") or [],
        "unsupported": plan.get("conflicts") or [],
        "row_expansion": expansion_report,
        "expansion": {"version": "3b.1", "blocks": expansion_blocks},
        "execution": execution,
        "stats": {
            "cells": stats["cells"],
            "rows_touched": stats["rows_touched"],
            "sheets": sorted(stats["sheets"]),
            "blocked": stats["blocked"],
            "preserved_formula_cells": stats["preserved"],
        },
        "policy": policy,
        "dictionary_version": dictionary["version"],
    }
    if fill_plan["row_expansion"]["needed"]:
        result.warn("ROW_EXPANSION_PLANNED",
                    "the source has more records than the evidenced blank "
                    "slots; row expansion is PLANNED ONLY and not applied in "
                    "this phase (never silent).",
                    requested=fill_plan["row_expansion"]["requested_rows"])
    if fill_plan["target_drift"]:
        result.warn("PLAN_TARGET_DRIFT",
                    "the target workbook's structure fingerprint differs from "
                    "the plan's; the dry-run follows the live target "
                    "(never silent).",
                    plan_fingerprint=plan.get("target_structure_fingerprint"),
                    target_fingerprint=target_structure)
    return fill_plan


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def write_out(path, payload):
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(common.canonical_json(payload) + "\n", encoding="utf-8")
    return target


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


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="FAZ 2B: build a source->target mapping plan and a "
                    "dry-run fill plan from a template profile. Writes "
                    "nothing: the target workbook is only read through the "
                    "FAZ 2A Document Map.")
    parser.add_argument("--profile", required=True, metavar="PATH",
                        help="template profile JSON from xlsx_semantics.py")
    parser.add_argument("--source", metavar="PATH",
                        help="source table (CSV or JSON)")
    parser.add_argument("--plan", metavar="PATH",
                        help="mapping plan JSON (required for --emit dry-run)")
    parser.add_argument("--target", metavar="PATH",
                        help="target workbook (.xlsx) -- read-only, via 2A")
    parser.add_argument("--doc-map", metavar="PATH",
                        help="target Document Map JSON instead of a workbook")
    parser.add_argument("--emit", choices=("plan", "dry-run"), default="plan")
    parser.add_argument("--policy-file", metavar="PATH",
                        help="JSON policy overrides (report_only/block, "
                             "max_rows)")
    parser.add_argument("--out", metavar="PATH")
    parser.add_argument("--pretty", action="store_true")
    parser.add_argument("--samples", type=int,
                        default=understand.SAMPLE_DEFAULT, metavar="N")
    parser.add_argument("--dictionary", metavar="PATH")
    args = parser.parse_args(argv)

    result = common.Result()
    dictionary = semantics.load_dictionary(args.dictionary)
    profile = semantics.load_profile(args.profile, result)
    policy = read_policy(args.policy_file, result)

    target_structure = profile.get("structure_fingerprint")
    if args.doc_map or args.target:
        if args.doc_map:
            doc_map = semantics.load_doc_map(args.doc_map, result)
        else:
            doc_map = semantics.build_map_from_workbook(
                args.target, result, args.samples)
        target_structure = (doc_map.get("workbook") or {}).get(
            "structure_fingerprint")
        if args.emit == "plan" and target_structure != profile.get(
                "structure_fingerprint"):
            result.warn("TARGET_PROFILE_DRIFT",
                        "the given target's structure fingerprint differs "
                        "from the profile's; the plan records the live "
                        "target (never silent).",
                        profile_fingerprint=profile.get(
                            "structure_fingerprint"),
                        target_fingerprint=target_structure)

    if args.emit == "dry-run":
        if not args.plan:
            raise common.XlsxError(
                "SPEC_INVALID", "--emit dry-run needs --plan PATH.",
                recovery="Run --emit plan first, then dry-run that plan.",
                context={})
        if not args.source:
            raise common.XlsxError(
                "SPEC_INVALID", "--emit dry-run needs --source PATH.",
                recovery="The dry-run previews real source values; without "
                         "the source it would have to invent them.",
                context={})
        plan = json.loads(Path(args.plan).read_text(encoding="utf-8"))
        source = load_source(args.source, result)
        source_cols = source_semantics(source, dictionary, result)
        fill_plan = build_dry_run(plan, profile, source, source_cols,
                                  dictionary, result,
                                  target_structure=target_structure)
        payload = {"ok": True, "mode": "mapping", "command": "dry-run",
                   "dictionary_version": dictionary["version"],
                   "fill_plan": fill_plan}
        if args.out:
            write_out(args.out, fill_plan)
            payload["out"] = str(Path(args.out))
        return emit(payload, result, args.pretty)

    if not args.source:
        raise common.XlsxError(
            "SPEC_INVALID", "--emit plan needs --source PATH.",
            recovery="Pass a CSV/JSON source table to map from.",
            context={})
    source = load_source(args.source, result)
    source_cols = source_semantics(source, dictionary, result)
    plan = build_plan(profile, source, source_cols, policy, result,
                      target_structure=target_structure)
    payload = {"ok": True, "mode": "mapping", "command": "plan",
               "dictionary_version": dictionary["version"], "plan": plan}
    if args.out:
        write_out(args.out, plan)
        payload["out"] = str(Path(args.out))
    return emit(payload, result, args.pretty)


if __name__ == "__main__":
    sys.exit(common.guard(main)())
