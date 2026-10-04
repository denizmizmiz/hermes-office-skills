#!/usr/bin/env python3
# MIT License. Part of the Hermes xlsx skill.
"""FAZ 3A -- template execution core: apply an APPROVED fill plan.

This is the first and only module in the skill allowed to modify a
workbook. Everything it writes comes from an approved FAZ 2B fill plan
(``xlsx_mapping.py --emit dry-run``); nothing is derived, guessed or
invented at execution time.

The write chain (I3A-7), all fail-closed:

    validate fill plan
      -> approval check (token = first 12 chars of plan_id)
      -> preflight (fingerprints, cells, merges, formulas, types,
         validations, duplicates, idempotency)
      -> snapshot (values / formulas / styles / structures)
      -> apply planned values in memory
      -> save to a STAGING file (Faz 1 safe-write: temp -> reopen -> atomic)
      -> deep QA on the staging file (formula inventory, unexpected
         modification detection, structural fidelity)
      -> only then atomic commit to the final output
      -> manifest record

QA failure means NO commit: the original is untouched and the staged
file is kept for diagnostics. In-place execution additionally takes a
backup before the commit and rolls back if the post-commit verification
fails.

Scope guardrails: FAZ 3A cell writes, plus the FAZ 3B expansion chain
ONLY when the plan is authorized (policy allow_expansion ->
execution.expansion.mode == "execute"): row expansion, copy-down
formula / style / validation propagation -- all through xlsx_expand.py,
the single implemented engine (never a second implementation). A plan
that needs expansion without authorization still stops with
ROW_EXPANSION_REQUIRED. Never here: formula family discovery, R1C1
engines, arbitrary formula synthesis, semantic remapping, automatic
unit conversion (FAZ 4).

Commands:
    xlsx_execute.py --fill-plan F.json --profile P.json --source S.csv \
                    --target T.xlsx [--emit preflight|execute]
    xlsx_execute.py ... --approve-token <first 12 chars of plan_id>
    xlsx_execute.py ... --in-place --approve-token ...   # in-place + backup
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_common as common          # noqa: E402
import xlsx_expand as expand          # noqa: E402  (FAZ 3B engine)
import xlsx_mapping as mapping        # noqa: E402
import xlsx_semantics as semantics    # noqa: E402
import xlsx_understand as understand  # noqa: E402

EXECUTE_VERSION = "3a.1"
BLOCKER_LIST_CAP = 500
CELL_LIST_CAP = 500
STYLE_SCAN_CAP = 50_000        # per workbook, for full style snapshot
DEFAULT_OUT_SUFFIX = "_filled"

CELL_RE = re.compile(r"^([A-Za-z]{1,3})([1-9][0-9]{0,6})$")
SOURCE_REF_RE = re.compile(r"^(?P<key>.+)\[(?P<offset>[0-9]+)\]$")
MAX_COLUMN = 16384
MAX_ROW = 1048576


# ---------------------------------------------------------------------------
# small shared helpers
# ---------------------------------------------------------------------------

def _col_index(letter: str) -> int:
    index = 0
    for ch in letter.upper():
        index = index * 26 + (ord(ch) - 64)
    return index


def parse_cell(ref) -> tuple[int, int] | None:
    """'B12' -> (column_index, row); None when the ref is not a cell."""
    match = CELL_RE.match(str(ref or "").strip())
    if not match:
        return None
    column, row = _col_index(match.group(1)), int(match.group(2))
    if column > MAX_COLUMN or row > MAX_ROW:
        return None
    return column, row


def cell_key(sheet: str, cell: str) -> str:
    return f"{sheet}!{cell}"


def _raw(wb):
    """Unwrap a LoadedWorkbook to the openpyxl workbook it guards."""
    inner = getattr(wb, "_wb", None)
    return inner if inner is not None else wb


def structure_fingerprint_of(wb) -> str:
    """Structure fingerprint computed EXACTLY like FAZ 2A (parity verified).

    FAZ 2A builds ``sheet|norm(dims)|tables`` parts through
    ``understand._norm_dims`` and hashes them with
    ``common.fingerprint_structure_parts``. Execution compares against the
    fingerprint stored in the fill plan, so the same recipe is reused here
    (a test pins the parity with a real profile fingerprint).
    """
    wb = _raw(wb)
    parts = []
    for ws in wb.worksheets:
        tables = ",".join(
            f"{name}={ref}"
            for name, ref in sorted(getattr(ws, "tables", {}).items()))
        parts.append(f"{ws.title}|{understand._norm_dims(ws.dimensions)}"
                     f"|{tables}")
    return common.fingerprint_structure_parts(parts)


def merged_bounds(ws):
    """Sorted list of (min_col, min_row, max_col, max_row) merged bounds."""
    from openpyxl.utils import range_boundaries  # noqa: PLC0415
    bounds = []
    for ref in ws.merged_cells.ranges:
        try:
            bounds.append(range_boundaries(str(ref)))
        except ValueError:
            continue
    return sorted(bounds)


def merged_status(ws, cell: str) -> str:
    """'not_merged' | 'top_left' | 'inside' for a cell in a sheet."""
    parsed = parse_cell(cell)
    if parsed is None:
        return "not_merged"
    column, row = parsed
    for min_col, min_row, max_col, max_row in merged_bounds(ws):
        if min_col <= column <= max_col and min_row <= row <= max_row:
            if (column, row) == (min_col, min_row):
                return "top_left"
            return "inside"
    return "not_merged"


def _defined_names(wb) -> list:
    names = []
    try:
        items = getattr(wb.defined_names, "items", None)
        if callable(items):
            for name, definition in wb.defined_names.items():
                text = getattr(definition, "attr_text", None)
                if text is None:
                    text = getattr(definition, "value", None)
                names.append(f"{name}={text}")
        else:  # pragma: no cover - defensive: other openpyxl shapes
            for name in wb.defined_names:
                names.append(str(name))
    except Exception:  # noqa: BLE001
        return []
    return sorted(names)


def _validations_of(ws) -> list:
    entries = []
    for dv in ws.data_validations.dataValidation:
        entries.append({
            "type": str(dv.type),
            "operator": str(dv.operator) if dv.operator else None,
            "formula1": None if dv.formula1 is None else str(dv.formula1),
            "formula2": None if dv.formula2 is None else str(dv.formula2),
            "ranges": sorted(str(rng) for rng in dv.sqref.ranges),
        })
    return sorted(entries, key=lambda item: json.dumps(item, sort_keys=True))


def _conditional_formatting_of(ws) -> list:
    entries = []
    try:
        for cf in ws.conditional_formatting:
            ranges = sorted(str(rng) for rng in cf.sqref.ranges)
            entries.append({"ranges": ranges,
                            "rules": len(list(cf.rules))})
    except Exception:  # noqa: BLE001
        return []
    return sorted(entries, key=lambda item: json.dumps(item, sort_keys=True))


def _hidden_of(ws) -> dict:
    rows = sorted(str(index) for index, dim in ws.row_dimensions.items()
                  if getattr(dim, "hidden", False))
    columns = sorted(str(index) for index, dim in ws.column_dimensions.items()
                     if getattr(dim, "hidden", False))
    return {"rows": rows, "columns": columns}


def style_signature(cell) -> str:
    """Everything the execution must NOT change about a cell's style."""
    font = cell.font
    fill = cell.fill
    border = cell.border
    alignment = cell.alignment
    protection = cell.protection
    font_color = None
    try:
        font_color = font.color.rgb if font.color is not None else None
    except Exception:  # noqa: BLE001
        font_color = None
    fill_color = None
    try:
        fill_color = (fill.fgColor.rgb if fill.fgColor is not None
                      else None)
    except Exception:  # noqa: BLE001
        fill_color = None
    parts = (
        font.bold, font.italic, font.underline, font.size, font.name,
        font_color,
        fill.patternType, fill_color,
        border.left.style, border.right.style, border.top.style,
        border.bottom.style,
        alignment.horizontal, alignment.vertical, alignment.wrap_text,
        cell.number_format, protection.locked, protection.hidden,
    )
    return "|".join("" if part is None else str(part) for part in parts)


def _cell_map(ws) -> dict:
    """{coordinate: value} for every existing cell that holds a value.

    openpyxl keeps all materialised cells in ``ws._cells``; iterating that
    dict is O(existing cells) and safe for very wide sheets, unlike
    ``iter_rows`` over a huge used range.
    """
    cells = getattr(ws, "_cells", None)
    if not isinstance(cells, dict):
        return {}
    out = {}
    for cell in cells.values():
        if cell.value is not None:
            out[cell.coordinate] = cell.value
    return out


def snapshot(wb, *, style_cells=(), result=None) -> dict:
    """Before/after state used for unexpected-modification detection."""
    wb = _raw(wb)
    snap = {
        "sheets": [ws.title for ws in wb.worksheets],
        "values": {},
        "formulas": {},
        "merges": {},
        "tables": {},
        "defined_names": _defined_names(wb),
        "validations": {},
        "conditional_formatting": {},
        "hidden": {},
        "styles": {},
        "style_scope": {},
    }
    style_cells = set(style_cells)
    total_cells = 0
    for ws in wb.worksheets:
        values = _cell_map(ws)
        total_cells += len(values)
        snap["values"][ws.title] = values
        snap["merges"][ws.title] = sorted(
            str(ref) for ref in ws.merged_cells.ranges)
        snap["tables"][ws.title] = {
            name: ref for name, ref in sorted(
                getattr(ws, "tables", {}).items())}
        snap["validations"][ws.title] = _validations_of(ws)
        snap["conditional_formatting"][ws.title] = \
            _conditional_formatting_of(ws)
        snap["hidden"][ws.title] = _hidden_of(ws)
        for cell in getattr(ws, "_cells", {}).values():
            if cell.data_type == "f" or (
                    isinstance(cell.value, str)
                    and cell.value.startswith("=")):
                snap["formulas"][cell_key(ws.title, cell.coordinate)] = \
                    str(cell.value)
    # style scope: all cells when the workbook is small, else only the
    # planned cells (honest limitation, reported through style_scope).
    full_style_scan = total_cells <= STYLE_SCAN_CAP
    for ws in wb.worksheets:
        keys = []
        if full_style_scan:
            for cell in getattr(ws, "_cells", {}).values():
                # phantom cells (created by read-only scans; no value, no
                # style) have no XML identity -- they cannot be a style
                # change; skipping them keeps before/after symmetric.
                if cell.value is None and not cell.has_style:
                    continue
                keys.append(cell.coordinate)
        else:
            keys = [key.split("!", 1)[1] for key in style_cells
                    if key.startswith(ws.title + "!")]
        for coordinate in keys:
            snap["styles"][cell_key(ws.title, coordinate)] = \
                style_signature(ws[coordinate])
        snap["style_scope"][ws.title] = (
            "all_cells" if full_style_scan else "planned_cells_only")
    snap["formula_count"] = len(snap["formulas"])
    snap["cell_count"] = total_cells
    if not full_style_scan and result is not None:
        result.diagnose(
            "style snapshot limited to planned cells (workbook is large)",
            style_scan_cap=STYLE_SCAN_CAP, cells=total_cells)
    return snap


def compare_snapshots(before: dict, after: dict, planned: set) -> dict:
    """Exact diff between two snapshots limited to the planned cells.

    Returns the four lists the QA stage needs:
      changed[]     -- planned cells whose value actually changed
      noop[]        -- planned cells whose value was already equal
      unexpected[]  -- ANY other modification (values, formulas, merges,
                       tables, defined names, validations, conditional
                       formatting, hidden state, styles, sheet list)
    """
    changed, noop = [], []
    for key in sorted(planned):
        sheet, cell = key.split("!", 1)
        before_value = before["values"].get(sheet, {}).get(cell)
        after_value = after["values"].get(sheet, {}).get(cell)
        if before_value == after_value:
            noop.append(key)
        else:
            changed.append(key)
    unexpected = []
    # 1. values outside the planned set
    for sheet in sorted(set(before["values"]) | set(after["values"])):
        before_values = before["values"].get(sheet, {})
        after_values = after["values"].get(sheet, {})
        for coordinate in sorted(set(before_values) | set(after_values)):
            key = cell_key(sheet, coordinate)
            if key in planned:
                continue
            if before_values.get(coordinate) != after_values.get(coordinate):
                unexpected.append({
                    "kind": "value", "cell": key,
                    "before": before_values.get(coordinate),
                    "after": after_values.get(coordinate)})
    # 2. formulas (count and text must not change at all)
    for key in sorted(set(before["formulas"]) | set(after["formulas"])):
        if before["formulas"].get(key) != after["formulas"].get(key):
            unexpected.append({
                "kind": "formula", "cell": key,
                "before": before["formulas"].get(key),
                "after": after["formulas"].get(key)})
    # 3. structural aspects (plan-outside changes must be zero)
    for aspect in ("sheets", "merges", "tables", "defined_names",
                   "validations", "conditional_formatting", "hidden"):
        if before.get(aspect) != after.get(aspect):
            unexpected.append({
                "kind": f"structure:{aspect}",
                "before": before.get(aspect),
                "after": after.get(aspect)})
    # 4. styles of the scanned cells
    if before.get("styles") != after.get("styles"):
        for key in sorted(set(before.get("styles", {}))
                          | set(after.get("styles", {}))):
            if before.get("styles", {}).get(key) != \
                    after.get("styles", {}).get(key):
                unexpected.append({
                    "kind": "style", "cell": key,
                    "before": before.get("styles", {}).get(key),
                    "after": after.get("styles", {}).get(key)})
    return {"changed": changed, "noop": noop, "unexpected": unexpected,
            "formula_count_before": before.get("formula_count"),
            "formula_count_after": after.get("formula_count")}


# ---------------------------------------------------------------------------
# fill plan loading + schema (I3A-1: only a valid plan reaches execution)
# ---------------------------------------------------------------------------

FILL_PLAN_ENTRY_SCHEMA = common.Schema(
    required=("sheet", "cell", "source", "write_policy"),
    optional=("preview", "type", "checks", "note"),
    types={"sheet": "str", "cell": "str", "source": "str",
           "write_policy": "str", "note": "str"},
)

ROW_EXPANSION_SCHEMA = common.Schema(
    required=("status",),
    optional=("needed", "requested_rows", "available_rows", "unknown_sheets",
              "actual_rows_added"),
    types={"status": "str", "actual_rows_added": "int"},
)

# FAZ 3B: the dry-run's deterministic structural evidence (D1) and the
# explicit execution opt-in namespace. Both additive: plans from the
# 2b.1 era carry neither key and behave exactly as before.
EXPANSION_BLOCK_SCHEMA = common.Schema(
    required=("sheet", "anchor_row", "planned_rows", "columns"),
    optional=("block_id", "slot_range", "available_rows", "needed",
              "available_mixed"),
    types={"sheet": "str", "anchor_row": "int", "planned_rows": "int"},
)

EXPANSION_SCHEMA = common.Schema(
    required=("version",),
    optional=("blocks",),
    many=("blocks",),
    children={"blocks": EXPANSION_BLOCK_SCHEMA},
)

EXECUTION_EXPANSION_SCHEMA = common.Schema(
    optional=("mode",),
    types={"mode": "str"})

EXECUTION_SCHEMA = common.Schema(
    optional=("expansion", "lookup"),
    children={"expansion": EXECUTION_EXPANSION_SCHEMA,
              "lookup": common.Schema(
                  optional=("mode", "table", "key", "result_column",
                            "target_column", "on_missing", "on_duplicate",
                            "result_kind", "allow_formula_overwrite"),
                  types={"mode": "str", "result_kind": "str",
                         "allow_formula_overwrite": "bool"})})

FILL_PLAN_SCHEMA = common.Schema(
    required=("dry_run_version", "plan_id", "would_write", "blocked",
              "preserved", "row_expansion", "stats", "policy"),
    optional=("generator", "profile_id", "target_structure_fingerprint",
              "plan_target_structure_fingerprint", "target_drift",
              "would_write_cap", "unresolved", "unsupported",
              "dictionary_version", "requires_approval", "expansion",
              "execution"),
    types={"dry_run_version": "str", "plan_id": "str", "policy": "dict",
           "stats": "dict", "requires_approval": "bool"},
    many=("would_write", "blocked", "preserved", "unresolved", "unsupported"),
    children={"would_write": FILL_PLAN_ENTRY_SCHEMA,
              "row_expansion": ROW_EXPANSION_SCHEMA,
              "expansion": EXPANSION_SCHEMA,
              "execution": EXECUTION_SCHEMA},
)


def load_fill_plan(path, result) -> dict:
    """Load and schema-validate a FAZ 2B dry-run fill plan (fail closed)."""
    target = Path(path)
    if not target.exists():
        raise common.XlsxError(
            "FILE_NOT_FOUND", f"fill plan not found: {target}",
            recovery="Run xlsx_mapping.py --emit dry-run first.",
            context={"fill_plan": str(target)})
    try:
        plan = json.loads(target.read_text(encoding="utf-8"))
    except (ValueError, UnicodeDecodeError) as exc:
        raise common.XlsxError(
            "VALIDATION_FAILED", f"fill plan is not valid JSON: {exc}",
            context={"fill_plan": str(target)}) from exc
    validation = common.validate_spec(plan, FILL_PLAN_SCHEMA, strict=True,
                                      path="fill_plan")
    if not validation["ok"]:
        first = validation["errors"][0]
        raise common.XlsxError(
            "PLAN_INVALID",
            f"fill plan schema error at '{first['path']}': expected "
            f"{first['expected']}, got {first['actual']}.",
            recovery=("Regenerate the fill plan with the current "
                      "xlsx_mapping.py; hand-edited plans are refused."),
            context={"fill_plan": str(target),
                     "errors": validation["errors"][:20]})
    plan_id = str(plan.get("plan_id") or "")
    if not re.fullmatch(r"[0-9a-f]{32,64}", plan_id):
        raise common.XlsxError(
            "PLAN_INVALID", "plan_id is not a content hash.",
            recovery="Regenerate the fill plan.", context={"plan_id": plan_id})
    return plan


def required_plan_fields(plan: dict) -> list:
    """Field names the preflight absolutely needs (honest gap list)."""
    needed = ["plan_id", "would_write", "row_expansion", "policy"]
    return [name for name in needed if plan.get(name) is None]


# ---------------------------------------------------------------------------
# source resolution + validation helpers
# ---------------------------------------------------------------------------

def resolve_source_ref(source: dict, ref: str):
    """'Ad[0]' -> (key, offset, raw value) or None when unresolvable."""
    match = SOURCE_REF_RE.match(str(ref or "").strip())
    if not match:
        return None
    key, offset = match.group("key"), int(match.group("offset"))
    if key not in source["headers"]:
        return None
    index = source["headers"].index(key)
    rows = source["rows"]
    if offset >= len(rows):
        return None
    row = rows[offset]
    raw = row[index] if index < len(row) else None
    return key, offset, raw


def converted_value(raw):
    """Deterministic scalar conversion (same inference as the dry-run)."""
    kind, value = mapping._infer_scalar(raw)
    return kind, value


def _validation_bounds(dv) -> list:
    from openpyxl.utils import range_boundaries  # noqa: PLC0415
    bounds = []
    for rng in dv.sqref.ranges:
        try:
            bounds.append(range_boundaries(str(rng)))
        except ValueError:
            continue
    return bounds


def live_validation_for(ws, cell: str):
    """The DataValidation object covering a cell, or None."""
    parsed = parse_cell(cell)
    if parsed is None:
        return None
    column, row = parsed
    for dv in ws.data_validations.dataValidation:
        for min_col, min_row, max_col, max_row in _validation_bounds(dv):
            if min_col <= column <= max_col and min_row <= row <= max_row:
                return dv
    return None


def _split_list_formula(formula1: str):
    """'\"GRS,RCS\"' -> ['GRS','RCS']; None when not a literal list."""
    text = str(formula1 or "").strip()
    if len(text) >= 2 and text.startswith('"') and text.endswith('"'):
        items = [item.strip() for item in text[1:-1].split(",")]
        return [item for item in items if item]
    return None


def _number(text):
    try:
        return float(str(text).replace(",", "."))
    except (TypeError, ValueError):
        return None


def check_validation(dv, cell: str, value) -> dict:
    """Deterministic, documented subset of validation checking.

    Supported: list (literal), whole, decimal, textLength.
    Everything else is reported as unsupported (never guessed).
    """
    kind = str(dv.type)
    text = "" if value is None else str(value)
    if kind == "list":
        allowed = _split_list_formula(dv.formula1)
        if allowed is None:
            return {"status": "unsupported",
                    "message": "list validation uses a reference/formula; "
                               "not evaluated in this phase"}
        if text.strip().casefold() in {item.casefold() for item in allowed}:
            return {"status": "ok", "allowed": allowed}
        return {"status": "invalid", "allowed": allowed, "value": text}
    if kind in ("whole", "decimal"):
        number = _number(text)
        if number is None:
            return {"status": "invalid", "value": text,
                    "message": f"{kind} validation expects a number"}
        if kind == "whole" and number != int(number):
            return {"status": "invalid", "value": text,
                    "message": "whole validation expects an integer"}
        low, high = _number(dv.formula1), _number(dv.formula2)
        op = str(dv.operator or "between")
        if op == "between" and low is not None and high is not None:
            if low <= number <= high:
                return {"status": "ok"}
            return {"status": "invalid", "value": text,
                    "bounds": [low, high]}
        if op == "notBetween" and low is not None and high is not None:
            if not (low <= number <= high):
                return {"status": "ok"}
            return {"status": "invalid", "value": text, "bounds": [low, high]}
        if op in ("greaterThan", "lessThan",
                  "greaterThanOrEqual", "lessThanOrEqual") and low is not None:
            checks = {
                "greaterThan": number > low,
                "lessThan": number < low,
                "greaterThanOrEqual": number >= low,
                "lessThanOrEqual": number <= low,
            }
            return ({"status": "ok"} if checks[op]
                    else {"status": "invalid", "value": text, "bound": low})
        return {"status": "unsupported",
                "message": f"{kind} validation operator '{op}' is not "
                           "evaluated in this phase"}
    if kind == "textLength":
        low, high = _number(dv.formula1), _number(dv.formula2)
        length = len(text)
        if low is not None and high is not None and low <= length <= high:
            return {"status": "ok"}
        return {"status": "invalid", "length": length,
                "bounds": [low, high]}
    return {"status": "unsupported",
            "message": f"validation type '{kind}' is not evaluated in this "
                       "phase (reported, never guessed)"}


def find_profile_column(profile: dict, sheet_name: str, letter: str):
    for sheet in profile.get("sheets") or []:
        if sheet.get("name") != sheet_name:
            continue
        for column in sheet.get("columns") or []:
            if column.get("letter") == letter:
                return column
    return None


def planned_region(writes) -> dict:
    """Per-sheet bounding box + combined region fingerprint for idempotency."""
    from openpyxl.utils import range_boundaries  # noqa: PLC0415
    per_sheet = {}
    for sheet, cell, _value, _entry in writes:
        parsed = parse_cell(cell)
        if parsed is None:
            continue
        column, row = parsed
        box = per_sheet.setdefault(
            sheet, [column, row, column, row])
        box[0] = min(box[0], column)
        box[1] = min(box[1], row)
        box[2] = max(box[2], column)
        box[3] = max(box[3], row)
    boxes = {}
    for sheet, box in sorted(per_sheet.items()):
        from openpyxl.utils import get_column_letter  # noqa: PLC0415
        bounds = (box[0], box[1], box[2], box[3])
        ref = (f"{get_column_letter(bounds[0])}{bounds[1]}:"
               f"{get_column_letter(bounds[2])}{bounds[3]}")
        boxes[sheet] = {"bounds": bounds, "ref": ref}
    return boxes


def region_fingerprint_of(wb, boxes: dict) -> str:
    """Combined deterministic fingerprint of the planned regions."""
    wb = _raw(wb)
    parts = []
    for sheet, info in sorted(boxes.items()):
        ws = wb[sheet]
        column, row, max_column, max_row = info["bounds"]
        from openpyxl.utils import get_column_letter  # noqa: PLC0415
        ref = (f"{get_column_letter(column)}{row}:"
               f"{get_column_letter(max_column)}{max_row}")
        parts.append(f"{sheet}|{ref}|{common.fingerprint_region(ws, ref)}")
    return common.fingerprint_structure_parts(parts)


def region_is_empty(wb, writes) -> bool:
    wb = _raw(wb)
    for sheet, cell, _value, _entry in writes:
        if wb[sheet][cell].value is not None:
            return False
    return True


# ---------------------------------------------------------------------------
# preflight (I3A-1 / I3A-4: every precondition before a single write)
# ---------------------------------------------------------------------------

def _blocker(code, message, **context):
    return {"code": code, "message": message,
            **(context if context else {})}


def preflight_execution(plan, profile, source, wb, *, target_path,
                        out_path, in_place, approve_token,
                        manifest_dir=None, result=None,
                        collect_writes=None,
                        expansion_state=None) -> dict:
    """All read-only checks; deterministic report (AC-3A21).

    ``collect_writes`` -- optional dict; when given, the resolved write
    list (sheet, cell, value, entry) is stored under ``["writes"]`` so the
    executor can reuse the exact writes the preflight validated.
    """
    wb = _raw(wb)
    blockers = []
    warnings = []

    def block(code, message, **context):
        blockers.append(_blocker(code, message, **context))

    # --- plan/profile/target identity -------------------------------------
    plan_id = str(plan.get("plan_id") or "")
    if plan.get("profile_id") and profile.get("profile_id") != plan.get(
            "profile_id"):
        block("PLAN_INVALID",
              "the profile does not match the profile the fill plan was "
              "built from.",
              plan_profile=plan.get("profile_id"),
              given_profile=profile.get("profile_id"))

    target_fp = common.fingerprint_file(target_path) \
        if Path(target_path).exists() else None
    live_structure = structure_fingerprint_of(wb)
    plan_structure = plan.get("plan_target_structure_fingerprint") or \
        plan.get("target_structure_fingerprint")
    structure_match = (plan_structure is None or
                       plan_structure == live_structure)
    exp_state = expansion_state or {}
    if exp_state.get("applied"):
        # the live workbook was expanded in memory BEFORE this preflight;
        # the fingerprint was already checked against the pre-expansion
        # workbook by the expansion gate (D1).
        structure_match = True
    if not structure_match:
        block("STALE_PLAN",
              "the target workbook's structure fingerprint no longer "
              "matches the plan's; nothing was written.",
              plan_fingerprint=plan_structure,
              live_fingerprint=live_structure)

    # --- output policy ----------------------------------------------------
    if out_path is not None and not in_place and \
            Path(out_path).resolve() == Path(target_path).resolve():
        block("PLAN_INVALID",
              "the output path equals the target; pass --in-place to "
              "confirm an in-place write (backup + approval are mandatory).")

    # --- fill plan scope gates -------------------------------------------
    # FAZ 3B: lookup config validation (only when authorized)
    execution = plan.get("execution") or {}
    lookup_cfg = execution.get("lookup") or {}
    if lookup_cfg.get("mode") == "execute":
        required = [("table", lookup_cfg.get("table")),
                    ("key", lookup_cfg.get("key")),
                    ("result_column", lookup_cfg.get("result_column")),
                    ("target_column", lookup_cfg.get("target_column"))]
        missing = [name for name, val in required if not val]
        if missing:
            block("PLAN_INVALID",
                  f"execution.lookup missing required fields: {', '.join(missing)}",
                  missing=missing)
        # Validate table kind
        table_spec = lookup_cfg.get("table") or {}
        if table_spec.get("kind") not in ("source_json_tables", "target_range"):
            block("PLAN_INVALID",
                  f"execution.lookup unsupported table kind: {table_spec.get('kind')!r}",
                  context={"kind": table_spec.get("kind")})
        # Validate key spec
        key_spec = lookup_cfg.get("key") or {}
        if key_spec.get("from") != "source" or not key_spec.get("column"):
            block("PLAN_INVALID",
                  "execution.lookup key must have from=source and column",
                  context={"key_spec": key_spec})
        # Validate target column is in a block
        target_col = lookup_cfg.get("target_column")
        if target_col:
            exp_blocks = (plan.get("expansion") or {}).get("blocks") or []
            found = any(target_col.upper() in (b.get("columns") or []) for b in exp_blocks)
            if not found:
                block("PLAN_INVALID",
                      f"execution.lookup target_column {target_col!r} not in any expansion block",
                      context={"target_column": target_col,
                               "blocks": [b.get("columns") for b in exp_blocks]})

    expansion = plan.get("row_expansion") or {}
    if exp_state.get("present"):
        # FAZ 3B: the plan carries structural evidence; the gate evaluated
        # it (dry-run flag OR live capacity math) before any write.
        if exp_state.get("needed") and not exp_state.get("authorized"):
            block("ROW_EXPANSION_REQUIRED",
                  "the plan needs row expansion but is not authorized "
                  "(execution.expansion.mode != 'execute'); regenerate "
                  "with policy allow_expansion or execute without it.",
                  requested_rows=expansion.get("requested_rows"),
                  available_rows=expansion.get("available_rows"),
                  rows_to_add=exp_state.get("rows_to_add", 0))
        for item in exp_state.get("blockers") or []:
            context = {key: value for key, value in item.items()
                       if key not in ("code", "message")}
            block(item["code"], item.get("message", ""), **context)
        if (exp_state.get("apply_required") and
                exp_state.get("authorized") and
                not exp_state.get("applied")):
            block("EXECUTION_BLOCKED",
                  "expansion was required and authorized but not applied; "
                  "the run stops instead of writing into unexpanded rows.")
    elif expansion.get("needed") is True:
        block("ROW_EXPANSION_REQUIRED",
              "the source holds more records than the evidenced blank "
              "slots; row expansion is FAZ 3B and is not performed.",
              requested_rows=expansion.get("requested_rows"),
              available_rows=expansion.get("available_rows"),
              actual_rows_added=0)
    cap = plan.get("would_write_cap") or {}
    if cap.get("truncated") is True:
        block("PLAN_TRUNCATED",
              "the dry-run capped the write list (truncated); a partial "
              "operation list is never executed -- regenerate with a "
              "higher cap or a smaller batch.",
              count_total=cap.get("count_total"),
              returned_count=cap.get("returned_count"))
    policy = plan.get("policy") or {}
    blocked_entries = plan.get("blocked") or []
    hard_blocked = [
        entry for entry in blocked_entries
        if "required" not in str(entry.get("reason", ""))
        or policy.get("missing_required") == "block"]
    if hard_blocked:
        block("EXECUTION_BLOCKED",
              f"the fill plan contains {len(hard_blocked)} blocked cell(s); "
              "resolve them before executing (no partial fill).",
              blocked=hard_blocked[:20])
    conflicts = plan.get("unsupported") or []
    if conflicts and policy.get("on_conflict", "block") == "block":
        block("EXECUTION_BLOCKED",
              f"the fill plan carries {len(conflicts)} conflict(s) and the "
              "policy is on_conflict=block.",
              conflicts=conflicts[:20])
    unresolved = plan.get("unresolved") or []
    if unresolved and policy.get("missing_required") == "block":
        block("EXECUTION_BLOCKED",
              f"the fill plan carries {len(unresolved)} unresolved mapping(s) "
              "and the policy is missing_required=block.",
              unresolved=unresolved[:20])

    # --- planned writes ---------------------------------------------------
    writes, duplicates = [], set()
    seen = set()
    for entry in plan.get("would_write") or []:
        sheet, cell = entry.get("sheet"), entry.get("cell")
        key = cell_key(str(sheet), str(cell))
        if key in seen:
            duplicates.add(key)
            continue
        seen.add(key)
        if str(entry.get("write_policy")) != "write_value":
            block("PLAN_INVALID",
                  f"would_write lists an entry with write policy "
                  f"'{entry.get('write_policy')}'; only write_value may be "
                  "executed (formula cells are never written).",
                  cell=key)
            continue
        parsed = parse_cell(cell)
        if parsed is None:
            block("PLAN_INVALID",
                  f"cell reference {cell!r} is not addressable.", cell=key)
            continue
        if sheet not in wb.sheetnames:
            block("EXECUTION_BLOCKED",
                  f"sheet {sheet!r} does not exist in the target.",
                  cell=key)
            continue
        ws = wb[sheet]
        merge = merged_status(ws, cell)
        if merge == "inside":
            block("MERGED_CELL_WRITE_FORBIDDEN",
                  "the cell lies inside a merged range and is not its "
                  "top-left cell.", cell=key)
            continue
        current = ws[cell]
        if current.data_type == "f" or (
                isinstance(current.value, str)
                and current.value.startswith("=")):
            block("FORMULA_MODIFICATION_FORBIDDEN",
                  "the cell currently holds a formula; FAZ 3A never "
                  "rewrites formulas.", cell=key,
                  formula=str(current.value)[:120])
            continue
        resolved = resolve_source_ref(source, entry.get("source"))
        if resolved is None:
            block("SOURCE_CHANGED",
                  f"source reference {entry.get('source')!r} can no longer "
                  "be resolved; the source table changed since the plan.",
                  cell=key, reference=entry.get("source"))
            continue
        _key, _offset, raw = resolved
        preview = entry.get("preview")
        if preview is not None and \
                str(raw)[:semantics.SNIPPET_MAX] != str(preview):
            block("SOURCE_CHANGED",
                  "the source value no longer matches the plan's preview.",
                  cell=key, expected=preview,
                  found=str(raw)[:semantics.SNIPPET_MAX])
            continue
        kind, value = converted_value(raw)
        letter = "".join(ch for ch in str(cell) if ch.isalpha())
        column = find_profile_column(profile, str(sheet), letter)
        if value is None or (isinstance(value, str)
                             and not value.strip()):
            if column is not None and column.get("required"):
                block("EXECUTION_BLOCKED",
                      "source value is empty for a required target; nothing "
                      "is invented.", cell=key)
                continue
        # type safety (re-checked at execution time, spec section 9)
        if column is None:
            block("PLAN_INVALID",
                  f"target column {sheet}!{letter} is not in the profile.",
                  cell=key)
            continue
        target_kind = column.get("value_kind")
        compat, _compat_reason = mapping._compat_of(kind, target_kind)
        checks = entry.get("checks") or {}
        if compat == "mismatch" or checks.get("format_compatible") is False:
            block("TYPE_MISMATCH",
                  f"value kind {kind!r} is not compatible with target kind "
                  f"{target_kind!r}.", cell=key, source_kind=kind,
                  target_kind=target_kind)
            continue
        if checks.get("role_compatible") is False:
            warnings.append({"code": "ROLE_MISMATCH_REPORTED",
                             "cell": key,
                             "message": "source and target semantic roles "
                                        "differ (reported; the format check "
                                        "passed)"})
        # validation / dropdown safety
        dv = live_validation_for(ws, str(cell))
        if dv is not None:
            outcome = check_validation(dv, str(cell), value)
            if outcome["status"] == "invalid":
                block("VALIDATION_MISMATCH",
                      f"the value {str(value)[:60]!r} violates the live "
                      f"{dv.type} validation on this cell.",
                      cell=key, validation=str(dv.type),
                      allowed=outcome.get("allowed"))
                continue
            if outcome["status"] == "unsupported":
                warnings.append({"code": "UNSUPPORTED_VALIDATION",
                                 "cell": key,
                                 "message": outcome["message"]})
        writes.append((str(sheet), str(cell), value, entry))
    for key in sorted(duplicates):
        block("PLAN_INVALID",
              "duplicate write operation for the same cell; a plan must "
              "address each cell at most once.", cell=key)

    # --- approval (I3A-2: state is never invented) ------------------------
    requires_approval = plan.get("requires_approval", True) is not False
    if in_place:
        requires_approval = True
    token = str(approve_token or "")
    expected = plan_id[:12]
    approval = {"required": requires_approval, "state": "not-required",
                "expected_token_hint": "first 12 chars of plan_id",
                "token_provided": bool(token)}
    if requires_approval:
        if not token:
            approval["state"] = "missing"
            block("APPROVAL_REQUIRED",
                  "this execution requires explicit approval; no approval "
                  "token was provided.",
                  )
        elif token != expected:
            approval["state"] = "invalid"
            block("APPROVAL_INVALID",
                  "the approval token does not match this plan (expected "
                  "the first 12 characters of plan_id).")
        else:
            approval["state"] = "provided"

    # --- idempotency (I4 via the Faz 1 manifest) --------------------------
    boxes = planned_region(writes)
    region_fp = region_fingerprint_of(wb, boxes) if boxes else None
    empty_region = region_is_empty(wb, writes) if writes else True
    idem = {"status": "no_writes", "message": "no planned writes"}
    if writes:
        idem = common.idempotency_decision(
            f"sha256:{plan_id}", region_fp,
            current_region_empty=empty_region, skill_dir=manifest_dir)
        if idem["status"] == "conflict":
            block("PLAN_CONFLICT", idem["message"],
                  prior_operation=(idem.get("prior") or {}).get(
                      "operation_id"))
        elif idem["status"] == "unknown":
            # FAZ 3C / D3: absent evidence is never evidence of absence.
            # An untouched (empty) target region stays safe to write; a
            # populated region with no usable manifest proof is refused
            # unless the plan explicitly accepts the risk.
            override = policy.get("allow_unverified_region_overwrite") is True
            fresh_workspace = bool(idem.get("fresh_workspace"))
            if empty_region or override or fresh_workspace:
                warnings.append({
                    "code": idem.get("error_code") or "MANIFEST_MISSING",
                    "cell": None,
                    "message": ("idempotency could not be verified ("
                                + str((idem.get("evidence") or {}).get(
                                    "manifest_status", "unknown"))
                                + "); the target region is "
                                + ("empty" if empty_region
                                   else "populated, but this workspace never "
                                        "had a manifest" if fresh_workspace
                                   else "populated but the plan allows an "
                                        "unverified overwrite")
                                + ", so the run may continue")})
            else:
                block("IDEMPOTENCY_UNVERIFIED", idem["message"],
                      cause=idem.get("error_code"),
                      evidence=idem.get("evidence"),
                      region_fingerprint=region_fp)
        elif idem.get("duplicate_entry"):
            warnings.append({
                "code": "MANIFEST_DUPLICATE", "cell": None,
                "message": idem.get("message", ""),
                "count": idem.get("match_count")})

    blockers.sort(key=lambda item: (item["code"],
                                    str(item.get("cell", ""))))
    report = {
        "executable": not blockers,
        "blockers": blockers[:BLOCKER_LIST_CAP],
        "blocker_count": len(blockers),
        "writes_planned": len(writes),
        "planned_cells": sorted(
            cell_key(sheet, cell) for sheet, cell, _v, _e in writes
        )[:CELL_LIST_CAP],
        "sheets": sorted({sheet for sheet, _c, _v, _e in writes}),
        "structure_fingerprint": live_structure,
        "plan_target_structure_fingerprint": plan_structure,
        "structure_match": structure_match,
        "target_file_fingerprint": target_fp,
        "region_fingerprint": region_fp,
        "region_empty": empty_region,
        "idempotency": {"status": idem["status"], "message": idem["message"],
                        "prior_operation": (idem.get("prior") or {}).get(
                            "operation_id"),
                        "cause": idem.get("error_code"),
                        "fresh_workspace": bool(idem.get("fresh_workspace")),
                        "duplicate_entry": bool(idem.get("duplicate_entry")),
                        "match_count": idem.get("match_count"),
                        "evidence": idem.get("evidence")},
        "approval": approval,
        "policy": policy,
        "row_expansion": expansion,
        "expansion": (None if not exp_state.get("present") else {
            "present": True,
            "authorized": bool(exp_state.get("authorized")),
            "needed": bool(exp_state.get("needed")),
            "rows_to_add": exp_state.get("rows_to_add", 0),
            "applied": bool(exp_state.get("applied")),
            "blocks": [{"sheet": item.get("sheet"),
                        "block_id": item.get("block_id"),
                        "rows_to_add": item.get("rows_to_add"),
                        "insert_at": item.get("insert_at")}
                       for item in
                       ((exp_state.get("plan") or {}).get("blocks")
                        or [])],
        }),
        "warnings": warnings[:BLOCKER_LIST_CAP],
        "warning_count": len(warnings),
    }
    if result is not None:
        for warning in report["warnings"]:
            result.warn(warning.get("code", "EXECUTION_WARNING"),
                        warning.get("message", ""), cell=warning.get("cell"))
        for item in report["blockers"]:
            result.diagnose(
                f"blocker {item['code']}: {item.get('cell') or '-'} "
                f"({item.get('message', '')[:160]})")
    if collect_writes is not None:
        collect_writes["writes"] = writes
    return report


# ---------------------------------------------------------------------------
# execution (I3A-4/5/7/8): staged write, deep QA, then atomic commit
# ---------------------------------------------------------------------------

def resolve_output(target_path, out_arg, in_place):
    """Final output path per the default policy (new file unless in-place)."""
    target = Path(target_path)
    if in_place:
        return target
    if out_arg:
        return Path(out_arg)
    return target.with_name(f"{target.stem}{DEFAULT_OUT_SUFFIX}{target.suffix}")


def _post_commit_verify(path, expect_sheets=None) -> dict:
    """Post-commit re-verification (module-level for fault injection)."""
    return common.verify_workbook(path, expect_sheets=expect_sheets)


def _primary_error_code(report: dict) -> str:
    codes = [item["code"] for item in report.get("blockers", [])]
    for preferred in ("APPROVAL_REQUIRED", "APPROVAL_INVALID",
                      "PLAN_TRUNCATED", "PLAN_CONFLICT",
                      "STALE_PLAN", "ROW_EXPANSION_REQUIRED",
                      "RECORD_KEY_UNAVAILABLE", "DUPLICATE_RECORD",
                      "UNSUPPORTED_FORMULA_PROPAGATION",
                      "MERGE_EXPANSION_UNSUPPORTED",
                      "TABLE_EXPANSION_UNSUPPORTED",
                      "EXPANSION_MULTI_BLOCK_UNSUPPORTED",
                      "EXPANSION_BOUNDARY_UNKNOWN",
                      "MERGED_CELL_WRITE_FORBIDDEN",
                      "FORMULA_MODIFICATION_FORBIDDEN", "SOURCE_CHANGED",
                      "TYPE_MISMATCH", "VALIDATION_MISMATCH",
                      "IDEMPOTENCY_UNVERIFIED", "PLAN_INVALID"):
        if preferred in codes:
            return preferred
    return "EXECUTION_BLOCKED"


def _raise_if_blocked(report):
    """Raise the canonical blocked-execution error (never a partial write)."""
    if report.get("executable"):
        return
    code = _primary_error_code(report)
    raise common.XlsxError(
        code,
        f"execution blocked: {report.get('blocker_count')} precondition "
        f"failure(s); nothing was written. First: "
        f"{(report['blockers'][0]['message'] if report['blockers'] else '')}",
        recovery=("Resolve the blockers listed in context.blockers and "
                  "re-run. Regenerate the fill plan when the target or "
                  "source changed."),
        context={"blockers": report.get("blockers"),
                 "blocker_count": report.get("blocker_count"),
                 "write_count": 0,
                 # FAZ 3C (AC-3C16): a blocked run is a failure payload
                 # too, so it carries the same operational diagnostics.
                 "stage": "preflight",
                 "failure_stage": "preflight",
                 "operation_id": common._new_operation_id("blocked"),
                 "commit_state": "NOT_STARTED",
                 "committed": False})


COMMIT_STATES = ("NOT_STARTED", "STAGED", "VALIDATED", "COMMITTED", "FAILED")


def record_failure(plan, state, exc, *, report=None, manifest_dir=None) -> dict:
    """Write the recovery manifest entry for a failed execution (3C, D2).

    Every field of spec section 15 is present; a temp/backup that does not
    exist is reported as ``NOT_AVAILABLE`` -- never invented.
    """
    context = exc.context or {}
    temp = state.get("temp") or context.get("staging") or context.get("tmp")
    backup = state.get("backup") or context.get("backup")
    plan_id = str(plan.get("plan_id") or "")
    target_before = (report or {}).get("target_file_fingerprint")
    region_before = (report or {}).get("region_fingerprint")
    entry = {
        "operation_id": state.get("operation_id"),
        "phase": "3c",
        "execute_version": EXECUTE_VERSION,
        "plan_id": plan_id,
        "plan_hash": f"sha256:{plan_id}",
        "status": "failed",
        "committed": False,
        "commit_state": "FAILED",
        "stage": state.get("stage"),
        "failure_stage": state.get("stage"),
        "error_code": exc.code,
        "error_message": str(exc)[:300],
        "original_fingerprint": region_before,
        "target_fingerprint": target_before,
        "target_fingerprint_before": target_before,
        "region_fingerprint_before": region_before,
        "region_fingerprint": None,
        "temp": temp or "NOT_AVAILABLE",
        "backup": backup or "NOT_AVAILABLE",
        "write_count": 0,
        "in_place": bool(state.get("in_place")),
        "timestamp": common._timestamp(),
    }
    info = record_execution(entry, manifest_dir)
    return {"entry": entry, "manifest": info}


def execute_fill_plan(plan, profile, source, loaded, report, writes, *,
                      target_path, out_path, in_place, approve_token,
                      manifest_dir=None, approval_source="cli",
                      result=None) -> dict:
    """Apply an approved, preflighted fill plan (FAZ 3C hardened entry).

    Wraps the body so ANY failure is fail-closed AND fully diagnosed:
    a ``status="failed"`` recovery record lands in the manifest and the
    raised error carries ``stage``, ``operation_id``, ``commit_state``,
    ``backup`` and ``temp``. The original file is never modified on a
    failure path.
    """
    state = {
        "stage": "preflight",
        "commit_state": "NOT_STARTED",
        "operation_id": common._new_operation_id("exec"),
        "plan_id": str(plan.get("plan_id") or ""),
        "backup": None,
        "temp": None,
        "write_count": 0,
        "in_place": bool(in_place),
    }
    try:
        return _execute_fill_plan_body(
            plan, profile, source, loaded, report, writes,
            target_path=target_path, out_path=out_path, in_place=in_place,
            approve_token=approve_token, manifest_dir=manifest_dir,
            approval_source=approval_source, result=result, state=state)
    except common.XlsxError as exc:
        state["commit_state"] = "FAILED"
        recorded = record_failure(plan, state, exc, report=report,
                                  manifest_dir=manifest_dir)
        if result is not None:
            result.diagnose(
                f"execution failed at stage '{state['stage']}' "
                f"({exc.code}); a recovery record was written to the manifest",
                stage=state["stage"], operation_id=state["operation_id"],
                commit_state="FAILED", error_code=exc.code)
        raise common.XlsxError(
            exc.code, str(exc), recovery=exc.recovery,
            context={
                **(exc.context or {}),
                "stage": state["stage"],
                "failure_stage": state["stage"],
                "operation_id": state["operation_id"],
                "commit_state": "FAILED",
                "plan_id": state["plan_id"],
                "backup": state.get("backup") or "NOT_AVAILABLE",
                "temp": (state.get("temp")
                         or (exc.context or {}).get("staging")
                         or (exc.context or {}).get("tmp")
                         or "NOT_AVAILABLE"),
                "committed": False,
                "write_count": 0,
                "failure_manifest": recorded["manifest"],
            }) from exc
    except Exception as exc:  # noqa: BLE001 - never leak an unstructured error
        state["commit_state"] = "FAILED"
        wrapped = common.XlsxError(
            "UNEXPECTED_ERROR",
            f"{type(exc).__name__} at stage '{state['stage']}': {exc}")
        recorded = record_failure(plan, state, wrapped, report=report,
                                  manifest_dir=manifest_dir)
        raise common.XlsxError(
            "UNEXPECTED_ERROR", str(wrapped), recovery=wrapped.recovery,
            context={"stage": state["stage"],
                     "operation_id": state["operation_id"],
                     "commit_state": "FAILED", "plan_id": state["plan_id"],
                     "backup": state.get("backup") or "NOT_AVAILABLE",
                     "temp": state.get("temp") or "NOT_AVAILABLE",
                     "committed": False, "write_count": 0,
                     "failure_manifest": recorded["manifest"]}) from exc


def _execute_fill_plan_body(plan, profile, source, loaded, report, writes, *,
                            target_path, out_path, in_place, approve_token,
                            manifest_dir=None, approval_source="cli",
                            result=None, state=None) -> dict:
    """The execution body: expansion -> lookup -> write -> QA -> commit.

    ``loaded`` is the Pass-C LoadedWorkbook (write authority); the raw
    openpyxl workbook is used for the snapshots and the in-memory writes.
    """
    state = state if state is not None else {}
    wb = _raw(loaded)
    started = time.perf_counter()
    target_path = Path(target_path)
    plan_id = str(plan["plan_id"])

    # --- gate: preflight must be clean (never a partial write) ------------
    _raise_if_blocked(report)

    # --- FAZ 3B: expansion chain (if authorized) ---------------------------
    state["stage"] = "expansion"
    common.fault("expansion", "before the expansion chain",
                 plan_id=plan_id)
    exp_state = report.get("expansion")
    exp_plan = None
    if exp_state and exp_state.get("present") and exp_state.get("authorized"):
        # plan_expansion is read-only; it returns the deterministic structural
        # plan with blockers -> then apply_expansion mutates the in-memory
        # workbook (via the single apply_shift engine). All coordinates are
        # live (D1), nothing is guessed.
        exp_plan = expand.plan_expansion(
            plan, profile, source, loaded.wb,
            result=common.Result(mode="expansion_plan"))
        if not exp_plan.get("ok"):
            raise common.XlsxError(
                "EXECUTION_BLOCKED",
                "expansion preflight produced blockers; nothing was written.",
                context={"expansion_blockers": exp_plan.get("blockers")})
        # resolve write locators against the post-expansion layout
        loc = expand.resolve_write_locators(plan, exp_plan)
        if not loc.get("ok"):
            raise common.XlsxError(
                "EXECUTION_BLOCKED",
                "locator resolution produced blockers; nothing was written.",
                context={"locator_blockers": loc.get("blockers")})
        # apply expansion in memory (staged workbook); no save yet
        loaded_wb = expand.apply_expansion(loaded, exp_plan,
                                           result=common.Result(mode="expansion"))
        exp_state = {**exp_state, "applied": True, "plan": exp_plan,
                     "locators": loc, "rows_added": loaded_wb.get("rows_added_total", 0),
                     "blocks": loaded_wb.get("blocks", [])}
        # the write list must be regenerated against the live coordinates
        # after expansion (preflight already validated them with the new
        # locators because it was re-run with expansion_state=exp_state).
        # We just keep the planned cells; the executor writes the SAME
        # values to the resolved cells (the values don't change).
        planned_set = {cell_key(sheet, cell) for sheet, cell, _v, _e in writes}
        new_writes = []
        for w in loc.get("writes") or []:
            key = cell_key(w["sheet"], w["resolved_cell"])
            if key in planned_set:
                # find original entry to keep the value
                for sheet, cell, value, entry in writes:
                    if cell_key(sheet, cell) == cell_key(w["sheet"], w["entry_cell"]):
                        new_writes.append((w["sheet"], w["resolved_cell"], value, entry))
                        break
        if new_writes:
            writes = new_writes

    # --- FAZ 3B: lookup execution (if authorized) -----------------------
    state["stage"] = "lookup"
    common.fault("lookup", "before lookup execution", plan_id=plan_id)
    lookup_result = None
    if plan.get("execution", {}).get("lookup", {}).get("mode") == "execute":
        lookup_result = expand.execute_lookup(
            plan, profile, source, loaded,
            result=common.Result(mode="lookup"))
        blockers = lookup_result.get("blockers") or []
        if blockers:
            # FAZ 3C: an unresolved lookup (missing key / ambiguous key /
            # unmapped column) is fail-closed -- nothing is written on a
            # guess. This includes the previously silent case of
            # executed=True with blockers.
            raise common.XlsxError(
                "EXECUTION_BLOCKED",
                "lookup execution produced blockers; nothing was written.",
                context={"lookup_blockers": blockers,
                         "reason": lookup_result.get("reason"),
                         "write_count": 0})
        if lookup_result.get("results"):
            # FAZ 3C: merging lookup results into the write list is NOT
            # implemented (3B left a silent no-op here). Report it loudly
            # instead of dropping the results without a trace.
            if result is not None:
                result.warn(
                    "LOOKUP_RESULTS_NOT_MERGED",
                    f"{len(lookup_result['results'])} lookup result(s) were "
                    "resolved but are NOT written (result merging is not "
                    "implemented in this phase); no cell was modified.",
                    count=len(lookup_result["results"]))
                result.unsupported_item(
                    "lookup result merging is not implemented; resolved "
                    "values were not written",
                    resolved=len(lookup_result["results"]))

    # --- idempotency: already applied -> zero writes ----------------------
    idem = report.get("idempotency") or {}
    if idem.get("status") == "already_applied":
        return {
            "ok": True, "mode": "execute", "command": "execute",
            "execute_version": EXECUTE_VERSION,
            "status": "already_applied", "write_count": 0,
            "idempotency_status": "already_applied",
            "plan_id": plan_id, "profile_id": profile.get("profile_id"),
            "output": None,
            "prior_operation": idem.get("prior_operation"),
            "message": idem.get("message"),
            "target_file_fingerprint": report.get("target_file_fingerprint"),
            "region_fingerprint": report.get("region_fingerprint"),
            "approval": report.get("approval"),
        }

    final = resolve_output(target_path, out_path, in_place)
    state["stage"] = "write"
    state["in_place"] = bool(in_place)
    state["output"] = str(final)
    write_started = time.perf_counter()

    planned = {cell_key(sheet, cell) for sheet, cell, _v, _e in writes}
    boxes = planned_region(writes)

    # --- snapshot BEFORE the in-memory write ------------------------------
    before = snapshot(wb, style_cells=planned, result=result)

    for sheet, cell, value, _entry in writes:
        wb[sheet][cell].value = value

    # --- staged save through the Faz 1 safe-write chain -------------------
    final.parent.mkdir(parents=True, exist_ok=True)
    common.fault("temp_create", "before creating the staging file",
                 target=str(final))
    handle, staging_name = tempfile.mkstemp(
        prefix=f".{final.stem}.stage-", suffix=".xlsx",
        dir=str(final.parent))
    os.close(handle)
    state["temp"] = staging_name
    save_info = common.save_workbook_safe(
        loaded, out_path=staging_name, backup=False, in_place=False,
        approve_token=None, require_approval=False,
        expect_sheets=list(wb.sheetnames), manifest_dir=manifest_dir,
        manifest_status="staged")
    state["commit_state"] = "STAGED"
    write_done = time.perf_counter()

    # --- deep QA on the staged output (still nothing committed) -----------
    state["stage"] = "qa"
    qa = post_validate(staging_name, before, writes, boxes, result=result)
    state["commit_state"] = "VALIDATED" if qa.get("ok") else "FAILED"
    qa_done = time.perf_counter()
    if not qa["ok"]:
        raise common.XlsxError(
            "QA_FAILED",
            "post-write QA failed on the staged output; nothing was "
            "committed and the original file is untouched.",
            recovery=("Inspect the staged file listed in context.staging, "
                      "fix the plan/source, and re-run."),
            context={"qa": qa, "staging": staging_name, "committed": False,
                     "write_count": 0, "unexpected": qa["unexpected"][:20]})

    # --- commit: backup (in-place) then atomic replace --------------------
    state["stage"] = "commit"
    backup_path = None
    if in_place:
        backup_path = common.backup_workbook(target_path, tag="3a")["backup"]
        state["backup"] = backup_path
    placed = common.atomic_replace(staging_name, final)
    state["temp"] = None
    state["commit_state"] = "COMMITTED"
    final_check = _post_commit_verify(
        final, expect_sheets=list(wb.sheetnames))
    if not final_check.get("ok"):
        rolled_back = False
        if in_place and backup_path:
            common.atomic_replace(backup_path, target_path)
            rolled_back = True
        raise common.XlsxError(
            "WRITE_FAILED",
            f"committed file failed re-verification: "
            f"{final_check.get('reason')}",
            recovery=("The backup was restored."
                      if rolled_back else
                      "The original file was never touched."),
            context={"output": str(final), "rolled_back": rolled_back,
                     "backup": backup_path, "verify": final_check,
                     "write_count": 0})
    commit_done = time.perf_counter()

    # --- fingerprints, region identity, provenance ------------------------
    target_after = common.fingerprint_file(final)
    region_after = qa.get("region_fingerprint_after")
    if in_place:
        backup_info = common.fingerprint_file(backup_path) \
            if backup_path else None
        original_unchanged = (backup_info ==
                              report.get("target_file_fingerprint"))
    else:
        original_unchanged = (common.fingerprint_file(target_path) ==
                              report.get("target_file_fingerprint"))

    changes_count = len(qa.get("changed") or [])
    noop_count = len(qa.get("noop") or [])
    if noop_count and result is not None:
        result.warn("CELL_ALREADY_EQUAL",
                    f"{noop_count} planned cell(s) already held the planned "
                    "value (no change; never silent).", count=noop_count)

    entry = {
        "operation_id": save_info.get("operation_id"),
        "phase": ("3b" if (report.get("expansion") or {}).get("applied") else "3a"),
        "execute_version": EXECUTE_VERSION,
        "plan_id": plan_id,
        "plan_hash": f"sha256:{plan_id}",
        "profile_id": profile.get("profile_id"),
        "source_content_fingerprint":
            (source.get("meta") or {}).get("content_fingerprint"),
        "target_fingerprint_before": report.get("target_file_fingerprint"),
        "target_fingerprint_after": target_after,
        "region_fingerprint": region_after,
        "region_fingerprint_before": report.get("region_fingerprint"),
        "region_fingerprint_after": region_after,
        "approval_state": (report.get("approval") or {}).get("state"),
        "approval_source": approval_source,
        "status": "committed",
        "idempotency_status": (report.get("idempotency") or {}).get("status"),
        "output": str(final),
        "in_place": bool(in_place),
        "backup": backup_path,
        "changes_count": changes_count,
        "noop_count": noop_count,
        "warnings": sorted({item.get("code")
                            for item in report.get("warnings") or []}
                           | {item.get("code")
                              for item in qa.get("warnings") or []}),
        "validation_status": qa.get("validation_status"),
        "qa_status": qa.get("qa_status"),
        "original_unchanged": original_unchanged,
    }
    manifest_info = record_execution(entry, manifest_dir)

    payload = {
        "ok": True, "mode": "execute", "command": "execute",
        "execute_version": EXECUTE_VERSION,
        "status": "committed",
        "idempotency_status": (report.get("idempotency") or {}).get("status"),
        "output": str(final),
        "write_count": len(writes),
        "changes_count": changes_count,
        "noop_count": noop_count,
        "cells": sorted(planned)[:CELL_LIST_CAP],
        "sheets": report.get("sheets"),
        "plan_id": plan_id,
        "profile_id": profile.get("profile_id"),
        "operation_id": save_info.get("operation_id"),
        "target_fingerprint_before": report.get("target_file_fingerprint"),
        "target_fingerprint_after": target_after,
        "region_fingerprint_before": report.get("region_fingerprint"),
        "region_fingerprint_after": region_after,
        "approval": report.get("approval"),
        "approval_source": approval_source,
        "backup": backup_path,
        "expansion": (exp_state if isinstance(exp_state, dict) and
                      exp_state.get("applied") else None),
        "original_unchanged": original_unchanged,
        "validation": qa.get("validation"),
        "qa": {"status": qa.get("qa_status"),
               "unexpected": qa.get("unexpected")[:20],
               "unexpected_count": len(qa.get("unexpected") or []),
               "formula_count_before": qa.get("formula_count_before"),
               "formula_count_after": qa.get("formula_count_after")},
        "manifest": manifest_info,
        "staged": True,
        "atomic": placed.get("atomic"),
        "same_filesystem": placed.get("same_filesystem"),
        "stage": "done",
        "commit_state": state.get("commit_state", "COMMITTED"),
        "timings_s": {
            "preflight": round(write_started - started, 3),
            "write": round(write_done - write_started, 3),
            "qa": round(qa_done - write_done, 3),
            "commit": round(commit_done - qa_done, 3),
            "total": round(time.perf_counter() - started, 3),
        },
    }
    return payload


def post_validate(staging_path, before, writes, boxes, *, result=None) -> dict:
    """Deep QA on the staged workbook (values/formulas/styles/structures)."""
    common.fault("reopen_validate", "before re-opening the staged output",
                 staging=str(staging_path))
    after_wb = common.load_workbook_safe(staging_path)
    common.fault("qa", "before the QA comparison", staging=str(staging_path))
    planned = {cell_key(sheet, cell) for sheet, cell, _v, _e in writes}
    after = snapshot(after_wb, style_cells=planned, result=result)
    comparison = compare_snapshots(before, after, planned)
    region_after = region_fingerprint_of(after_wb, boxes) if boxes else None
    unexpected = comparison["unexpected"]
    formulas_preserved = (comparison["formula_count_before"] ==
                          comparison["formula_count_after"])
    structure_ok = not any(str(item.get("kind", "")).startswith("structure:")
                           for item in unexpected)
    style_ok = not any(item.get("kind") == "style" for item in unexpected)
    validation_status = "pass" if (formulas_preserved and structure_ok) \
        else "fail"
    qa_status = "pass" if not unexpected else "fail"
    warnings = []
    if comparison["noop"]:
        warnings.append({"code": "CELL_ALREADY_EQUAL",
                         "message": f"{len(comparison['noop'])} planned "
                                    "cell(s) already held the planned value"})
    return {
        "ok": validation_status == "pass" and qa_status == "pass",
        "validation_status": validation_status,
        "qa_status": qa_status,
        "validation": {
            "reopens": True,
            "formulas_preserved": formulas_preserved,
            "structure_preserved": structure_ok,
            "styles_preserved": style_ok,
        },
        "formula_count_before": comparison["formula_count_before"],
        "formula_count_after": comparison["formula_count_after"],
        "changed": comparison["changed"],
        "noop": comparison["noop"],
        "unexpected": unexpected,
        "region_fingerprint_after": region_after,
        "warnings": warnings,
    }


def record_execution(entry: dict, manifest_dir=None) -> dict:
    """Append the FAZ 3A operation record to the Faz 1 manifest."""
    return common.manifest_append(entry, skill_dir=manifest_dir)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv=None):
    parser = argparse.ArgumentParser(
        description="FAZ 3A: execute an APPROVED fill plan on a template "
                    "workbook. Default output is a new file; in-place "
                    "requires --in-place, an approval token and a backup. "
                    "Row expansion, formula propagation and lookup execution "
                    "are enabled when policy allow_expansion=true and the "
                     "approval token is provided (in-place mandates backup).")
    parser.add_argument("--fill-plan", required=True, metavar="PATH",
                        help="fill plan JSON from xlsx_mapping.py --emit "
                             "dry-run")
    parser.add_argument("--profile", required=True, metavar="PATH",
                        help="the same template profile the plan was built "
                             "from")
    parser.add_argument("--source", required=True, metavar="PATH",
                        help="source table (CSV/JSON) with the real values")
    parser.add_argument("--target", required=True, metavar="PATH",
                        help="target workbook (.xlsx)")
    parser.add_argument("--emit", choices=("preflight", "execute"),
                        default="execute")
    parser.add_argument("--out", metavar="PATH",
                        help=f"output workbook (default: "
                             f"<target>{DEFAULT_OUT_SUFFIX}.xlsx)")
    parser.add_argument("--in-place", action="store_true",
                        help="overwrite the target itself (approval token + "
                             "backup are mandatory)")
    parser.add_argument("--approve-token", metavar="TOKEN",
                        help="first 12 characters of the plan_id; proves the "
                             "plan was approved before execution")
    parser.add_argument("--approval-source", default="cli", metavar="TEXT",
                        help="recorded in the manifest (e.g. hermes-agent)")
    parser.add_argument("--manifest-dir", metavar="PATH",
                        help="override the manifest directory (tests)")
    parser.add_argument("--pretty", action="store_true")
    args = parser.parse_args(argv)

    result = common.Result()
    plan = load_fill_plan(args.fill_plan, result)
    profile = semantics.load_profile(args.profile, result)
    source = mapping.load_source(args.source, result)
    target = Path(args.target)
    if not target.exists():
        raise common.XlsxError(
            "FILE_NOT_FOUND", f"target workbook not found: {target}",
            context={"target": str(target)})

    out_path = resolve_output(target, args.out, args.in_place)
    collect = {}
    wb = common.load_workbook_safe(target)
    loaded_wb = wb

    # FAZ 3B: expansion gate (read-only structural evaluation)
    policy = plan.get("policy") or {}
    exp_state = None
    if policy.get("allow_expansion") is True:
        exp_state = expand.build_expansion_state(
            plan, profile, source, wb.wb,  # raw workbook for read-only gate
            authorized=True,
            result=common.Result(mode="expansion_gate"))

    def run_preflight(extra_state=None):
        return preflight_execution(
            plan, profile, source, loaded_wb, target_path=target,
            out_path=(None if args.in_place else out_path),
            in_place=bool(args.in_place),
            approve_token=args.approve_token,
            manifest_dir=args.manifest_dir, result=result,
            collect_writes=collect, expansion_state=extra_state)

    try:
        report = run_preflight(None)
        if exp_state and exp_state.get("needed"):
            # Evaluate + apply in memory, then re-run preflight on the
            # EXPANDED workbook so live coordinates match the post-expansion
            # layout (D1: logical -> actual A1).
            temp_exp_plan = expand.plan_expansion(
                plan, profile, source, loaded_wb.wb,
                result=common.Result(mode="expansion_plan"))
            if temp_exp_plan.get("ok"):
                expand.apply_expansion(loaded_wb, temp_exp_plan,
                                       result=common.Result(mode="expansion"))
                # re-run preflight on the EXPANDED workbook
                report = run_preflight({
                    **exp_state, "applied": True, "plan": temp_exp_plan})
            else:
                report = run_preflight({
                    **exp_state, "applied": False,
                    "blockers": temp_exp_plan.get("blockers", [])})
                _raise_if_blocked(report)

        if args.emit == "preflight":
            payload = {
                "ok": True, "mode": "execute", "command": "preflight",
                "execute_version": EXECUTE_VERSION,
                "plan_id": plan.get("plan_id"),
                "profile_id": profile.get("profile_id"),
                "report": report,
            }
            return mapping.emit(payload, result, args.pretty)
        _raise_if_blocked(report)
        writes = collect.get("writes") or []
        if not writes:
            payload = {
                "ok": True, "mode": "execute", "command": "execute",
                "execute_version": EXECUTE_VERSION,
                "status": "no_writes", "write_count": 0,
                "plan_id": plan.get("plan_id"),
                "report": report,
            }
            return mapping.emit(payload, result, args.pretty)
        payload = execute_fill_plan(
            plan, profile, source, loaded_wb, report, writes,
            target_path=target, out_path=args.out,
            in_place=bool(args.in_place),
            approve_token=args.approve_token,
            manifest_dir=args.manifest_dir,
            approval_source=args.approval_source, result=result)
        return mapping.emit(payload, result, args.pretty)
    finally:
        wb.close()


if __name__ == "__main__":
    sys.exit(common.guard(main)())


# ---------------------------------------------------------------------------
# FAZ 4 / step 6b -- approved formula-synthesis plan execution (D7-A).
#
# This is a THIN adapter over the Faz 1/3C safe-write chain. Every formula
# text in a plan was already validated in `xlsx_formula_synth` (translation +
# signature re-derivation); nothing here computes or infers a formula -- it
# only stages writes, runs pre/post QA and commits atomically. Unknown plan
# types are rejected LOUDLY at admission (validated lesson: an unrecognized
# plan must die at admission, not mid-write).
# ---------------------------------------------------------------------------

SYNTH_PLAN_TYPE = "formula_synthesis"
SYNTH_WRITE_POLICY = "write_formula"


def load_synthesis_plan(path) -> dict:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def _assert_synthesis_plan(plan) -> list:
    contract = plan.get("executor_contract") or {}
    if contract.get("plan_type") != SYNTH_PLAN_TYPE:
        raise common.XlsxError(
            "PLAN_INVALID",
            f"plan_type {contract.get('plan_type')!r} is not a "
            f"formula synthesis plan",
            recovery=("Pass a plan produced by the synthesis planner "
                      "(plan_type=formula_synthesis)."))
    if plan.get("write_policy") != SYNTH_WRITE_POLICY:
        raise common.XlsxError(
            "PLAN_INVALID",
            f"write_policy {plan.get('write_policy')!r} not allowed for "
            f"synthesis execution",
            recovery="Only write_formula plans are executable here.")
    actions = plan.get("actions") or []
    if len(actions) != plan.get("generated_formula_count"):
        raise common.XlsxError(
            "PLAN_INVALID",
            "generated_formula_count does not match the action list",
            recovery="Re-plan the synthesis; the plan is inconsistent.")
    return actions


def execute_synthesis_plan(plan, *, target_path, out_path=None,
                           in_place=False, approve_token=None,
                           manifest_dir=None, result=None) -> dict:
    """Execute an approved synthesis plan: stage -> QA -> atomic commit.

    Guarantees (spec section 24; AC-4.21/4.22/4.24):
      * the original file is never touched before `atomic_replace`;
      * any failure leaves partial commit = 0 and the original file intact
        (a failed in-place commit-verification rolls back from the backup);
      * `values_recalculated` stays False -- openpyxl does not evaluate
        formulas and nothing here claims otherwise.
    """
    state = {
        "stage": "admission",
        "commit_state": "NOT_STARTED",
        "operation_id": common._new_operation_id("synth"),
        "plan_id": str(plan.get("plan_id") or ""),
        "backup": None,
        "temp": None,
        "write_count": 0,
        "in_place": bool(in_place),
    }
    try:
        if in_place and approve_token != plan.get("plan_id"):
            raise common.XlsxError(
                "APPROVAL_REQUIRED",
                "an in-place synthesis write needs the plan's approval token",
                recovery=("Re-run with --approve <plan_id> after reviewing "
                          "the plan, or write to --out instead."))
        return _execute_synthesis_body(plan, target_path=target_path,
                                       out_path=out_path, in_place=in_place,
                                       manifest_dir=manifest_dir,
                                       result=result, state=state)
    except common.XlsxError as exc:
        state["commit_state"] = "FAILED"
        recorded = record_failure(plan, state, exc, report=None,
                                  manifest_dir=manifest_dir)
        raise common.XlsxError(
            exc.code, str(exc), recovery=exc.recovery,
            context={**(exc.context or {}),
                     "stage": state["stage"],
                     "operation_id": state["operation_id"],
                     "commit_state": "FAILED",
                     "plan_id": state["plan_id"],
                     "backup": state.get("backup") or "NOT_AVAILABLE",
                     "temp": state.get("temp") or "NOT_AVAILABLE",
                     "committed": False,
                     "write_count": 0,
                     "values_recalculated": False,
                     "failure_manifest": recorded["manifest"]}) from exc
    except Exception as exc:                        # noqa: BLE001
        state["commit_state"] = "FAILED"
        wrapped = common.XlsxError(
            "UNEXPECTED_ERROR",
            f"{type(exc).__name__} at stage '{state['stage']}': {exc}")
        recorded = record_failure(plan, state, wrapped, report=None,
                                  manifest_dir=manifest_dir)
        raise common.XlsxError(
            "UNEXPECTED_ERROR", str(wrapped), recovery=wrapped.recovery,
            context={"stage": state["stage"],
                     "operation_id": state["operation_id"],
                     "commit_state": "FAILED",
                     "plan_id": state["plan_id"],
                     "backup": state.get("backup") or "NOT_AVAILABLE",
                     "temp": state.get("temp") or "NOT_AVAILABLE",
                     "committed": False,
                     "write_count": 0,
                     "values_recalculated": False,
                     "failure_manifest": recorded["manifest"]}) from exc


def _execute_synthesis_body(plan, *, target_path, out_path, in_place,
                            manifest_dir, result, state) -> dict:
    actions = _assert_synthesis_plan(plan)
    state["stage"] = "load"
    final = resolve_output(target_path, out_path, in_place)
    # `load_workbook_safe` returns the Pass-C LoadedWorkbook write authority;
    # snapshots and in-memory writes go through the RAW openpyxl workbook
    # (never read_only, never data_only -- spec 14/15).
    loaded = common.load_workbook_safe(target_path, data_only=False,
                                       read_only=False)
    wb = _raw(loaded)
    planned = {cell_key(action["target"]["sheet"], action["target"]["cell"])
               for action in actions}

    state["stage"] = "precheck"
    for action in actions:
        sheet = action["target"]["sheet"]
        cell = action["target"]["cell"]
        if sheet not in wb.sheetnames:
            raise common.XlsxError(
                "SHEET_NOT_FOUND", f"sheet {sheet!r} vanished since planning",
                recovery="Re-plan against the current workbook.")
        current = wb[sheet][cell].value
        if isinstance(current, str) and current.startswith("="):
            raise common.XlsxError(
                "FORMULA_MODIFICATION_FORBIDDEN",
                f"{sheet}!{cell} became a formula since planning; refusing "
                "to overwrite it",
                recovery="Re-plan; the target is no longer empty.")

    state["stage"] = "write"
    before = snapshot(wb, style_cells=planned, result=result)
    for action in actions:
        wb[action["target"]["sheet"]][action["target"]["cell"]].value = (
            action["formula_text"])

    state["stage"] = "staging"
    final.parent.mkdir(parents=True, exist_ok=True)
    handle, staging_name = tempfile.mkstemp(
        prefix=f".{final.stem}.stage-", suffix=".xlsx", dir=str(final.parent))
    os.close(handle)
    state["temp"] = staging_name
    # The ONLY sanctioned write path (Faz 1 gate): save_workbook_safe writes
    # the staging file through its own backup/temp/validate chain. Raw .save()
    # is forbidden in this module (scope-guard AST test) and unused here; the
    # staging file reaches the target through `atomic_replace` below.
    common.save_workbook_safe(loaded, out_path=staging_name, backup=False,
                              record_manifest=False)

    state["stage"] = "qa"
    staged = _raw(common.load_workbook_safe(staging_name, data_only=False,
                                            read_only=False))
    qa = compare_snapshots(before,
                           snapshot(staged, style_cells=planned,
                                    result=result),
                           planned)
    # `compare_snapshots` is the Faz 3A VALUE-fill comparator: it treats any
    # formula change as unexpected because fill plans never write formulas.
    # A synthesis plan writes formulas AT its planned cells by definition, so
    # entries inside the planned set are expected here -- every entry OUTSIDE
    # it keeps its full strictness (AC-4.24: unrelated formulas never change).
    qa["unexpected"] = [entry for entry in qa.get("unexpected", [])
                        if getattr(entry, "get", lambda *_: {})("cell")
                        not in planned]
    texts_ok = all(
        staged[action["target"]["sheet"]][action["target"]["cell"]].value
        == action["formula_text"] for action in actions)
    staged.close()
    count_before = qa.get("formula_count_before")
    count_after = qa.get("formula_count_after")
    delta = (count_after - count_before
             if isinstance(count_before, int) and isinstance(count_after, int)
             else None)
    qa["formula_count_delta"] = delta
    count_ok = delta == len(actions)
    if qa["unexpected"] or not texts_ok or not count_ok:
        raise common.XlsxError(
            "QA_FAILED",
            "pre-commit QA refused the staged workbook",
            recovery=("Nothing was committed; the original file is "
                      "untouched. Fix the plan and re-run."),
            context={"qa_unexpected": qa["unexpected"][:20],
                     "texts_ok": texts_ok,
                     "formula_count_delta": delta,
                     "formula_count_delta_expected": len(actions),
                     "staging": staging_name,
                     "committed": False, "write_count": 0})

    state["stage"] = "commit"
    backup_path = None
    if in_place:
        backup_path = common.backup_workbook(target_path, tag="3c")["backup"]
        state["backup"] = backup_path
    common.atomic_replace(staging_name, final)
    state["temp"] = None
    state["commit_state"] = "COMMITTED"
    final_check = _post_commit_verify(final, expect_sheets=list(wb.sheetnames))
    if not final_check.get("ok"):
        rolled_back = False
        if in_place and backup_path:
            common.atomic_replace(backup_path, target_path)
            rolled_back = True
        raise common.XlsxError(
            "WRITE_FAILED",
            "committed file failed re-verification: "
            f"{final_check.get('reason')}",
            recovery=("The backup was restored." if rolled_back else
                      "The original file was never touched."))
    state["write_count"] = len(actions)
    report = {
        "status": "ok",
        "plan_id": state["plan_id"],
        "committed": True,
        "write_count": len(actions),
        "output": str(final),
        "backup": backup_path,
        "qa": qa,
        "values_recalculated": False,
        "unexpected_changes": qa.get("unexpected", []),
        "formula_count_delta": qa.get("formula_count_delta"),
        "formula_count_delta_expected": len(actions),
    }
    entry = {"operation_id": state["operation_id"],
             "plan_id": state["plan_id"], "status": "ok",
             "target": str(target_path), "output": str(final),
             "write_count": len(actions),
             "commit_state": state["commit_state"]}
    entry.update(record_execution(entry, manifest_dir=manifest_dir) or {})
    return report
