#!/usr/bin/env python3
"""FAZ 3B -- row expansion + formula/style/validation propagation engine.

Turns an APPROVED fill plan's planned-only row expansion into real
workbook structure, reusing the single shift engine of
``xlsx_restructure.py`` (``apply_shift``) and its reference parser.

Scope (binding):
  * row expansion inside the plan's evidenced block only (FAZ 3B-1)
  * formula propagation = copy-down translation of the formula that
    EXISTS in the template row (relative rows shift, ``$`` stays)
    -- never formula invention (FAZ 3B-2 / FAZ 4 boundary)
  * style / row-height / validation / conditional-format / table /
    merge / hidden propagation (FAZ 3B-3)
  * NOT in scope: formula family discovery, R1C1 engines, window
    classification, arbitrary synthesis (FAZ 4), semantic remapping.

The 3B execution locator model (D1): the dry-run's absolute A1 cells
are *preview evidence only*. Execution resolves logical locators
(block_id + row_offset + column) against the LIVE, post-expansion
workbook; a mismatch between the resolved coordinate and the plan's
preview cell is a hard blocker -- never a silent write.
"""
from __future__ import annotations

from copy import copy  # style-array copies (never shared styles)
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_common as common  # noqa: E402
import xlsx_mapping as mapping  # noqa: E402  (bounds helpers; no writes)
import xlsx_restructure as restructure  # noqa: E402  (shift engine+parser)

try:  # dependency guard: structured error, never a traceback
    from openpyxl.cell.cell import MergedCell  # noqa: F401
    from openpyxl.utils import get_column_letter, range_boundaries  # noqa: F401
except ImportError as _exc:  # pragma: no cover - environment guard
    sys.exit(common.dependency_error(_exc))

EXPAND_VERSION = "3b.1"

#: minimum evidenced record-key score for duplicate verification (D3).
#: Below this the key is treated as UNAVAILABLE -- no dedupe, no overwrite.
RECORD_KEY_MIN_SCORE = 60

#: OOXML rule types whose conditional formatting may be sqref-extended;
#: anything else is preserved and reported (PRESERVE_WITH_WARNING), never
#: guessed. openpyxl reports every CF rule as class ``Rule`` -- the stable
#: discriminator is ``rule.type`` (cellIs, expression, ...).
_TESTED_CF_RULE_TYPES = ("cellIs", "expression", "colorScale", "dataBar",
                         "iconSet")
_TESTED_DV_TYPES = ("list", "whole", "decimal", "date", "time",
                    "textLength", "custom")


def _masked(text: str) -> str:
    """Formula text with string literals blanked (parser parity with the
    restructure engine -- the only reference scanner in the codebase)."""
    chars = list(text)
    for literal in restructure.STRING_RE.finditer(text):
        for index in range(literal.start(), literal.end()):
            chars[index] = " "
    return "".join(chars)


def classify_formula(text):
    """(supported, reason) for copy-down propagation of a template formula.

    Supported = plain A1 references (relative/absolute/mixed, ranges,
    same-sheet and cross-sheet). Everything else is refused explicitly
    (spec section 11) -- never propagated by guess.
    """
    if not isinstance(text, str) or not text.startswith("="):
        return (False, "not a formula")
    masked = _masked(text)
    if "{" in masked or "}" in masked:
        return (False, "array or constant formula")
    if "_xlfn." in masked:
        return (False, "future function (_xlfn)")
    if "#REF!" in masked:
        return (False, "broken reference (#REF!)")
    if "[" in masked or "]" in masked:
        return (False, "structured or external reference")
    if masked.lstrip("=").lstrip().startswith(("+", "-", "*", "/")):
        return (False, "formula without a leading reference is not a "
                       "supported copy-down pattern")
    return (True, None)


def translate_formula(text: str, drow: int) -> str:
    """Copy-down translation: every relative ROW reference shifts by
    ``drow`` (any sheet), absolute rows keep their position; string
    literals are untouched. Columns are never touched (vertical copy).

    This is a mechanical application of ``restructure``'s parser --
    deliberately NOT a formula engine (no R1C1, no families).
    """
    out, pos = [], 0

    def sub(match):
        prefix = match.group("sheet") or ""

        def shift(coord):
            col_abs, col, row_abs, row = restructure.COORD_RE.match(
                coord).groups()
            new_row = row if row_abs else str(int(row) + drow)
            return f"{col_abs}{col}{row_abs}{new_row}"

        start, end = match.group("start"), match.group("end")
        if end is None:
            return prefix + shift(start)
        return f"{prefix}{shift(start)}:{shift(end)}"

    for literal in restructure.STRING_RE.finditer(text):
        out.append(restructure.REF_RE.sub(sub, text[pos:literal.start()]))
        out.append(literal.group(0))
        pos = literal.end()
    out.append(restructure.REF_RE.sub(sub, text[pos:]))
    return "".join(out)


def translate_ref_line(ref, drow):
    """Translate the row part of an A1 range string (no dollar flags)."""
    min_col, min_row, max_col, max_row = range_boundaries(ref)
    return (f"{get_column_letter(min_col)}{min_row + drow}:"
            f"{get_column_letter(max_col)}{max_row + drow}")


# ---------------------------------------------------------------------------
# block / capacity / record-key evidence
# ---------------------------------------------------------------------------

def find_block(profile, sheet_name, block_id):
    """The profile's block record (by id, else by sheet)."""
    for sheet in profile.get("sheets") or []:
        if sheet.get("name") != sheet_name:
            continue
        for block in sheet.get("blocks") or []:
            if block_id is not None and block.get("block_id") == block_id:
                return block
        return None
    return None


def block_bounds(block):
    return mapping._bounds_of_ref(block.get("range")) if block else None


def record_key_column(block):
    """Usable record key column letter, or None (D3: never a weak guess)."""
    candidate = (block or {}).get("record_key_candidate") or {}
    score = candidate.get("candidate_key_score") or 0
    column = candidate.get("column")
    if column and score >= RECORD_KEY_MIN_SCORE:
        return column
    return None


def _column_values(ws, letter, first_row, last_row):
    values = []
    for row in range(first_row, last_row + 1):
        value = ws[f"{letter}{row}"].value
        values.append(None if value is None else str(value))
    return values


def plan_expansion(fill_plan, profile, source, wb, result=None) -> dict:
    """Read-only, deterministic structural plan (spec sections 4/5/6).

    Returns per-block: boundary, rows_to_add, insert_at, template row,
    affected columns, the live inventories the expansion will touch and
    every check (merge safety, formula support, table mode, record key
    / duplicate evidence). Nothing is modified; nothing is guessed.
    """
    expansion = (fill_plan.get("expansion") or {})
    blocks_evidence = expansion.get("blocks") or []
    plan = {
        "version": EXPAND_VERSION,
        "ok": True,
        "blocks": [],
        "blockers": [],
        "warnings": [],
        "rows_added_total": 0,
    }

    def block_error(code, message, **context):
        plan["ok"] = False
        plan["blockers"].append({"code": code, "message": message,
                                 **context})

    if not blocks_evidence:
        plan["warnings"].append({"code": "NO_EXPANSION_EVIDENCE",
                                 "message": "the fill plan carries no "
                                            "expansion block evidence; "
                                            "nothing to expand."})
        return plan

    rows = source.get("rows") or []
    for evidence in blocks_evidence:
        sheet_name = evidence.get("sheet")
        if not sheet_name or sheet_name not in wb.sheetnames:
            block_error("EXPANSION_BOUNDARY_UNKNOWN",
                        f"expansion block sheet {sheet_name!r} is not in "
                        "the target workbook.", sheet=sheet_name)
            continue
        block = find_block(profile, sheet_name, evidence.get("block_id"))
        bounds = block_bounds(block)
        if bounds is None:
            block_error("EXPANSION_BOUNDARY_UNKNOWN",
                        f"no profile block range for "
                        f"{sheet_name}!{evidence.get('block_id')!r}; the "
                        "boundary is never guessed.", sheet=sheet_name)
            continue
        anchor_row = evidence.get("anchor_row")
        if bounds[1] != anchor_row:
            block_error("EXPANSION_BOUNDARY_UNKNOWN",
                        "the plan anchor row does not match the profile "
                        "block range start.",
                        sheet=sheet_name, anchor_row=anchor_row,
                        block_start=bounds[1])
            continue

        planned_rows = int(evidence.get("planned_rows") or 0)
        existing_rows = bounds[3] - bounds[1] + 1
        rows_to_add = max(0, planned_rows - existing_rows)
        insert_at = bounds[3] + 1
        new_last_row = anchor_row + planned_rows - 1 if planned_rows else \
            bounds[3]
        columns = [c for c in (evidence.get("columns") or []) if c]
        ws = wb[sheet_name]
        records = {
            "sheet": sheet_name,
            "block_id": evidence.get("block_id"),
            "anchor_row": anchor_row,
            "block_range": block.get("range"),
            "existing_rows": existing_rows,
            "planned_rows": planned_rows,
            "rows_to_add": rows_to_add,
            "insert_at": insert_at,
            "new_last_row": new_last_row,
            "template_row": bounds[3],
            "columns": sorted(set(columns)),
            "column_span": (get_column_letter(bounds[0]),
                            get_column_letter(bounds[2])),
            "checks": {},
            "warnings": [],
            "blockers": [],
        }

        # -- merge safety: nothing may straddle the insertion boundary ----
        straddling = []
        for merged in ws.merged_cells.ranges:
            if merged.min_row < insert_at <= merged.max_row:
                straddling.append(str(merged))
        records["checks"]["merge_safe"] = not straddling
        if straddling:
            records["blockers"].append({
                "code": "MERGE_EXPANSION_UNSUPPORTED",
                "message": "merged ranges straddle the insertion boundary; "
                           "the merge pattern is ambiguous.",
                "ranges": straddling[:10]})

        # -- template formulas: copy-down support --------------------------
        formula_support = {"supported": [], "unsupported": []}
        for column in range(bounds[0], bounds[2] + 1):
            cell = ws.cell(row=bounds[3], column=column)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                ok, reason = classify_formula(cell.value)
                item = {"cell": cell.coordinate, "formula": cell.value}
                if ok:
                    formula_support["supported"].append(item)
                else:
                    formula_support["unsupported"].append(
                        dict(item, reason=reason))
        records["checks"]["formula_support"] = formula_support
        if formula_support["unsupported"]:
            records["blockers"].append({
                "code": "UNSUPPORTED_FORMULA_PROPAGATION",
                "message": "the template row holds formulas that cannot be "
                           "propagated by copy-down translation; the "
                           "expansion stops instead of guessing.",
                "formulas": formula_support["unsupported"][:10]})

        # -- table mode -----------------------------------------------------
        table_info = []
        for table in ws.tables.values():
            t_bounds = mapping._bounds_of_ref(table.ref)
            if not t_bounds:
                continue
            intersects = (t_bounds[1] <= bounds[3] and
                          t_bounds[3] >= bounds[1] and
                          t_bounds[0] <= bounds[2] and
                          t_bounds[2] >= bounds[0])
            if not intersects:
                continue
            mode = "report_only"
            if getattr(table, "totalsRowShown", False) or \
                    getattr(table, "totalsRowCount", 0):
                mode = "unsupported_totals"
            elif t_bounds[3] == bounds[3] and t_bounds[1] <= anchor_row:
                mode = "extend"
            elif t_bounds[3] <= insert_at - 1:
                mode = "shift_only"
            table_info.append({"name": table.displayName, "ref": table.ref,
                               "mode": mode})
            if mode == "unsupported_totals":
                records["blockers"].append({
                    "code": "TABLE_EXPANSION_UNSUPPORTED",
                    "message": "table with a totals row intersects the "
                               "expansion; not tested -- refused, not "
                               "guessed.",
                    "table": table.displayName})
        records["checks"]["tables"] = table_info

        # -- aggregate pointer hint (informational; never auto-changes) ----
        if rows_to_add > 0:
            aggregate_hints = []
            scan_end = min(ws.max_row, bounds[3] + 50)
            for row in range(insert_at, scan_end + 1):
                for column in range(bounds[0], bounds[2] + 1):
                    value = ws.cell(row=row, column=column).value
                    if not (isinstance(value, str)
                            and value.startswith("=")):
                        continue
                    masked = _masked(value)
                    for match in restructure.REF_RE.finditer(masked):
                        if match.group("end") is None:
                            continue
                        start_b = mapping._bounds_of_ref(
                            match.group("start"))
                        end_b = mapping._bounds_of_ref(match.group("end"))
                        if not start_b or not end_b:
                            continue
                        if end_b[3] == bounds[3] and \
                                start_b[1] >= bounds[1]:
                            aggregate_hints.append({
                                "cell": ws.cell(row=row,
                                                column=column).coordinate,
                                "formula": value})
                            break
            if aggregate_hints:
                records["warnings"].append({
                    "code": "AGGREGATE_MAY_NOT_COVER_NEW_ROWS",
                    "message": "a formula below the block aggregates a "
                               "range ending at the template row; it is "
                               "left unchanged on purpose (never rewritten "
                               "by guess) -- the new rows may not be "
                               "covered.",
                    "items": aggregate_hints[:5]})

        # -- validation / conditional-format coverage ----------------------
        dv_hits, cf_hits = [], []
        for dv in ws.data_validations.dataValidation:
            for part in str(dv.sqref).split():
                db = mapping._bounds_of_ref(part)
                if not db:
                    continue
                if db[1] <= bounds[3] <= db[3] and \
                        not (db[2] < bounds[0] or db[0] > bounds[2]):
                    dv_hits.append({"type": dv.type, "sqref": str(dv.sqref),
                                    "covered": True})
                    break
        for cf in ws.conditional_formatting:
            for part in str(cf.sqref).split():
                cb = mapping._bounds_of_ref(part)
                if not cb:
                    continue
                if cb[1] <= bounds[3] <= cb[3] and \
                        not (cb[2] < bounds[0] or cb[0] > bounds[2]):
                    rules = [rule.type for rule in cf.rules]
                    cf_hits.append({"sqref": str(cf.sqref), "rules": rules,
                                    "tested": all(
                                        name in _TESTED_CF_RULE_TYPES
                                        for name in rules)})
                    break
        records["checks"]["validations"] = dv_hits
        records["checks"]["conditional_formats"] = cf_hits

        # -- row state ------------------------------------------------------
        dim = ws.row_dimensions.get(bounds[3])
        records["checks"]["template_row_state"] = {
            "hidden": bool(getattr(dim, "hidden", False)),
            "outline_level": (int(getattr(dim, "outlineLevel", 0) or 0) or
                              None),
            "height": (float(dim.height) if dim is not None and
                       dim.height is not None else None),
        }

        # -- record key / duplicates (D3, spec section 29) -----------------
        key_letter = record_key_column(block)
        rc = {"key_column": key_letter,
              "state": ("unavailable" if key_letter is None else "available")}
        if rows_to_add > 0:
            source_key_name, key_index = None, None
            if key_letter:
                for entry in fill_plan.get("would_write") or []:
                    if (entry.get("sheet") == sheet_name and
                            str(entry.get("cell", "")).rstrip(
                                "0123456789").upper() == key_letter.upper()):
                        source_key_name = str(entry.get("source", "")
                                              ).split("[", 1)[0]
                        break
                header_names = [str(name) for name in
                                (source.get("headers") or [])]
                if not header_names:
                    for column in ((source.get("meta") or {}).get("columns")
                                   or []):
                        header_names.append(
                            str(column.get("key"))
                            if isinstance(column, dict) else str(column))
                if source_key_name in header_names:
                    key_index = header_names.index(source_key_name)
            if not key_letter or key_index is None:
                rc["state"] = "unavailable"
                rc["reason"] = ("no evidenced record key"
                                if not key_letter else
                                "the key column is not mapped from the "
                                "source table")
                records["blockers"].append({
                    "code": "RECORD_KEY_UNAVAILABLE",
                    "message": "the record key cannot be verified "
                               "(key column: " + str(key_letter) + "); "
                               "existing records may not be overwritten "
                               "and duplicates cannot be detected.",
                })
            else:
                existing_keys = [k for k in
                                 _column_values(ws, key_letter, bounds[1],
                                                bounds[3]) if k is not None]
                source_keys = []
                for row_values in rows[:planned_rows]:
                    value = (row_values[key_index]
                             if key_index < len(row_values) else None)
                    source_keys.append("" if value is None else str(value))
                append_zone = source_keys[existing_rows:]
                duplicates = sorted(set(k for k in append_zone
                                        if k and k in set(existing_keys)))
                internal = sorted({k for k in source_keys
                                   if k and source_keys.count(k) > 1})
                rc["source_key"] = source_key_name
                rc["existing_distinct"] = len(set(existing_keys))
                rc["duplicates_in_target"] = duplicates[:20]
                rc["duplicates_in_source"] = internal[:20]
                if duplicates:
                    rc["state"] = "duplicate_record"
                    records["blockers"].append({
                        "code": "DUPLICATE_RECORD",
                        "message": "source record keys already exist in "
                                   "the target block (append zone); "
                                   "appending would duplicate them -- "
                                   "never overwritten by default.",
                        "keys": duplicates[:10]})
                if internal:
                    rc["state"] = "duplicate_record"
                    records["blockers"].append({
                        "code": "DUPLICATE_RECORD",
                        "message": "the source table itself holds "
                                   "duplicate record keys; the correct "
                                   "row cannot be chosen "
                                   "deterministically.",
                        "keys": internal[:10]})
        records["checks"]["record_key"] = rc

        if evidence.get("needed") is True and rows_to_add == 0:
            records["warnings"].append({
                "code": "ROW_EXPANSION_FLAG_RECONCILED",
                "message": "the dry-run flagged expansion, but the live "
                           "block capacity covers all planned records; "
                           "no rows are added (live evidence wins).",
                "planned_rows": planned_rows,
                "existing_rows": existing_rows})
        if evidence.get("needed") in (False, None) and rows_to_add > 0:
            records["warnings"].append({
                "code": "ROW_EXPANSION_FLAG_RECONCILED",
                "message": "the dry-run did not flag expansion, but the "
                           "live block capacity is insufficient; the "
                           "expansion is planned from live evidence.",
                "planned_rows": planned_rows,
                "existing_rows": existing_rows})

        if records["blockers"]:
            plan["ok"] = False
            plan["blockers"].extend(
                [dict(item, sheet=sheet_name,
                      block_id=evidence.get("block_id"))
                 for item in records["blockers"]])
        plan["warnings"].extend(
            [dict(item, sheet=sheet_name, block_id=evidence.get("block_id"))
             for item in records["warnings"]])
        plan["rows_added_total"] += rows_to_add
        plan["blocks"].append(records)

    if result is not None:
        for item in plan["warnings"]:
            extra = ({"sheet": item["sheet"]} if item.get("sheet")
                     else {})
            result.warn(item.get("code", "EXPANSION_WARNING"),
                        item.get("message", ""), **extra)
        for item in plan["blockers"]:
            result.diagnose(f"expansion blocker {item['code']}: "
                            f"{item.get('message', '')[:160]}")
    return plan


# ---------------------------------------------------------------------------
# execution locators (D1): logical -> actual A1, resolved live
# ---------------------------------------------------------------------------

_CELL_RE = re.compile(r"^([A-Za-z]{1,3})([0-9]{1,7})$")


def split_cell(cell):
    match = _CELL_RE.match(str(cell or ""))
    if not match:
        return None
    return match.group(1).upper(), int(match.group(2))


def resolve_write_locators(fill_plan, expansion_plan) -> dict:
    """Resolve every planned write's logical locator against the LIVE
    layout AFTER the planned expansion (D1).

    * entries inside a block: resolved = anchor + row_offset (the
      coordinate only becomes a provisioned row once the expansion has
      run -- ``requires_expansion`` marks the tail),
    * entries below an insertion point (any mode): resolved = original
      row + (rows added by insertions above it),
    * at most one expanding block per sheet (fail-closed otherwise).
    """
    records = expansion_plan.get("blocks") or []
    active = [b for b in records if (b.get("rows_to_add") or 0) > 0]
    blockers = []
    per_sheet = {}
    for block in active:
        per_sheet.setdefault(block["sheet"], []).append(block)
    for sheet, blocks in sorted(per_sheet.items()):
        if len(blocks) > 1:
            blockers.append({
                "code": "EXPANSION_MULTI_BLOCK_UNSUPPORTED",
                "message": f"{sheet}: more than one block needs rows "
                           "added; multi-block expansion is not tested "
                           "and is refused (never guessed).",
                "sheet": sheet,
                "blocks": [b.get("block_id") for b in blocks]})

    writes = []
    for index, entry in enumerate(fill_plan.get("would_write") or []):
        sheet = str(entry.get("sheet") or "")
        cell = str(entry.get("cell") or "")
        parts = split_cell(cell)
        if parts is None or not sheet:
            blockers.append({"code": "PLAN_INVALID",
                             "message": "cell reference is not "
                                        "addressable.",
                             "entry": index, "cell": cell})
            continue
        letter, row = parts
        inside = None
        for block in records:
            if block["sheet"] != sheet:
                continue
            if letter in (block.get("columns") or []):
                last = block["anchor_row"] + block["planned_rows"] - 1
                if block["anchor_row"] <= row <= last:
                    inside = block
                    break
        if inside is not None:
            offset = row - inside["anchor_row"]
            shift = 0
            requires = offset >= inside["existing_rows"]
        else:
            offset = None
            shift = sum(b["rows_to_add"] for b in active
                        if b["sheet"] == sheet and b["insert_at"] <= row)
            requires = False
        writes.append({
            "index": index,
            "sheet": sheet,
            "entry_cell": cell,
            "resolved_cell": f"{letter}{row + shift}",
            "row_shift": shift,
            "block_id": (inside or {}).get("block_id"),
            "row_offset": offset,
            "requires_expansion": requires,
        })
    return {"version": EXPAND_VERSION,
            "ok": not blockers,
            "writes": writes,
            "blockers": blockers,
            "sheets": sorted({w["sheet"] for w in writes})}


# ---------------------------------------------------------------------------
# staged expansion (in-memory): the only mutating section
# ---------------------------------------------------------------------------

def _provision_rows(ws, record):
    """Copy the template row's style / height / hidden state into the
    freshly inserted rows (never shared-style mutation: every cell gets
    its own style-array copy)."""
    template_row = record["template_row"]
    span = (mapping._bounds_of_ref(record["block_range"]))
    first_col, last_col = span[0], span[2]
    new_rows = list(range(record["insert_at"],
                          record["insert_at"] + record["rows_to_add"]))
    cells = 0
    for row in new_rows:
        for column in range(first_col, last_col + 1):
            source = ws.cell(row=template_row, column=column)
            if isinstance(source, MergedCell):
                # style of a merged field lives on the range master
                for rng in ws.merged_cells.ranges:
                    if (rng.min_row <= template_row <= rng.max_row and
                            rng.min_col <= column <= rng.max_col):
                        source = ws.cell(row=rng.min_row,
                                         column=rng.min_col)
                        break
            target = ws.cell(row=row, column=column)
            target._style = copy(source._style)
            cells += 1
    template_dim = ws.row_dimensions.get(template_row)
    height = (template_dim.height if template_dim is not None else None)
    hidden = bool(getattr(template_dim, "hidden", False))
    outline = int(getattr(template_dim, "outlineLevel", 0) or 0)
    collapsed = bool(getattr(template_dim, "collapsed", False))
    for row in new_rows:
        dim = ws.row_dimensions[row]
        if height is not None:
            dim.height = height
        if hidden:
            dim.hidden = True
        if outline:
            dim.outlineLevel = outline
            if collapsed:
                dim.collapsed = True
    return {"cells": cells, "rows": new_rows,
            "height": height, "hidden": hidden, "outline_level": outline}


def _propagate_formulas(ws, record):
    common.fault("formula_propagation", "before propagating formulas",
                 sheet=getattr(ws, "title", None))
    """Copy-down translation of the formulas that exist in the template
    row (FAZ 3B-2). Unsupported formulas were refused in plan_expansion;
    a surprise here is a hard error, never a guess."""
    template_row = record["template_row"]
    span = mapping._bounds_of_ref(record["block_range"])
    propagated, skipped = [], []
    for offset, row in enumerate(range(record["insert_at"],
                                       record["insert_at"] +
                                       record["rows_to_add"]), start=1):
        drow = offset
        for column in range(span[0], span[2] + 1):
            source = ws.cell(row=template_row, column=column)
            value = source.value
            if not (isinstance(value, str) and value.startswith("=")):
                continue
            ok, reason = classify_formula(value)
            if not ok:
                raise common.XlsxError(
                    "UNSUPPORTED_FORMULA_PROPAGATION",
                    f"formula {value!r} at "
                    f"{source.coordinate} cannot be propagated: {reason}",
                    context={"cell": source.coordinate, "formula": value,
                             "reason": reason})
            target = ws.cell(row=row, column=column)
            target.value = translate_formula(value, drow)
            propagated.append({"cell": target.coordinate,
                               "from": value,
                               "formula": target.value})
    return {"count": len(propagated),
            "items": propagated[:50],
            "skipped": skipped}


def _extend_validations(ws, record):
    """Extend the sqref span of the template row's validations to cover
    the new rows -- formula1/formula2 are never reinterpreted."""
    template_row, new_last = record["template_row"], record["new_last_row"]
    span = mapping._bounds_of_ref(record["block_range"])
    extended, warnings = [], []
    for dv in ws.data_validations.dataValidation:
        if dv.type not in _TESTED_DV_TYPES:
            warnings.append({"code": "UNSUPPORTED_VALIDATION_PROPAGATION",
                             "message": f"validation type {dv.type!r} is "
                                        "not tested for propagation; the "
                                        "new rows stay outside it "
                                        "(never guessed).",
                             "sqref": str(dv.sqref)})
            continue
        parts, changed = [], False
        for part in str(dv.sqref).split():
            bounds = mapping._bounds_of_ref(part)
            if bounds is None:
                parts.append(part)
                continue
            overlap = not (bounds[2] < span[0] or bounds[0] > span[2])
            if overlap and bounds[1] <= template_row <= bounds[3] and \
                    bounds[3] < new_last:
                parts.append(f"{get_column_letter(bounds[0])}"
                             f"{bounds[1]}:{get_column_letter(bounds[2])}"
                             f"{new_last}")
                changed = True
            else:
                parts.append(part)
        if changed:
            old = str(dv.sqref)
            dv.sqref = " ".join(parts)
            extended.append({"from": old, "to": str(dv.sqref),
                             "type": dv.type})
    return {"extended": extended, "warnings": warnings}


def _extend_conditional_formats(ws, record):
    template_row, new_last = record["template_row"], record["new_last_row"]
    span = mapping._bounds_of_ref(record["block_range"])
    extended, warnings = [], []
    try:
        from openpyxl.formatting.formatting import ConditionalFormattingList
    except ImportError:  # pragma: no cover
        return {"extended": [], "warnings": []}
    rebuilt = ConditionalFormattingList()
    for cf in ws.conditional_formatting:
        rules = [rule.type for rule in cf.rules]
        tested = all(name in _TESTED_CF_RULE_TYPES for name in rules)
        parts, changed = [], False
        for part in str(cf.sqref).split():
            bounds = mapping._bounds_of_ref(part)
            if bounds is None:
                parts.append(part)
                continue
            overlap = not (bounds[2] < span[0] or bounds[0] > span[2])
            if tested and overlap and bounds[1] <= template_row <= bounds[3] \
                    and bounds[3] < new_last:
                parts.append(f"{get_column_letter(bounds[0])}"
                             f"{bounds[1]}:{get_column_letter(bounds[2])}"
                             f"{new_last}")
                changed = True
            else:
                parts.append(part)
        new_sqref = " ".join(parts)
        for rule in cf.rules:
            rebuilt.add(new_sqref, rule)
        if changed:
            extended.append({"from": str(cf.sqref), "to": new_sqref,
                             "rules": rules})
        elif not tested:
            warnings.append({"code": "PRESERVE_WITH_WARNING",
                             "message": "conditional-formatting rule "
                                        f"types {rules} are not tested "
                                        "for propagation; preserved "
                                        "with a warning.",
                             "sqref": str(cf.sqref)})
    ws.conditional_formatting = rebuilt
    return {"extended": extended, "warnings": warnings}


def _extend_tables(ws, record):
    template_row, new_last = record["template_row"], record["new_last_row"]
    extended = []
    for plan_entry in record["checks"].get("tables") or []:
        if plan_entry.get("mode") != "extend":
            continue
        table = ws.tables.get(plan_entry["name"])
        if table is None:
            continue
        bounds = mapping._bounds_of_ref(table.ref)
        if bounds is None or bounds[3] != template_row:
            continue
        new_ref = (f"{get_column_letter(bounds[0])}{bounds[1]}:"
                   f"{get_column_letter(bounds[2])}{new_last}")
        old_ref = table.ref
        table.ref = new_ref
        entry = {"name": table.displayName, "from": old_ref, "to": new_ref}
        auto = getattr(table, "autoFilter", None)
        if auto is not None and getattr(auto, "ref", None):
            a_bounds = mapping._bounds_of_ref(auto.ref)
            if a_bounds and a_bounds[3] == template_row:
                auto.ref = (f"{get_column_letter(a_bounds[0])}"
                            f"{a_bounds[1]}:{get_column_letter(a_bounds[2])}"
                            f"{new_last}")
                entry["autofilter_to"] = auto.ref
        extended.append(entry)
    return {"extended": extended}


def _replicate_merges(ws, record):
    """Replicate the template row's own one-row merges into every new row
    (multi-row merges over the block were refused in plan_expansion)."""
    template_row = record["template_row"]
    span = mapping._bounds_of_ref(record["block_range"])
    pattern = [rng for rng in list(ws.merged_cells.ranges)
               if rng.min_row == rng.max_row == template_row and
               rng.min_col >= span[0] and rng.max_col <= span[2]]
    replicated = []
    for row in range(record["insert_at"],
                     record["insert_at"] + record["rows_to_add"]):
        for rng in pattern:
            ref = (f"{get_column_letter(rng.min_col)}{row}:"
                   f"{get_column_letter(rng.max_col)}{row}")
            ws.merge_cells(ref)
            replicated.append(ref)
    return {"replicated": replicated, "pattern": [str(r) for r in pattern]}


def apply_expansion(loaded, expansion_plan, result=None) -> dict:
    """Perform the staged, in-memory expansion + propagation for every
    planned block (bottom-up per sheet so coordinates stay stable).

    The caller keeps full control of the write chain: this function only
    mutates the in-memory workbook of ``loaded`` -- saving, QA and the
    atomic commit stay in ``xlsx_execute`` (FAZ 3A safe chain, reused,
    never re-implemented).
    """
    wb = loaded.wb
    report = {"version": EXPAND_VERSION, "ok": True, "blocks": [],
              "warnings": [], "rows_added_total": 0}
    ordered = sorted(expansion_plan.get("blocks") or [],
                     key=lambda item: (item["sheet"], -item["insert_at"]))
    for record in ordered:
        if (record.get("rows_to_add") or 0) <= 0:
            continue
        ws = wb[record["sheet"]]
        shift_result = common.Result(
            mode="expand", sheet=record["sheet"], op="insert", axis="rows",
            index=record["insert_at"], count=record["rows_to_add"],
            formulas=[], merges=[], tables={}, defined_names={},
            validations=[], conditional_formats=[],
            not_shifted=["chart anchors", "images",
                         "conditional-format rule formulas"])
        restructure.apply_shift(ws, wb, "rows", record["insert_at"],
                                record["rows_to_add"], False, shift_result)
        provisioned = _provision_rows(ws, record)
        propagated = _propagate_formulas(ws, record)
        validations = _extend_validations(ws, record)
        cfs = _extend_conditional_formats(ws, record)
        tables = _extend_tables(ws, record)
        merges = _replicate_merges(ws, record)
        block_entry = {
            "sheet": record["sheet"],
            "block_id": record.get("block_id"),
            "template_row": record["template_row"],
            "block_range": record["block_range"],
            "insert_at": record["insert_at"],
            "rows_added": record["rows_to_add"],
            "new_rows": [record["insert_at"],
                         record["insert_at"] + record["rows_to_add"] - 1],
            "new_last_row": record["new_last_row"],
            "provisioned_cells": provisioned["cells"],
            "row_state": {"height": provisioned["height"],
                          "hidden": provisioned["hidden"],
                          "outline_level": provisioned["outline_level"]},
            "propagated_formulas": propagated["count"],
            "formula_samples": propagated["items"][:5],
            "validations_extended": validations["extended"],
            "conditional_formats_extended": cfs["extended"],
            "tables_extended": tables["extended"],
            "merges_replicated": merges["replicated"],
            "shifted_references": {
                "formulas": len(shift_result.get("formulas") or []),
                "merges": len(shift_result.get("merges") or []),
                "defined_names": len(shift_result.get("defined_names")
                                     or {}),
            },
        }
        for warning in validations["warnings"] + cfs["warnings"]:
            report["warnings"].append(dict(warning,
                                           sheet=record["sheet"]))
        report["blocks"].append(block_entry)
        report["rows_added_total"] += record["rows_to_add"]
    if result is not None:
        result.diagnose(f"expansion applied: {report['rows_added_total']} "
                        f"row(s) across {len(report['blocks'])} block(s)")
        for warning in report["warnings"]:
            extra = ({"sheet": warning["sheet"]} if warning.get("sheet")
                     else {})
            result.warn(warning.get("code", "EXPANSION_WARNING"),
                        warning.get("message", ""), **extra)
    return report


def build_expansion_state(fill_plan, profile, source, wb, *, authorized=False, result=None) -> dict:
    """Read-only structural gate (pre-expansion).
    
    Evaluates the plan's expansion evidence against the LIVE workbook
    and returns the authoritative state:
      - present: the plan carries expansion evidence (expansion.blocks)
      - needed:  live capacity math says rows are required
      - rows_to_add: how many rows per block
      - authorized: policy allow_expansion == True (caller sets)
      - blockers: any hard blockers that would stop execution
    
    This is the single source of truth for expansion decisions (D1).
    The dry-run flag (row_expansion.needed) is advisory only; live math wins.
    """
    plan = plan_expansion(fill_plan, profile, source, wb, result=result)
    state = {
        "present": bool(fill_plan.get("expansion")),
        "needed": False,
        "rows_to_add": 0,
        "authorized": authorized,
        "blockers": [],
        "plan": plan,
    }
    if not state["present"]:
        return state
    if not plan.get("ok"):
        state["blockers"].extend(plan.get("blockers", []))
        return state
    total_added = plan.get("rows_added_total", 0)
    state["needed"] = total_added > 0
    state["rows_to_add"] = total_added
    if state["needed"]:
        # Summarize per-block for preflight
        state["block_summary"] = [
            {"sheet": b.get("sheet"),
             "block_id": b.get("block_id"),
             "rows_to_add": b.get("rows_to_add"),
             "insert_at": b.get("insert_at")}
            for b in plan.get("blocks", []) if b.get("rows_to_add", 0) > 0
        ]
    return state



# ---------------------------------------------------------------------------
# FAZ 3B-4: limited exact-match lookup execution
# ---------------------------------------------------------------------------

def _load_source_table(source, table_spec):
    """Load a lookup table from the source data.
    
    table_spec: {"kind": "source_json_tables" | "target_range",
                 "name_or_range": "...", "key_column": "..."}
    """
    kind = table_spec.get("kind")
    if kind == "source_json_tables":
        # The source JSON may have a "tables" key with named tables
        tables = source.get("tables") or {}
        name = table_spec.get("name_or_range")
        if name not in tables:
            raise common.XlsxError(
                "LOOKUP_TABLE_NOT_FOUND",
                f"source table {name!r} not found in source.tables",
                context={"available": list(tables.keys())})
        table = tables[name]
        headers = table.get("headers") or []
        rows = table.get("rows") or []
        if not headers or not rows:
            raise common.XlsxError(
                "LOOKUP_TABLE_EMPTY",
                f"source table {name!r} has no headers or rows")
        key_col = table_spec.get("key_column")
        if key_col not in headers:
            raise common.XlsxError(
                "LOOKUP_KEY_COLUMN_MISSING",
                f"key column {key_col!r} not in table headers",
                context={"headers": headers})
        key_idx = headers.index(key_col)
        # Build dict: key -> row dict
        lookup = {}
        for row in rows:
            key = row[key_idx] if key_idx < len(row) else None
            if key is not None:
                lookup[str(key)] = {headers[i]: row[i] for i in range(len(headers))}
        return {"headers": headers, "lookup": lookup, "key_column": key_col}
    elif kind == "target_range":
        # Not implemented in 3B: would require reading from the target workbook
        raise common.XlsxError(
            "UNSUPPORTED_LOOKUP",
            "target_range lookup tables are not supported in FAZ 3B "
            "(requires cross-workbook read at execution time)",
            context={"kind": kind})
    else:
        raise common.XlsxError(
            "UNSUPPORTED_LOOKUP",
            f"unknown lookup table kind: {kind!r}",
            context={"kind": kind})


def _load_dv_range(ws, dv_spec):
    """Load allowed values from a data-validation list range.
    
    dv_spec: {"sqref": "F2:F10"} or {"table": "TableName", "column": "..."}
    """
    if "sqref" in dv_spec:
        sqref = dv_spec["sqref"]
        # Parse the sqref (single range)
        bounds = mapping._bounds_of_ref(sqref)
        if not bounds:
            raise common.XlsxError(
                "INVALID_DV_RANGE",
                f"could not parse DV sqref: {sqref!r}")
        values = set()
        for row in range(bounds[1], bounds[3] + 1):
            for col in range(bounds[0], bounds[2] + 1):
                cell = ws.cell(row=row, column=col)
                if cell.value is not None:
                    values.add(str(cell.value))
        return values
    else:
        raise common.XlsxError(
            "UNSUPPORTED_LOOKUP",
            "DV range lookup only supports explicit sqref in FAZ 3B")


def execute_lookup(fill_plan, profile, source, loaded, result=None) -> dict:
    """Execute all lookup operations defined in the fill plan.
    
    Lookup is only performed when the plan has execution.lookup.mode == "execute"
    (set by mapping when policy.allow_lookup is True).
    
    Returns a dict with results and any warnings/blockers.
    """
    common.fault("lookup", "at the start of lookup execution")
    execution = fill_plan.get("execution") or {}
    lookup_cfg = execution.get("lookup") or {}
    mode = lookup_cfg.get("mode")
    if mode != "execute":
        return {"executed": False, "reason": "lookup not authorized",
                "results": [], "blockers": [], "warnings": []}
    
    table_spec = lookup_cfg.get("table") or {}
    key_spec = lookup_cfg.get("key") or {}
    result_col = lookup_cfg.get("result_column")
    on_missing = lookup_cfg.get("on_missing", "LOOKUP_NOT_FOUND")
    on_duplicate = lookup_cfg.get("on_duplicate", "LOOKUP_AMBIGUOUS")
    result_kind = lookup_cfg.get("result_kind", "literal")
    allow_formula_overwrite = lookup_cfg.get("allow_formula_overwrite", False)
    
    if not table_spec or not key_spec or not result_col:
        return {"executed": False, "reason": "incomplete lookup config",
                "blockers": [{"code": "PLAN_INVALID",
                              "message": "execution.lookup missing table/key/result_column"}],
                "warnings": []}
    
    wb = loaded.wb
    blockers = []
    warnings = []
    results = []
    
    # Load the lookup table
    try:
        table_data = _load_source_table(source, table_spec)
    except common.XlsxError as e:
        return {"executed": False, "reason": "table load failed",
                "blockers": [e.as_dict()], "warnings": []}
    
    # Determine which cells need lookup (would_write entries with write_policy="lookup")
    # For 3B, we match by finding entries that target the result_column in the block
    key_from = key_spec.get("from")  # "source"
    key_column = key_spec.get("column")  # source column name
    
    if key_from != "source" or not key_column:
        blockers.append({"code": "PLAN_INVALID",
                         "message": "lookup key must have from=source and column"})
        return {"executed": False, "reason": "invalid key spec",
                "blockers": blockers, "warnings": warnings}
    
    # Find the source key index
    header_names = [str(name) for name in (source.get("headers") or [])]
    if not header_names:
        for column in ((source.get("meta") or {}).get("columns") or []):
            header_names.append(str(column.get("key")) if isinstance(column, dict) else str(column))
    
    if key_column not in header_names:
        blockers.append({"code": "LOOKUP_KEY_COLUMN_MISSING",
                         "message": f"lookup key column {key_column!r} not in source headers",
                         "context": {"headers": header_names}})
        return {"executed": False, "reason": "key column not found",
                "blockers": blockers, "warnings": warnings}
    
    key_idx = header_names.index(key_column)
    result_idx = header_names.index(result_col) if result_col in header_names else None
    if result_idx is None:
        blockers.append({"code": "LOOKUP_RESULT_COLUMN_MISSING",
                         "message": f"result column {result_col!r} not in source headers",
                         "context": {"headers": header_names}})
        return {"executed": False, "reason": "result column not found",
                "blockers": blockers, "warnings": warnings}
    
    # Get the lookup table (key -> full row)
    table_lookup = table_data["lookup"]
    
    # Iterate through source rows and resolve lookups for each
    rows = source.get("rows") or []
    for row_offset, row_values in enumerate(rows):
        key_value = row_values[key_idx] if key_idx < len(row_values) else None
        if key_value is None:
            continue
        key_str = str(key_value)
        
        matches = table_lookup.get(key_str)
        if matches is None:
            if on_missing == "LOOKUP_NOT_FOUND":
                blockers.append({"code": "LOOKUP_NOT_FOUND",
                                 "message": f"no match for key {key_str!r} in lookup table",
                                 "key": key_str})
            elif on_missing == "null":
                # Record null result
                results.append({"key": key_str, "result": None, "status": "not_found"})
            # else "skip" - just don't record
            continue
        
        if isinstance(matches, list) and len(matches) > 1:
            if on_duplicate == "LOOKUP_AMBIGUOUS":
                blockers.append({"code": "LOOKUP_AMBIGUOUS",
                                 "message": f"multiple matches for key {key_str!r} in lookup table",
                                 "key": key_str, "count": len(matches)})
            elif on_duplicate == "first":
                matches = matches[0]
            else:
                matches = matches[0]
        elif isinstance(matches, list):
            matches = matches[0]
        
        # matches is now a single row dict
        result_value = matches.get(result_col)
        if result_value is None:
            result_value = ""
        
        results.append({
            "key": key_str,
            "result": str(result_value) if result_value is not None else "",
            "status": "found"
        })
    
    if blockers and result.get("ok") is not False:
        # Blockers will be raised by caller
        pass
    
    return {"executed": True, "mode": mode, "results": results,
            "blockers": blockers, "warnings": warnings,
            "result_kind": result_kind,
            "allow_formula_overwrite": allow_formula_overwrite}

