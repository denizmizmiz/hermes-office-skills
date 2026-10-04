#!/usr/bin/env python3
# MIT License. Part of the Hermes xlsx skill.
"""FAZ 4 / step 6a -- controlled formula synthesis: PLANNING + VALIDATION.

D7 (locked): synthesis lives in THIS module; `xlsx_execute.py` is not
touched and will only ever apply an approved plan through its existing
3C safe-write contract (step 6b). Nothing here writes a workbook.

Every synthesis action must clear the spec's nine preconditions (section
22); any missing one yields `FORMULA_SYNTHESIS_BLOCKED` with a code, never a
guess. The translated formula is re-parsed and must reproduce the source
family's normalized signature exactly -- a mismatch blocks the action.

No source family -> zero generated formulas (AC-4.11). Unsupported family ->
zero (AC-4.10). Absolute references survive translation (AC-4.13); relative
offsets translate through the existing `translate_formula` machinery only
(AC-4.14) -- no new shifting code.
"""
from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_expand as expand            # noqa: E402
import xlsx_formula_family as family_mod  # noqa: E402
import xlsx_formula_parse as fp         # noqa: E402

SYNTH_VERSION = "4.6a"
WRITE_POLICY = "write_formula"
BLOCKED = "FORMULA_SYNTHESIS_BLOCKED"

BLOCK_CODES = (
    "NO_SOURCE_FAMILY",
    "FAMILY_CONFIDENCE_BELOW_HIGH",
    "FAMILY_UNSUPPORTED_CONSTRUCT",
    "WINDOW_PATTERN_NOT_DETERMINISTIC",
    "TARGET_STATE_UNKNOWN",
    "TARGET_ALREADY_FORMULA",
    "TARGET_NOT_ALLOWED",
    "TARGET_COLUMN_MISMATCH",
    "REFERENCE_TRANSLATION_MISMATCH",
    "SOURCE_TEXT_UNAVAILABLE",
    "TRANSLATION_FAILED",
)


def _cell(cell: str):
    return fp.parse_cell(cell)


def select_source_family(family_report, sheet: str, target_cell: str):
    """Nearest HIGH-confidence, synthesis-eligible family above the target in
    the same sheet+column (deterministic tie-break: member_count, then id)."""
    target = _cell(target_cell)
    if target is None:
        return None, "TARGET_STATE_UNKNOWN"
    col, row = target
    candidates = []
    for fam in family_report.get("families") or []:
        if fam["sheet"] != sheet or fam.get("orientation") != "vertical":
            continue
        if not fam.get("synthesis_allowed"):
            continue
        block_end = _cell(fam["block"]["end"])
        block_start = _cell(fam["block"]["start"])
        if block_end is None or block_start is None:
            continue
        if block_end[0] != col or block_start[0] != col:
            continue
        if block_end[1] >= row:                      # must sit above the target
            continue
        candidates.append((row - block_end[1], -fam["member_count"],
                           fam["family_id"], fam))
    if not candidates:
        return None, "NO_SOURCE_FAMILY"
    candidates.sort()
    return candidates[0][3], None


def _family_source_text(family):
    """The family's representative formula text (must exist verbatim)."""
    for member in family.get("members") or []:
        if member.get("cell") == family["representative"]["cell"]:
            text = (member.get("text")
                    or member.get("formula_text"))
            if text:
                return text
    return None


def plan_synthesis(family_report, targets, *, families_with_text=None,
                   allowed_targets=None, target_states=None) -> dict:
    """Build a deterministic synthesis plan.

    `target_states`: {("SHEET","A1"): {"has_formula": bool, "has_value": bool}}
    -- supplied by a read-only workbook scan; unknown state blocks the action.
    `families_with_text`: {family_id: formula_text} when member views do not
    carry the text (the family report's member views keep cells only).
    """
    families_by_id = {fam["family_id"]: fam
                      for fam in (family_report.get("families") or [])}
    allowed = set(allowed_targets) if allowed_targets is not None else None
    actions, blocked_records = [], []
    for target in targets:
        sheet, cell = target.get("sheet"), target.get("cell")
        record = {"target": {"sheet": sheet, "cell": cell}}
        fam = None
        if target.get("family_id"):
            fam = families_by_id.get(target["family_id"])
            if fam is None:
                record.update({"code": BLOCKED,
                               "reason_code": "NO_SOURCE_FAMILY"})
                blocked_records.append(record)
                continue
        else:
            fam, reason = select_source_family(family_report, sheet, cell)
            if fam is None:
                record.update({"code": BLOCKED, "reason_code": reason})
                blocked_records.append(record)
                continue
        if fam.get("orientation") != "vertical":
            record.update({"code": BLOCKED,
                           "reason_code": "TARGET_COLUMN_MISMATCH"})
            blocked_records.append(record)
            continue
        if fam.get("confidence") != family_mod.CONF_HIGH:
            record.update({"code": BLOCKED, "family_id": fam["family_id"],
                           "reason_code": "FAMILY_CONFIDENCE_BELOW_HIGH"})
            blocked_records.append(record)
            continue
        if not fam.get("synthesis_eligible"):
            record.update({"code": BLOCKED, "family_id": fam["family_id"],
                           "reason_code": "FAMILY_UNSUPPORTED_CONSTRUCT"})
            blocked_records.append(record)
            continue
        window_type = (fam.get("window") or {}).get("window_type")
        if window_type in (None, "UNKNOWN_WINDOW"):
            record.update({"code": BLOCKED, "family_id": fam["family_id"],
                           "reason_code": "WINDOW_PATTERN_NOT_DETERMINISTIC"})
            blocked_records.append(record)
            continue
        if allowed is not None and (sheet, cell) not in allowed:
            record.update({"code": BLOCKED, "family_id": fam["family_id"],
                           "reason_code": "TARGET_NOT_ALLOWED"})
            blocked_records.append(record)
            continue
        state = (target_states or {}).get((sheet, cell))
        if state is None:
            record.update({"code": BLOCKED, "family_id": fam["family_id"],
                           "reason_code": "TARGET_STATE_UNKNOWN"})
            blocked_records.append(record)
            continue
        if state.get("has_formula"):
            record.update({"code": BLOCKED, "family_id": fam["family_id"],
                           "reason_code": "TARGET_ALREADY_FORMULA"})
            blocked_records.append(record)
            continue
        source_text = None
        if families_with_text:
            source_text = families_with_text.get(fam["family_id"])
        if source_text is None:
            source_text = _family_source_text(fam)
        if not source_text:
            record.update({"code": BLOCKED, "family_id": fam["family_id"],
                           "reason_code": "SOURCE_TEXT_UNAVAILABLE"})
            blocked_records.append(record)
            continue
        rep_cell = _cell(fam["representative"]["cell"])
        target_cell = _cell(cell)
        if rep_cell is None or target_cell is None:
            record.update({"code": BLOCKED, "family_id": fam["family_id"],
                           "reason_code": "TARGET_STATE_UNKNOWN"})
            blocked_records.append(record)
            continue
        if rep_cell[0] != target_cell[0]:
            record.update({"code": BLOCKED, "family_id": fam["family_id"],
                           "reason_code": "TARGET_COLUMN_MISMATCH"})
            blocked_records.append(record)
            continue
        drow = target_cell[1] - rep_cell[1]
        try:
            translated = expand.translate_formula(source_text, drow)
        except Exception as exc:                    # pragma: no cover
            record.update({"code": BLOCKED, "family_id": fam["family_id"],
                           "reason_code": "TRANSLATION_FAILED",
                           "detail": str(exc)[:160]})
            blocked_records.append(record)
            continue
        parsed = fp.parse_formula(translated, sheet=sheet, cell=cell)
        if parsed.get("unsupported"):
            record.update({"code": BLOCKED, "family_id": fam["family_id"],
                           "reason_code": "FAMILY_UNSUPPORTED_CONSTRUCT"})
            blocked_records.append(record)
            continue
        if parsed.get("signature") != fam.get("signature"):
            record.update({"code": BLOCKED, "family_id": fam["family_id"],
                           "reason_code": "REFERENCE_TRANSLATION_MISMATCH",
                           "detail": "translated signature %s != family %s"
                                     % (parsed.get("signature"),
                                        fam.get("signature"))})
            blocked_records.append(record)
            continue
        actions.append({
            "target": {"sheet": sheet, "cell": cell},
            "family_id": fam["family_id"],
            "source_cell": fam["representative"]["cell"],
            "formula_text": translated,
            "drow": drow,
            "validation": {
                "signature": parsed["signature"],
                "signature_matches_family": True,
                "unsupported": [],
                "window_type": window_type,
                "confidence": fam["confidence"],
            },
            "evidence": ["source family %s (HIGH, %s)" % (fam["family_id"],
                                                          window_type),
                         "row translation %+d via translate_formula" % drow,
                         "translated signature reproduced the family "
                         "signature exactly"],
        })
    payload = {"synth_version": SYNTH_VERSION,
               "actions": actions, "blocked": blocked_records}
    canonical = json.dumps(payload, sort_keys=True, ensure_ascii=False)
    plan_id = "SP4-" + hashlib.sha1(canonical.encode("utf-8")).hexdigest()[:16]
    plan = dict(payload)
    plan["plan_id"] = plan_id
    plan["write_policy"] = WRITE_POLICY
    plan["generated_formula_count"] = len(actions)
    plan["qa_expectations"] = qa_expectations(actions)
    plan["executor_contract"] = {
        "plan_type": "formula_synthesis",
        "safe_write": "phase-3c",
        "writes_formula_cells": True,
        "note": "applied by xlsx_execute through the existing plan/safe-write "
                "contract only (D7-A); synthesis logic stays in this module",
    }
    return plan


def qa_expectations(actions) -> dict:
    """Pre/post QA expectations for an approved plan (spec section 25)."""
    return {
        "formula_count_delta_expected": len(actions),
        "checks": [
            "formula exists at every target after write",
            "formula syntax preserved (re-parse equals planned text)",
            "expected family signature matched at every target",
            "expected relative offsets matched at every target",
            "absolute references preserved (byte compare vs plan)",
            "sheet references preserved (byte compare vs plan)",
            "no unrelated formula text changed (full inventory compare)",
        ],
        "values_recalculated": False,
        "note": "openpyxl does not evaluate formulas; recalculation is a "
                "separate step (LibreOffice) and is never claimed here",
    }


def main(argv=None) -> int:
    import argparse
    parser = argparse.ArgumentParser(description="plan formula synthesis")
    parser.add_argument("--families", required=True,
                        help="family report JSON (xlsx_formula_family)")
    parser.add_argument("--targets", required=True,
                        help="JSON list: [{sheet, cell, family_id?}]")
    parser.add_argument("--target-states",
                        help="JSON map {'SHEET!A1': {has_formula, has_value}}")
    parser.add_argument("--out")
    args = parser.parse_args(argv)
    family_report = json.loads(Path(args.families).read_text(encoding="utf-8"))
    targets = json.loads(Path(args.targets).read_text(encoding="utf-8"))
    raw_states = {}
    if args.target_states:
        raw = json.loads(Path(args.target_states).read_text(encoding="utf-8"))
        raw_states = {tuple(key.split("!")): value
                      for key, value in raw.items()}
    plan = plan_synthesis(family_report, targets,
                          target_states=raw_states)
    if args.out:
        Path(args.out).write_text(json.dumps(plan, indent=1,
                                             ensure_ascii=False),
                                  encoding="utf-8")
    print(json.dumps({key: plan[key] for key in
                      ("plan_id", "generated_formula_count", "write_policy")},
                     indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
