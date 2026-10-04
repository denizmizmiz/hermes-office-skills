#!/usr/bin/env python3
# MIT License. Part of the Hermes xlsx skill.
"""FAZ 4 / step 2 -- formula family discovery, confidence, outliers.

A *family* is a set of formulas with the same normalized signature
(`xlsx_formula_parse.parse_formula`) sitting in one contiguous line of one
sheet. Nothing is inferred from a single formula; every verdict carries an
`evidence[]` list and a `HIGH/MEDIUM/LOW` confidence -- never a probability.

Locked decisions honoured here:
  D2  identity comes from the parse module's deterministic R1C1 signature;
  D3  synthesis later accepts only HIGH confidence families (this module just
      reports the level);
  D5  intent classes are NOT defined here yet (step 4, `INTENT_CLASSES`);
  D8  unsupported members are reported, never silently normalised away.

Deterministic: entries are sorted before processing and family ids are
assigned after a final deterministic sort, so the same workbook always yields
the same report (AC-4.22/AC-4.23 groundwork).
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_expand as expand        # noqa: E402
import xlsx_formula_parse as fp  # noqa: E402

FAMILY_VERSION = "4.2"
CONF_HIGH, CONF_MEDIUM, CONF_LOW = "HIGH", "MEDIUM", "LOW"
GAP_MERGE_MAX = 2          # an interruption this small keeps one family
_CELL_RE = re.compile(r"^([A-Za-z]{1,3})(\d{1,7})$")


def _cell_parts(cell: str):
    match = _CELL_RE.match(str(cell).strip().lstrip("$"))
    if not match:
        raise ValueError("not an A1 cell: %r" % (cell,))
    return fp._col_to_index(match.group(1)), int(match.group(2))


def _col_letter(index: int) -> str:
    return fp._index_to_col(index)


def build_entry(sheet: str, cell: str, text: str) -> dict:
    """One workbook cell -> parse result + provenance."""
    result = fp.parse_formula(text, sheet=sheet, cell=cell)
    return {"sheet": sheet, "cell": cell, "result": result}


def build_entries(sheet_cells: dict, sheet: str) -> list[dict]:
    return [build_entry(sheet, cell, text)
            for cell, text in sheet_cells.items()]


def _sort_key(entry):
    col, row = _cell_parts(entry["cell"])
    return (entry["sheet"], col, row)


def _signature_of(entry):
    return entry["result"].get("signature")


def _contiguous_runs(cells_indexed, gap: int = 0):
    """cells_indexed: [(index, cell)] sorted by index -> runs of gap<=`gap`."""
    runs = []
    current = []
    for index, cell in cells_indexed:
        if not current:
            current.append((index, cell))
            continue
        previous = current[-1][0]
        if index - previous <= 1 + gap:
            current.append((index, cell))
            continue
        runs.append(current)
        current = [(index, cell)]
    if current:
        runs.append(current)
    return runs




def _dominant_signature(entries):
    """Most frequent signature; ties break on the lexicographically smallest."""
    counts = {}
    for entry in entries:
        signature = _signature_of(entry)
        counts[signature] = counts.get(signature, 0) + 1
    ranked = sorted(counts.items(), key=lambda kv: (-kv[1], str(kv[0])))
    if not ranked:
        return None, counts
    return ranked[0], counts


def _confidence(member_count, member_gaps, run_gaps):
    """HIGH/MEDIUM/LOW from counted evidence -- never a probability."""
    if member_count >= 3 and member_gaps == 0 and run_gaps == 0:
        return CONF_HIGH
    if member_count == 2 and member_gaps == 0 and run_gaps == 0:
        return CONF_MEDIUM
    if member_count >= 3:              # interrupted line or member gaps
        return CONF_MEDIUM
    return CONF_LOW


# --- window classification (D4) -------------------------------------------
# fixed / expanding / rolling / single_row / single_cell / whole_column /
# whole_row / cross_sheet_fixed / cross_sheet_relative / UNKNOWN_WINDOW.
# A verdict needs >= 3 instances on a regularly spaced line AND a unanimous
# start/end behaviour; anything else stays UNKNOWN_WINDOW -- a single formula
# is never classified (D4), and no ambiguity is forced into fixed or rolling.
WINDOW_TYPES = ("fixed", "expanding", "rolling", "single_row", "single_cell",
                "whole_column", "whole_row", "cross_sheet_fixed",
                "cross_sheet_relative", "UNKNOWN_WINDOW")
MIN_WINDOW_INSTANCES = 3


def _axis_value(cell: str, orientation: str) -> int:
    col, row = _cell_parts(cell)
    return row if orientation == "vertical" else col


def _tracks(coord: dict, orientation: str) -> bool:
    """True when this endpoint moves with the formula (relative axis)."""
    return not (coord["abs_row"] if orientation == "vertical" else coord["abs_col"])


def _spacing_of(member_cells, orientation):
    values = sorted(_axis_value(cell, orientation) for cell in member_cells)
    diffs = [b - a for a, b in zip(values, values[1:])]
    regular = bool(diffs) and all(diff == diffs[0] and diff > 0 for diff in diffs)
    return {"step": diffs[0] if regular else None, "regular": regular,
            "spacing_diffs": sorted(set(diffs))}


def _unknown_window(reason, instances, spacing=None):
    return {"window_type": "UNKNOWN_WINDOW", "instances": instances,
            "spacing": spacing, "evidence": [], "ambiguity": [reason]}


def classify_window(member_entries, orientation):
    """Window verdict from counted evidence; UNKNOWN_WINDOW is a real result."""
    cells = [entry["cell"] for entry in member_entries]
    spacing = _spacing_of(cells, orientation)
    instances = len(cells)
    if instances < MIN_WINDOW_INSTANCES:
        return _unknown_window(
            "insufficient instances for a window verdict (%d < %d)"
            % (instances, MIN_WINDOW_INSTANCES), instances, spacing)
    if not spacing["regular"]:
        return _unknown_window(
            "irregular instance spacing %s; start/end behaviour cannot be "
            "measured reliably" % spacing["spacing_diffs"], instances, spacing)
    behaviours, cross_sheet = set(), False
    for entry in member_entries:
        result = entry["result"]
        range_refs = [ref for ref in (result.get("refs") or [])
                      if ref["kind"] != "ref"
                      or (ref.get("offsets") or {}).get("end")]
        home = result.get("sheet")
        for ref in range_refs:
            if ref.get("sheet") and ref["sheet"] != home:
                cross_sheet = True
            if ref["kind"] == "whole_col":
                behaviours.add("whole_column")
                continue
            if ref["kind"] == "whole_row":
                behaviours.add("whole_row")
                continue
            start, end = ref["offsets"]["start"], ref["offsets"]["end"]
            if start["row"] == end["row"] and start["col"] != end["col"]:
                behaviours.add("single_row")
                continue
            start_tracks = _tracks(start, orientation)
            end_tracks = _tracks(end, orientation)
            if not start_tracks and not end_tracks:
                behaviours.add("fixed")
            elif start_tracks and end_tracks:
                behaviours.add("rolling")
            elif not start_tracks and end_tracks:
                behaviours.add("expanding")
            else:
                behaviours.add("contracting")
        if not range_refs:
            behaviours.add("single_cell")
    evidence = [
        "%d instance(s) on a regular line (step %s)" % (instances,
                                                        spacing["step"]),
        "start/end behaviour: %s" % ", ".join(sorted(behaviours)),
    ]
    if "contracting" in behaviours:
        return {"window_type": "UNKNOWN_WINDOW", "instances": instances,
                "spacing": spacing, "evidence": evidence,
                "ambiguity": ["start endpoint moves while end endpoint is "
                              "anchored (contracting window is not a "
                              "classified type)"]}
    if len(behaviours) != 1:
        return {"window_type": "UNKNOWN_WINDOW", "instances": instances,
                "spacing": spacing, "evidence": evidence,
                "ambiguity": ["mixed window behaviours in the same family: %s"
                              % ", ".join(sorted(behaviours))]}

    behaviour = next(iter(behaviours))
    if behaviour == "single_cell":
        window_type = "single_cell"
    elif behaviour in ("whole_column", "whole_row"):
        window_type = behaviour
    elif behaviour == "single_row":
        window_type = "single_row"
    else:
        window_type = behaviour                  # fixed / expanding / rolling
    if cross_sheet:
        window_type = ("cross_sheet_fixed"
                       if behaviour in ("fixed", "single_row")
                       else "cross_sheet_relative")
        evidence.append("cross-sheet window; classified as %s" % window_type)
    return {"window_type": window_type, "instances": instances,
            "spacing": spacing, "evidence": evidence, "ambiguity": [],
            "raw_behaviour": behaviour, "cross_sheet": cross_sheet}


# --- intent inference (D5, spec sections 18-19) ---------------------------
# The tuple below IS the taxonomy: documents never restate a count, they
# reference this constant (drift test: tests/test_xlsx_formula_intent.py).
# `unknown` is a real result and is never silently coerced into another
# class; function names alone are never sufficient evidence (spec section 18).
INTENT_CLASSES = ("calculation", "aggregation", "lookup", "conditional",
                  "date_logic", "validation_logic", "reference", "summary",
                  "ratio", "percentage", "unknown")
INTENT_VERSION = "4.4"

_AGG_FUNCS = frozenset({"SUM", "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS",
                        "COUNT", "COUNTA", "MAX", "MIN", "AVERAGE",
                        "AVERAGEIF", "AVERAGEIFS"})
_LOOKUP_FUNCS = frozenset({"VLOOKUP", "HLOOKUP", "XLOOKUP", "LOOKUP",
                           "MATCH", "INDEX"})
_DATE_FUNCS = frozenset({"TODAY", "NOW", "DATE", "EDATE", "EOMONTH", "YEAR",
                         "MONTH", "DAY", "WORKDAY", "NETWORKDAYS"})
_VALID_FUNCS = frozenset({"ISERROR", "ISNUMBER", "ISBLANK", "ISTEXT",
                          "ISLOGICAL", "EXACT", "LEN"})
_COND_FUNCS = frozenset({"IF", "IFS", "SWITCH", "AND", "OR", "NOT", "IFNA"})
_ERROR_WRAPPERS = frozenset({"IFERROR", "IFNA"})
_COMPARE_RE = re.compile(r"(<=|>=|<>|[=<>])")
_ARITH = frozenset("+-*/^")
_FUNC_SCAN_RE = re.compile(r"([A-Z_][A-Z0-9_.]*)\s*\(")
_BASE_CONFIDENCE = {"conditional": CONF_HIGH, "aggregation": CONF_HIGH,
                    "lookup": CONF_HIGH, "date_logic": CONF_HIGH,
                    "calculation": CONF_HIGH, "reference": CONF_HIGH,
                    "summary": CONF_MEDIUM, "validation_logic": CONF_MEDIUM,
                    "percentage": CONF_MEDIUM, "ratio": CONF_MEDIUM,
                    "unknown": CONF_LOW}
_DOWNGRADE = {CONF_HIGH: CONF_MEDIUM, CONF_MEDIUM: CONF_LOW,
              CONF_LOW: CONF_LOW}


def _masked_signature(result) -> str:
    return expand._masked(result.get("signature") or "")


def _outermost_function(result):
    """First function token in document order (the one that governs)."""
    masked = _masked_signature(result)
    functions = set(result.get("functions") or [])
    for match in _FUNC_SCAN_RE.finditer(masked):
        name = match.group(1)
        if name in functions:
            return name
    return None


def _summary_evidence(result):
    """Aggregation over the same column ending directly above the formula."""
    cell = fp.parse_cell(result.get("cell") or "")
    if cell is None:
        return False
    col, row = cell
    for ref in result.get("refs") or []:
        offsets = ref.get("offsets") or {}
        start, end = offsets.get("start"), offsets.get("end")
        if not start or not end:
            continue
        if (ref["kind"] == "ref" and start["col"] == col
                and end["col"] == col and end["row"] == row - 1):
            return True
    return False


def _literal_count(result) -> int:
    """Numeric literals outside references/function names (evidence: `=1+1`
    is arithmetic over CONSTANT operands just like over references)."""
    sig = result.get("signature") or ""
    for ref in result.get("refs") or []:
        sig = sig.replace(ref.get("token") or "", " ")
    for name in result.get("functions") or []:
        sig = re.sub(r"\b%s\s*\(" % re.escape(name), "(", sig, flags=re.I)
    return len(re.findall(r"(?<![A-Za-z_$])\d+(?:\.\d+)?", sig))


def classify_intent(result) -> dict:
    """Evidence-based intent; `unknown` when no rule matches (never forced)."""
    if not result or not result.get("is_formula"):
        return {"intent": "unknown", "confidence": CONF_LOW,
                "intent_version": INTENT_VERSION, "evidence": [],
                "ambiguity": ["not a formula"]}
    functions = set(result.get("functions") or [])
    body = (result.get("op_text") or "").strip()
    if body.startswith("="):               # the formula prefix is not a test
        body = body[1:]
    has_compare = bool(_COMPARE_RE.search(body))
    has_arith = any(char in _ARITH for char in body)
    has_div = "/" in body
    times100 = "*100" in body or "%" in body
    outermost = _outermost_function(result)
    refs_total = result.get("refs_total") or 0
    has_range = any(ref["kind"] != "ref" or (ref.get("offsets") or {}).get("end")
                    for ref in result.get("refs") or [])
    evidence, ambiguity = [], []
    decision = None

    def consider(name, condition, note):
        if condition:
            evidence.append(note)
            return name
        return None

    if outermost in _ERROR_WRAPPERS:
        evidence.append("outermost %s is an error-handling wrapper; the "
                        "inner expression governs the intent" % outermost)
    elif outermost in _COND_FUNCS:
        if has_compare:
            decision = consider("conditional", True,
                                "outermost %s with a comparison operator"
                                % outermost)
        else:
            ambiguity.append("outermost %s without a comparison operator"
                             % outermost)
    if decision is None and outermost in _AGG_FUNCS:
        if _summary_evidence(result):
            decision = "summary"
            evidence.append("outermost %s over the same column, range ends "
                            "directly above the formula (summary row)"
                            % outermost)
        else:
            decision = "aggregation"
            evidence.append("outermost %s over a %s reference"
                            % (outermost, "range" if has_range else "point"))
    if decision is None and (outermost in _LOOKUP_FUNCS
                             or {"INDEX", "MATCH"} <= functions):
        decision = "lookup"
        evidence.append("lookup chain: %s"
                        % ", ".join(sorted(functions & _LOOKUP_FUNCS)))
    if decision is None and functions & _DATE_FUNCS:
        decision = "date_logic"
        evidence.append("date function(s): %s"
                        % ", ".join(sorted(functions & _DATE_FUNCS)))
    if decision is None and functions & _VALID_FUNCS and (
            has_compare or outermost in _VALID_FUNCS):
        decision = "validation_logic"
        evidence.append("validation function(s) %s%s"
                        % (", ".join(sorted(functions & _VALID_FUNCS)),
                           " with a comparison" if has_compare
                           else " as the outermost function"))
    if decision is None and has_div:
        if times100:
            decision = "percentage"
            evidence.append("division combined with *100 or a percent sign")
        else:
            decision = "ratio"
            evidence.append("division between references without a percent "
                            "pattern")
        evidence.append("number format unavailable at parse time: classified "
                        "from the operator pattern only")
    if decision is None and has_arith and (refs_total + _literal_count(
            result)) >= 2:
        decision = "calculation"
        evidence.append("arithmetic over %d reference(s) and %d constant "
                        "operand(s), no aggregation function"
                        % (refs_total, _literal_count(result)))
    if decision is None and not functions and refs_total and not (
            has_arith or has_compare):
        decision = "reference"
        evidence.append("reference-only formula (no function, no operator)")
    if decision is None:
        decision = "unknown"
        ambiguity.append("no intent rule matched the available evidence")
    # record competing candidates instead of hiding them
    for name, matches, note in (
            ("conditional", has_compare and bool(functions & _COND_FUNCS),
             "contains a conditional function with a comparison"),
            ("aggregation", bool(functions & _AGG_FUNCS),
             "contains an aggregation function"),
            ("lookup", bool(functions & _LOOKUP_FUNCS),
             "contains a lookup function"),
            ("date_logic", bool(functions & _DATE_FUNCS),
             "contains a date function"),
            ("validation_logic", bool(functions & _VALID_FUNCS),
             "contains a validation function"),
            ("percentage", has_div and times100, "division with *100"),
            ("ratio", has_div, "division"),
            ("calculation", has_arith, "arithmetic operator present"),
            ("reference", not functions and bool(refs_total)
             and not has_arith and not has_compare, "reference-only")):
        subsumed = {("summary", "aggregation"), ("percentage", "ratio"),
                    ("conditional", "reference"), ("ratio", "calculation"),
                    ("percentage", "calculation")}
        if matches and name != decision and (decision, name) not in subsumed:
            ambiguity.append("%s evidence also present" % name)
    confidence = _BASE_CONFIDENCE.get(decision, CONF_LOW)
    if ambiguity and any("also present" in item for item in ambiguity):
        confidence = _DOWNGRADE[confidence]
    return {"intent": decision, "confidence": confidence,
            "intent_version": INTENT_VERSION, "evidence": evidence,
            "ambiguity": ambiguity}


def _family_intent(members):
    verdicts = [classify_intent(entry["result"]) for entry in members]
    intents = {verdict["intent"] for verdict in verdicts}
    if len(intents) == 1:
        return verdicts[0]
    return {"intent": "unknown", "confidence": CONF_LOW,
            "intent_version": INTENT_VERSION,
            "evidence": ["family members disagree on intent: %s"
                         % ", ".join(sorted(intents))],
            "ambiguity": ["members disagree; intent left unknown"]}


def _gaps_within(run):
    gaps = 0
    for (previous, _), (current, _) in zip(run, run[1:]):
        if current - previous > 1:
            gaps += 1
    return gaps


def _member_view(entry):
    return {"cell": entry["cell"],
            "signature": _signature_of(entry),
            "unsupported": entry["result"].get("unsupported") or [],
            "synthesis_eligible": bool(
                entry["result"].get("synthesis_eligible"))}


def _build_block(sheet, orientation, line_index, run):
    """One gap-bounded run -> (family, member_entries, outlier_entries)|None.

    A block only yields a family when at least two of its formulas share one
    normalized signature; everything else is either separated as an outlier
    (spec section 20-21) or -- when no signature repeats -- left as singles,
    never silently absorbed into a family.
    """
    entries = [entry for _, entry in run]
    (signature, count), _counts = _dominant_signature(entries)
    if signature is None or count < 2:
        return None
    members = [entry for entry in entries if _signature_of(entry) == signature]
    out = [entry for entry in entries if _signature_of(entry) != signature]
    gaps = _gaps_within(run)
    member_gaps = _gaps_within([pair for pair in run
                                if pair[1] in members])
    contiguous = gaps == 0 and member_gaps == 0
    confidence = _confidence(len(members), member_gaps, gaps)
    first_cell, last_cell = members[0]["cell"], members[-1]["cell"]
    cross = sorted({sheet_name for entry in members
                    for sheet_name in (entry["result"].get("cross_sheet") or [])})
    unsupported_members = [entry for entry in members
                           if not entry["result"].get("synthesis_eligible")]
    evidence = [
        "same normalized signature %s (sig_version %s)"
        % (signature, members[0]["result"].get("sig_version")),
        "%s line continuity: %d member(s) over %s..%s"
        % (orientation, len(members), first_cell, last_cell),
    ]
    if gaps:
        evidence.append("interrupted continuity: %d gap(s) inside the line"
                        % gaps)
    if member_gaps:
        evidence.append("member continuity: %d gap(s) between family members"
                        % member_gaps)
    if out:
        evidence.append("block purity: %d/%d carry the family signature; "
                        "%d separated as outlier(s)"
                        % (count, len(entries), len(out)))
    else:
        evidence.append("block purity: no outliers in block")
    if cross:
        evidence.append("cross-sheet topology identical: %s"
                        % ", ".join(cross))
    if unsupported_members:
        evidence.append("unsupported member(s): %s"
                        % ", ".join(
                            "%s (%s)" % (entry["cell"],
                                         ",".join(entry["result"]["unsupported"]))
                            for entry in unsupported_members))
    member_pairs = [pair for pair in run if pair[1] in members]
    family = {
        "family_id": None,                     # assigned after the final sort
        "sheet": sheet,
        "orientation": orientation,
        "line": line_index,
        "signature": signature,
        "sig_version": members[0]["result"].get("sig_version"),
        "member_count": len(members),
        "members": [_member_view(entry) for entry in members],
        "representative": {"cell": members[0]["cell"]},
        "block": {"start": first_cell, "end": last_cell},
        "contiguous": contiguous,
        "gaps": gaps,
        "confidence": confidence,
        "evidence": evidence,
        "cross_sheet": cross,
        "unsupported_any": bool(unsupported_members),
        "synthesis_eligible": not unsupported_members,
        "synthesis_allowed": (not unsupported_members
                              and confidence == CONF_HIGH),
    }
    family["window"] = classify_window(members, orientation)
    family["intent"] = _family_intent(members)
    return family, members, out


def discover(entries) -> dict:
    """Group parsed formulas into families; deterministic for any input order."""
    formula_entries = [entry for entry in entries
                       if entry["result"].get("is_formula")]
    formula_entries.sort(key=_sort_key)
    per_sheet = {}
    for entry in formula_entries:
        per_sheet.setdefault(entry["sheet"], []).append(entry)

    blocks, claimed = [], set()
    for sheet in sorted(per_sheet):
        sheet_entries = per_sheet[sheet]
        for orientation in ("vertical", "horizontal"):
            lines = {}
            for entry in sheet_entries:
                col, row = _cell_parts(entry["cell"])
                key = col if orientation == "vertical" else row
                lines.setdefault(key, []).append(entry)
            for line_index in sorted(lines):
                indexed = []
                for entry in lines[line_index]:
                    if orientation == "horizontal" and entry["cell"] in claimed:
                        continue
                    col, row = _cell_parts(entry["cell"])
                    index = row if orientation == "vertical" else col
                    indexed.append((index, entry))
                indexed.sort(key=lambda pair: pair[0])
                for run in _contiguous_runs(indexed, gap=GAP_MERGE_MAX):
                    if len(run) < 2:
                        continue
                    built = _build_block(sheet, orientation, line_index, run)
                    if built is None:
                        continue
                    blocks.append(built)
                    for _, entry in run:
                        claimed.add(entry["cell"])

    families = [family for family, _members, _out in blocks]
    families.sort(key=lambda family: (family["sheet"],
                                      family["orientation"], family["line"],
                                      family["block"]["start"]))
    for index, family in enumerate(families, 1):
        family["family_id"] = "F%04d" % index

    outliers = []
    for family, _members, out in blocks:
        for entry in out:
            outliers.append({
                "sheet": family["sheet"], "cell": entry["cell"],
                "signature": _signature_of(entry),
                "outlier_of": family["family_id"],
                "unsupported": entry["result"].get("unsupported") or [],
                "synthesis_eligible": bool(
                    entry["result"].get("synthesis_eligible")),
                "intent": classify_intent(entry["result"]),
                "evidence": ["signature differs from the dominant family "
                             "signature %s" % family["signature"]],
            })

    singles = []
    for entry in formula_entries:
        if entry["cell"] in claimed:
            continue
        singles.append({"sheet": entry["sheet"], "cell": entry["cell"],
                        "signature": _signature_of(entry),
                        "unsupported": entry["result"].get("unsupported") or [],
                        "synthesis_eligible": bool(
                            entry["result"].get("synthesis_eligible")),
                        "window": _unknown_window(
                            "single formula: window classification needs a "
                            "family with >= %d instances" % MIN_WINDOW_INSTANCES,
                            1),
                        "intent": classify_intent(entry["result"]),
                        "evidence": ["no family: no other formula on this "
                                     "line shares its signature"]})
    outliers.sort(key=lambda record: (record["sheet"], record["cell"]))
    singles.sort(key=lambda record: (record["sheet"], record["cell"]))

    covered = sum(family["member_count"] for family in families)
    report = {
        "family_version": FAMILY_VERSION,
        "families": families,
        "outliers": outliers,
        "singles": singles,
        "stats": {
            "formulas": len(formula_entries),
            "family_members": covered,
            "families": len(families),
            "high": len([f for f in families if f["confidence"] == CONF_HIGH]),
            "medium": len([f for f in families
                           if f["confidence"] == CONF_MEDIUM]),
            "low": len([f for f in families if f["confidence"] == CONF_LOW]),
            "outliers": len(outliers),
            "singles": len(singles),
            "coverage_fraction": (round(covered / len(formula_entries), 4)
                                  if formula_entries else 0.0),
        },
    }
    return report


def main(argv=None) -> int:
    import argparse
    import json
    parser = argparse.ArgumentParser(description="formula family discovery")
    parser.add_argument("--cells", help="JSON dump: {sheet: {cell: formula}}")
    parser.add_argument("--out")
    args = parser.parse_args(argv)
    if not args.cells:
        print(json.dumps({"error": "usage: xlsx_formula_family.py --cells "
                                   "cells.json [--out report.json]"}))
        return 2
    cells_by_sheet = json.loads(Path(args.cells).read_text(encoding="utf-8"))
    entries = []
    for sheet, cells in cells_by_sheet.items():
        entries.extend(build_entries(cells, sheet))
    report = discover(entries)
    console = dict(report)
    console["families"] = [{key: family[key] for key in
                            ("family_id", "sheet", "block", "member_count",
                             "confidence", "signature", "synthesis_allowed")}
                           for family in report["families"]]
    if args.out:
        Path(args.out).write_text(json.dumps(report, indent=1,
                                             ensure_ascii=False),
                                  encoding="utf-8")
    print(json.dumps(console, indent=1, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
