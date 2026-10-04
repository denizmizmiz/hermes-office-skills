#!/usr/bin/env python3
# MIT License. Part of the Hermes xlsx skill.
"""FAZ 4 / step 1 -- formula tokenizer, normalizer and signature.

Reuses the ONLY A1 parser in the codebase (`xlsx_restructure.REF_RE` +
`COORD_RE`, masked exactly like `RefRewriter.rewrite()` through
`xlsx_expand._masked`) and adds the spans that parser does not own:
whole-column / whole-row ranges (`R:R`, `3:3`), which large real-world sheets hit by
the thousand.

Locked decisions implemented here:
  D1  module name `xlsx_formula_parse.py`
  D2  deterministic R1C1-string signature (`SIG_VERSION`-stamped, never a hash)
  D8  closed unsupported taxonomy; anything unlisted becomes `UNKNOWN_CONSTRUCT`

Nothing here writes a workbook, shifts a formula or duplicates
`RefRewriter`/`translate_formula`: shifting stays in `xlsx_expand`.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_expand as expand          # noqa: E402
import xlsx_restructure as restructure  # noqa: E402

SIG_VERSION = "4.1"

# --- closed unsupported taxonomy (D8) -------------------------------------
UNSUPPORTED_CODES = (
    "UNSUPPORTED_STRUCTURED_FORMULA",
    "UNSUPPORTED_EXTERNAL_REFERENCE",
    "UNSUPPORTED_DYNAMIC_ARRAY",
    "UNSUPPORTED_FUTURE_FUNCTION",
    "UNSUPPORTED_ARRAY_CONSTANT",
    "UNSUPPORTED_LATE_BOUND",
    "UNSUPPORTED_BROKEN_REFERENCE",
    "UNKNOWN_NAMED_REFERENCE",
    "UNKNOWN_CONSTRUCT",
)
LATE_BOUND_FUNCS = frozenset({"INDIRECT", "OFFSET", "GETPIVOTDATA", "CELL",
                              "INFO", "RTD"})
DYNAMIC_FUNCS = frozenset({"SORT", "SORTBY", "FILTER", "UNIQUE", "SEQUENCE",
                          "RANDARRAY", "XLOOKUP", "XMATCH", "LET", "LAMBDA",
                          "TEXTSPLIT", "TOCOL", "TOROW", "WRAPROWS",
                          "WRAPCOLS", "TAKE", "DROP", "EXPAND",
                          "CHOOSEROWS", "CHOOSECOLS"})
RESERVED_WORDS = frozenset({"TRUE", "FALSE", "NULL"})

# --- spans the shared parser does not own ---------------------------------
_WHOLE_QUALIFIED = r"'(?:[^']|'')+'!|[A-Za-z_][A-Za-z0-9_.]*!"
WHOLE_COL_RE = re.compile(
    r"(?<![\w$.:])(?P<prefix>" + _WHOLE_QUALIFIED + r")?"
    r"(?P<ac0>\$?)(?P<col0>[A-Za-z]{1,3}):(?P<ac1>\$?)(?P<col1>[A-Za-z]{1,3})"
    r"(?![\w(])")
WHOLE_ROW_RE = re.compile(
    r"(?<![\w$.:])(?P<prefix>" + _WHOLE_QUALIFIED + r")?"
    r"(?P<ar0>\$?)(?P<row0>\d{1,7}):(?P<ar1>\$?)(?P<row1>\d{1,7})"
    r"(?![\w(])")
# external workbook ref: [Book1.xlsx]Sheet1! / [1]Report! -- a bare
# structured-ref bracket (Table_5[X]) is NOT external
EXTERNAL_RE = re.compile(
    r"\[[^\]]{1,120}\]" r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!")
QUOTED_SPAN_RE = re.compile(r"'(?:[^']|'')*'")
# structured-ref payload: [Volume] is a table column, never an identifier
BRACKET_RE = re.compile(r"\[[^\]]*\]")
FUNC_RE = re.compile(r"(?<![\w.$])(?P<name>[A-Za-z_][A-Za-z0-9_.]*)(?=\s*\()")
IDENT_RE = re.compile(
    r"(?<![\w.$!'\"])(?P<name>[A-Za-z_][A-Za-z0-9_.]*)(?![\w(!\[])")
_PLAIN_SHEET_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
_PLAIN_CELL_RE = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d{1,7})")


def _col_to_index(letters: str) -> int:
    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - 64)
    return index


def _index_to_col(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def parse_cell(cell: str):
    """'H10' -> (col_index, row_number); None when not a plain A1 cell."""
    if not cell:
        return None
    match = _PLAIN_CELL_RE.fullmatch(str(cell).strip())
    if not match:
        return None
    return _col_to_index(match.group(1)), int(match.group(2))


def _whole_token(kind: str, coord, base) -> str:
    """#COL[..] / #COL18 / #ROW[..] / #ROW3 -- a span, not a rectangle."""
    index = coord["col"] if kind == "whole_col" else coord["row"]
    absolute = coord["abs_col"] if kind == "whole_col" else coord["abs_row"]
    axis = "COL" if kind == "whole_col" else "ROW"
    if absolute:
        return "#%s%d" % (axis, index)
    if base is None:
        return "#%s?" % axis
    delta = index - (base[0] if kind == "whole_col" else base[1])
    return "#%s" % axis if delta == 0 else "#%s[%d]" % (axis, delta)


def _r1c1(row, col, abs_row: bool, abs_col: bool, base) -> str:
    """One R1C1 token; base=None means offsets are unknown (no home cell)."""
    if abs_row:
        row_part = "R" + str(row)
    elif base is None or row is None:
        row_part = "R?"
    else:
        delta = row - base[1]
        row_part = "R" if delta == 0 else "R[%d]" % delta
    if abs_col:
        col_part = "C" + str(col)
    elif base is None or col is None:
        col_part = "C?"
    else:
        delta = col - base[0]
        col_part = "C" if delta == 0 else "C[%d]" % delta
    return row_part + col_part


def _sheet_token(prefix: str):
    """Return (prefix_text, sheet_name) for a REF_RE sheet prefix."""
    if not prefix:
        return "", None
    name = prefix[:-1]
    if name.startswith("'"):
        name = name[1:-1].replace("''", "'")
    if _PLAIN_SHEET_RE.match(name):
        text = name.upper() + "!"
    else:
        text = "'" + name.replace("'", "''") + "'!"
    return text, name


def _offset_view(coord) -> dict:
    """Serialisable offsets of one coordinate (rows/cols + absolute flags)."""
    return {"row": coord["row"], "col": coord["col"],
            "abs_row": bool(coord["abs_row"]), "abs_col": bool(coord["abs_col"])}


def _coord(raw: str):
    """COORD_RE parse of one A1 coordinate -> dict or None."""
    match = restructure.COORD_RE.match(raw or "")
    if not match:
        return None
    return {"col": _col_to_index(match.group(2)), "row": int(match.group(4)),
            "abs_col": match.group(1) == "$", "abs_row": match.group(3) == "$"}


def scan_spans(text: str):
    """Reference spans in document order, shared parser first.

    Each span: {kind, start, end, raw, sheet, sheet_prefix, coord, coord_end}.
    `kind` is 'ref' (A1, from REF_RE) or 'whole_col'/'whole_row' (the FAZ 4
    addition). Overlapping matches are dropped so every character is
    substituted exactly once.
    """
    masked = expand._masked(text)
    spans = []
    for match in restructure.REF_RE.finditer(masked):
        coord = _coord(match.group("start"))
        if coord is None:                       # defensive: never guess
            continue
        coord_end = _coord(match.group("end")) if match.group("end") else None
        prefix_text, sheet_name = _sheet_token(match.group("sheet") or "")
        spans.append({"kind": "ref", "start": match.start(),
                      "end": match.end(), "raw": match.group(0),
                      "sheet": sheet_name, "sheet_prefix": prefix_text,
                      "coord": coord, "coord_end": coord_end})
    for kind, pattern, key0, key1 in (
            ("whole_col", WHOLE_COL_RE, "col0", "col1"),
            ("whole_row", WHOLE_ROW_RE, "row0", "row1")):
        for match in pattern.finditer(masked):
            if any(s["start"] < match.end() and match.start() < s["end"]
                   for s in spans):
                continue
            if kind == "whole_col":
                coord = {"col": _col_to_index(match.group(key0)), "row": None,
                         "abs_col": bool(match.group("ac0")), "abs_row": False}
                coord_end = {"col": _col_to_index(match.group(key1)),
                             "row": None, "abs_col": bool(match.group("ac1")),
                             "abs_row": False}
            else:
                coord = {"col": None, "row": int(match.group(key0)),
                         "abs_col": False, "abs_row": bool(match.group("ar0"))}
                coord_end = {"col": None, "row": int(match.group(key1)),
                             "abs_col": False, "abs_row": bool(match.group("ar1"))}
            prefix_text, sheet_name = _sheet_token(match.group("prefix") or "")
            spans.append({"kind": kind, "start": match.start(),
                          "end": match.end(), "raw": match.group(0),
                          "sheet": sheet_name, "sheet_prefix": prefix_text,
                          "coord": coord, "coord_end": coord_end})
    spans.sort(key=lambda s: (s["start"], -s["end"]))
    return masked, spans


def function_tokens(text: str) -> list[str]:
    """Function names outside literals, quoted sheet names and [brackets]."""
    masked = expand._masked(text)
    chars = list(masked)
    for pattern in (QUOTED_SPAN_RE, EXTERNAL_RE):
        for span in pattern.finditer(masked):
            for index in range(span.start(), span.end()):
                chars[index] = " "
    return sorted({m.group("name").upper()
                   for m in FUNC_RE.finditer("".join(chars))})


def unsupported_codes(text: str, functions) -> list[str]:
    """Closed-taxonomy verdict for one formula (D8); never invents a code."""
    codes = []
    masked = expand._masked(text)
    external = EXTERNAL_RE.search(masked)
    if external:
        codes.append("UNSUPPORTED_EXTERNAL_REFERENCE")
    rest = masked
    if external:                                  # brackets already explained
        rest = masked[:external.start()] + " " * (external.end()
                                                  - external.start()) \
            + masked[external.end():]
    if "[" in rest or "]" in rest:
        codes.append("UNSUPPORTED_STRUCTURED_FORMULA")
    if "{" in masked or "}" in masked:
        codes.append("UNSUPPORTED_ARRAY_CONSTANT")
    if "_xlfn." in masked.lower():
        codes.append("UNSUPPORTED_FUTURE_FUNCTION")
    if "#REF!" in masked.upper():
        codes.append("UNSUPPORTED_BROKEN_REFERENCE")
    if set(functions) & LATE_BOUND_FUNCS:
        codes.append("UNSUPPORTED_LATE_BOUND")
    if set(functions) & DYNAMIC_FUNCS:
        codes.append("UNSUPPORTED_DYNAMIC_ARRAY")
    known = [c for c in codes if c in UNSUPPORTED_CODES]
    if len(known) != len(set(codes)):            # defensive: fail closed
        return sorted(set(known) | {"UNKNOWN_CONSTRUCT"})
    return sorted(set(known))


def named_references(text: str, functions=None) -> list[str]:
    """Identifiers that are genuinely named refs (not refs/sheets/tables)."""
    functions = set(functions if functions is not None
                    else function_tokens(text))
    masked = expand._masked(text)
    chars = list(masked)
    for pattern in (restructure.REF_RE, QUOTED_SPAN_RE, WHOLE_COL_RE,
                    WHOLE_ROW_RE, EXTERNAL_RE, BRACKET_RE):
        for span in pattern.finditer(masked):
            for index in range(span.start(), span.end()):
                chars[index] = " "
    out = []
    for match in IDENT_RE.finditer("".join(chars)):
        name = match.group("name")
        end = match.end()
        if end < len(text) and text[end] in "[!(":
            continue
        if name.upper() in functions or name.upper() in RESERVED_WORDS:
            continue
        if restructure.COORD_RE.match(name):
            continue
        out.append(name)
    return sorted(set(out))


def _upper_outside_literals(text: str) -> str:
    """Upper-case everything except the bytes inside "..." literals."""
    masked = expand._masked(text)
    return "".join(char if masked[index] == " " and char != " " else char.upper()
                   for index, char in enumerate(text))


def parse_formula(text, *, sheet=None, cell=None) -> dict:
    """Tokenize + normalize one formula cell into a deterministic dict.

    `sheet` is the formula's own sheet (cross-sheet detection); `cell` is its
    A1 address (relative offsets need it). Same input -> byte-identical output;
    no hashing, no clock, no dict-order dependence.
    """
    if text is None:
        return {"is_formula": False}
    raw = str(text).strip()
    if not raw.startswith("="):
        return {"is_formula": False, "text": raw[:400]}
    base = parse_cell(cell) if cell else None
    functions = function_tokens(raw)
    codes = unsupported_codes(raw, functions)
    named = named_references(raw, functions)
    if named and not codes:
        codes.append("UNKNOWN_NAMED_REFERENCE")
    masked, spans = scan_spans(raw)
    pieces, refs = [], []
    cross_sheets, shapes = [], []
    absolute = relative = mixed = ranges = 0
    cursor = 0
    for span in spans:
        pieces.append(raw[cursor:span["start"]])
        prefix = span["sheet_prefix"] or ""
        if span["sheet"] and (sheet is None or span["sheet"] != sheet):
            cross_sheets.append(span["sheet"])
        start = span["coord"]
        end = span["coord_end"]
        if span["kind"] != "ref":
            pieces.append(prefix + _whole_token(span["kind"], start, base))
            cursor = span["end"]
            ranges += 1
            shapes.append("whole_column" if span["kind"] == "whole_col"
                          else "whole_row")
            if span["kind"] == "whole_col":
                abs_col, abs_row = start["abs_col"], False
            else:
                abs_col, abs_row = False, start["abs_row"]
            if abs_col and abs_row:
                absolute += 1
            elif abs_col or abs_row:
                mixed += 1
            else:
                relative += 1
            refs.append({"kind": span["kind"], "raw": span["raw"],
                         "token": _whole_token(span["kind"], start, base),
                         "sheet": span["sheet"],
                         "offsets": {"start": _offset_view(start),
                                     "end": _offset_view(end)}})
            continue
        start_token = _r1c1(start["row"], start["col"], start["abs_row"],
                           start["abs_col"], base)
        if end is None:
            token = start_token
            shapes.append("single_cell")
            abs_col, abs_row = start["abs_col"], start["abs_row"]
        else:
            end_token = _r1c1(end["row"], end["col"], end["abs_row"],
                             end["abs_col"], base)
            token = start_token + ":" + end_token
            ranges += 1
            if start["row"] == end["row"] and start["col"] != end["col"]:
                shapes.append("horizontal_range")
                abs_col = start["abs_col"] and end["abs_col"]
                abs_row = start["abs_row"] and end["abs_row"]
            elif start["col"] == end["col"] and start["row"] != end["row"]:
                shapes.append("vertical_range")
                abs_col = start["abs_col"] and end["abs_col"]
                abs_row = start["abs_row"] and end["abs_row"]
            else:
                shapes.append("two_dimensional")
                abs_col = start["abs_col"] and end["abs_col"]
                abs_row = start["abs_row"] and end["abs_row"]
        pieces.append(prefix + token)
        cursor = span["end"]
        if abs_col and abs_row:
            absolute += 1
        elif abs_col or abs_row:
            mixed += 1
        else:
            relative += 1
        refs.append({"kind": span["kind"], "raw": span["raw"],
                     "token": token, "sheet": span["sheet"],
                     "offsets": {"start": _offset_view(start),
                                 "end": _offset_view(end) if end else None}})
    pieces.append(raw[cursor:])
    signature = _upper_outside_literals("".join(pieces))
    # operator view: literals and references blanked, so "RC[-2]"/"#COL[3]"
    # tokens can never masquerade as arithmetic operators (measured bug)
    operator_chars = list(masked)
    for span in spans:
        for index in range(span["start"], span["end"]):
            operator_chars[index] = " "
    op_text = "".join(operator_chars)
    return {
        "is_formula": True,
        "text": raw[:400],
        "sheet": sheet,
        "cell": cell,
        "sig_version": SIG_VERSION,
        "signature": signature,
        "functions": functions,
        "named_refs": named,
        "refs": refs,
        "op_text": op_text,
        "refs_total": len(refs),
        "absolute_refs": absolute,
        "relative_refs": relative,
        "mixed_refs": mixed,
        "ranges": ranges,
        "cross_sheet": sorted(set(cross_sheets)),
        "range_shapes": shapes,
        "unsupported": codes,
        "synthesis_eligible": not codes,
        "base_cell_known": base is not None,
    }


def main(argv=None) -> int:
    import argparse
    import json
    parser = argparse.ArgumentParser(description="tokenize/normalize formulas")
    parser.add_argument("formulas", nargs="+")
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--cell", default=None)
    args = parser.parse_args(argv)
    for formula in args.formulas:
        print(json.dumps(parse_formula(formula, sheet=args.sheet,
                                       cell=args.cell),
                         ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    sys.exit(main())
