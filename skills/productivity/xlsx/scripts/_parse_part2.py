

def unsupported_codes(text: str, functions: set[str]) -> list[str]:
    """Closed-taxonomy verdict for one formula (D8)."""
    codes = []
    masked = expand._masked(text)
    if EXTERNAL_RE.search(masked):
        codes.append("UNSUPPORTED_EXTERNAL_REFERENCE")
    if "[" in masked or "]" in masked:
        codes.append("UNSUPPORTED_STRUCTURED_FORMULA")
    if "{" in masked or "}" in masked:
        codes.append("UNSUPPORTED_ARRAY_CONSTANT")
    if "_xlfn." in masked.lower():
        codes.append("UNSUPPORTED_FUTURE_FUNCTION")
    if "#REF!" in masked.upper():
        codes.append("UNSUPPORTED_BROKEN_REFERENCE")
    late = sorted(functions & LATE_BOUND_FUNCS)
    if late:
        codes.append("UNSUPPORTED_LATE_BOUND")
    dynamic = sorted(functions & DYNAMIC_FUNCS)
    if dynamic:
        codes.append("UNSUPPORTED_DYNAMIC_ARRAY")
    unknown = sorted(c for c in codes if c not in UNSUPPORTED_CODES)
    if unknown:                                   # defensive: never open-ended
        codes = [c for c in codes if c in UNSUPPORTED_CODES]
        codes.append("UNKNOWN_CONSTRUCT")
    return sorted(set(codes))


def named_references(text: str) -> list[str]:
    """Identifiers that are genuinely named refs (not refs/sheets/tables)."""
    functions = set(function_tokens(text))
    masked = expand._masked(text)
    chars = list(masked)
    for pattern in (restructure.REF_RE, QUOTED_SPAN_RE, WHOLE_COL_RE,
                    WHOLE_ROW_RE, BARE_WHOLE_COL_RE, BARE_WHOLE_ROW_RE,
                    EXTERNAL_RE):
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


def parse_formula(text: str, *, sheet: str | None = None,
                  cell: str | None = None) -> dict:
    """Tokenize + normalize one formula cell.

    `sheet` is the formula's own sheet (used to decide whether a reference is
    cross-sheet); `cell` is its A1 address (needed for relative offsets).
    Returns a deterministic dict: same input -> byte-identical output.
    """
    if text is None:
        return {"is_formula": False}
    raw = str(text).strip()
    if not raw.startswith("="):
        return {"is_formula": False, "text": raw[:400]}
    base = _parse_cell(cell) if cell else None
    functions = function_tokens(raw)
    codes = unsupported_codes(raw, set(functions))
    named = named_references(raw)
    if named and not codes:
        codes.append("UNKNOWN_NAMED_REFERENCE")
    masked, spans = _scan_spans(raw, sheet)
    pieces = []
    cursor = 0
    refs = []
    cross_sheets = []
    absolute = relative = mixed = ranges = 0
    shapes = []
    for span in spans:
        pieces.append(raw[cursor:span["start"]])
        if span["kind"] == "ref":
            start_token = _r1c1(span["start_row"], span["start_col"],
                               span["start_abs_row"], span["start_abs_col"],
                               base)
            if span.get("end_col") is not None and span.get("end_row") is not None:
                end_token = _r1c1(span["end_row"], span["end_col"],
                                 span["start_abs_row"], span["start_abs_col"],
                                 base)
                token = start_token + ":" + end_token
                ranges += 1
            else:
                token = start_token
            end_abs_col = span["start_abs_col"]
            end_abs_row = span["start_abs_row"]
        elif span["kind"] == "whole_col":
            if base is None:
                col_part = "C?"
            elif span["start_abs_col"]:
                col_part = "C" + str(span["start_col"])
            else:
                delta = span["start_col"] - base[0]
                col_part = "C" if delta == 0 else f"C[{delta}]"
            token = col_part + ":" + col_part
            end_abs_col, end_abs_row = span["start_abs_col"], False
            shapes.append("whole_column")
            ranges += 1
        else:
            if base is None:
                row_part = "R?"
            elif span["start_abs_row"]:
                row_part = "R" + str(span["start_row"])
            else:
                delta = span["start_row"] - base[1]
                row_part = "R" if delta == 0 else f"R[{delta}]"
            token = row_part + ":" + row_part
            end_abs_col, end_abs_row = False, span["start_abs_row"]
            shapes.append("whole_row")
            ranges += 1
        prefix = ""
        if span.get("sheet") and sheet and span["sheet"].strip("'") != sheet:
            prefix = _sheet_prefix(span["sheet"].strip("'"))
            cross_sheets.append(span["sheet"].strip("'"))
        elif span.get("sheet"):
            prefix = _sheet_prefix(span["sheet"].strip("'"))
            cross_sheets.append(span["sheet"].strip("'"))
        pieces.append(prefix + token)
        cursor = span["end"]
        if span["kind"] == "ref":
            if span["start_abs_col"] and span["start_abs_row"] and end_abs_col and end_abs_row:
                absolute += 1
            elif (span["start_abs_col"] or span["start_abs_row"]
                  or end_abs_col or end_abs_row):
                mixed += 1
            else:
                relative += 1
            if span.get("end_col") is not None:
                shapes.append("range")
            else:
                shapes.append("single_cell")
        refs.append({"kind": span["kind"], "token": token,
                     "sheet": span.get("sheet")})
    pieces.append(raw[cursor:])
    body = "".join(pieces)
    # upper-case outside string literals (Excel is case-insensitive there);
    # literals keep their exact bytes
    signature = _upper_outside_literals(body)
    return {
        "is_formula": True,
        "text": raw[:400],
        "sheet": sheet,
        "cell": cell,
        "sig_version": SIG_VERSION,
        "signature": signature,
        "functions": functions,
        "named_refs": named,
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


def _upper_outside_literals(text: str) -> str:
    """Upper-case everything except the contents of "..." literals."""
    chars = list(text)
    for literal in restructure.STRING_RE.finditer(text):
        chars[literal.start():literal.end()] = text[literal.start():literal.end()]
    masked = expand._masked(text)
    out = []
    for index, char in enumerate(chars):
        if masked[index] == " " and char != " ":
            out.append(char)                      # inside a literal
        else:
            out.append(char.upper())
    return "".join(out)


def main(argv=None) -> int:
    import argparse
    import json
    parser = argparse.ArgumentParser(description="tokenize/normalize formulas")
    parser.add_argument("formulas", nargs="+")
    parser.add_argument("--sheet")
    parser.add_argument("--cell")
    args = parser.parse_args(argv)
    for formula in args.formulas:
        print(json.dumps(parse_formula(formula, sheet=args.sheet,
                                       cell=args.cell),
                         ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    sys.exit(main())
