#!/usr/bin/env python3
# MIT License. Part of the Hermes xlsx skill.
"""FAZ 4 preflight instrument: build a real formula corpus, deterministically.

Read-only. Every input file's SHA-256 is measured before and after; the
script never opens a workbook for writing (`read_only=True`, no save), so
the corpus cannot change a file (AC-4.20).

For every formula cell it records the text plus a structural analysis that
reuses the ONLY parser in the codebase (`xlsx_restructure.REF_RE` /
`COORD_RE` / `STRING_RE` via ``extract_refs``) -- no second parser:

  functions, references (absolute / relative / mixed, ranges, cross-sheet),
  structured refs, external refs, array constants, future functions
  (`_xlfn.`), broken refs (`#REF!`), named references, cache availability,
  and the copy-down support verdict from ``xlsx_expand.classify_formula``.

Aggregates are complete; stored *samples* are capped and every capped list
reports ``count_total / returned_count / truncated`` (spec section 36).

Usage:
  python xlsx_formula_corpus.py --files a.xlsx b.xlsx --out corpus.json
  python xlsx_formula_corpus.py --sweep sweep_2b_results.json --out corpus.json
  python xlsx_formula_corpus.py --files ... --no-values      # skip cached values
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import time
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_common as common  # noqa: E402
import xlsx_expand as expand  # noqa: E402
import xlsx_restructure as restructure  # noqa: E402

try:
    from openpyxl import load_workbook
except ImportError as _exc:  # pragma: no cover - environment guard
    sys.exit(common.dependency_error(_exc))

CORPUS_VERSION = "4.preflight.1"

# function token: identifier immediately followed by "(" -- checked AFTER
# string literals are blanked, and excluding tokens the ref parser already
# claims are references (the parser's guards do that for us: a ref is never
# followed by "(").
FUNC_RE = re.compile(r"(?<![\w.$])(?P<name>[A-Za-z_][A-Za-z0-9_.]*)(?=\s*\()")
# workbook-qualified reference, e.g. [Book1.xlsx]Sheet1!A1 or '[Book1]S'!A1
EXTERNAL_RE = re.compile(r"\[[^\]'\s]+\]")
# single-quoted sheet name inside a reference, e.g. 'Ürün Listesi'!A1
QUOTED_SHEET_RE = re.compile(r"'(?:[^']|'')+'")
# whole-column / whole-row spans: REF_RE needs digits, so R:R must be
# masked too or the bare letters leak into the identifier scan as "named
# refs" (measured: 92k bogus hits on a large real-world sheet before this mask existed)
WHOLE_COL_RE = re.compile(
    r"(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
    r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}")
WHOLE_ROW_RE = re.compile(r"\$?[0-9]{1,7}:\$?[0-9]{1,7}")
# bare identifier that is not a function call, not a reference, not a literal
IDENT_RE = re.compile(
    r"(?<![\w.$!'\"])(?P<name>[A-Za-z_][A-Za-z0-9_.]*)"
    r"(?![\w(!\[])")
# ^ "!" and "[" exclusions matter: an unquoted sheet qualifier
#   (Listas!$A$1) is a reference prefix, and a structured table name
#   (Table_5[...]) is already counted as a structured reference -- both
#   measured as bogus named refs before (4,092 + 92,491 + 7 hits fixed here).
LITERALS = {"TRUE", "FALSE", "NULL"}
# functions that are inherently dynamic-array / late-bound
DYNAMIC_FUNCS = {"SORT", "SORTBY", "FILTER", "UNIQUE", "SEQUENCE",
                 "RANDARRAY", "XLOOKUP", "XMATCH", "LET", "LAMBDA",
                 "TEXTSPLIT", "TOCOL", "TOROW", "WRAPROWS", "WRAPCOLS",
                 "TAKE", "DROP", "EXPAND", "CHOOSEROWS", "CHOOSECOLS"}
UNSUPPORTED_PRONE = {"INDIRECT", "OFFSET", "GETPIVOTDATA", "CELL", "INFO"}

SAMPLE_CAP = 400        # stored unique formula samples per file
EXAMPLE_CAP = 60        # stored examples per unsupported reason (global)
SINGLE_EXAMPLE_CAP = 40


def sha256_of(path: Path) -> str:
    hasher = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def masked(text: str) -> str:
    """Formula text with string literals blanked (parser parity)."""
    chars = list(text)
    for literal in restructure.STRING_RE.finditer(text):
        for index in range(literal.start(), literal.end()):
            chars[index] = " "
    return "".join(chars)


FUNC_MASK_RE = re.compile(r"\[[^\]]*\]")


def function_tokens(text: str) -> list[str]:
    """Function names outside literals, quoted sheet names and [brackets]."""
    scan = masked(text)
    chars = list(scan)
    for pattern in (QUOTED_SHEET_RE, FUNC_MASK_RE):
        for span in pattern.finditer(scan):
            for index in range(span.start(), span.end()):
                chars[index] = " "
    return sorted({m.group("name").upper()
                   for m in FUNC_RE.finditer("".join(chars))})


def named_refs(text: str, functions: set[str]) -> list[str]:
    """Identifiers that are neither functions, refs, booleans, nor cell-ish.

    Quoted sheet names are blanked first: their inner words (e.g. DENET
    inside 'İÇ DENETİM RAPORU'!A1) are not named references.
    """
    scan = masked(text)
    chars = list(scan)
    # blank everything that is NOT an identifier candidate: parsed refs
    # (F$5 leaks its letter otherwise), quoted sheets, whole-col/row spans
    # and structured-ref brackets (Table_5[Substrate ...] is one ref)
    for pattern in (restructure.REF_RE, QUOTED_SHEET_RE, WHOLE_COL_RE,
                    WHOLE_ROW_RE, FUNC_MASK_RE):
        for span in pattern.finditer(scan):
            for index in range(span.start(), span.end()):
                chars[index] = " "
    out = []
    for match in IDENT_RE.finditer("".join(chars)):
        name = match.group("name")
        # adjacency must be judged on the ORIGINAL text: masking blanks the
        # very "["/"!" that makes Table_5/Listas a non-identifier
        end = match.end()
        if end < len(text) and text[end] in "[!(":
            continue
        upper = name.upper()
        if upper in functions or upper in LITERALS:
            continue
        if restructure.COORD_RE.match(name):      # A1-style coordinate
            continue
        if upper.startswith("_XLFN") or upper.startswith("_XLWS"):
            continue
        out.append(name)
    return sorted(set(out))


def ref_profile(text: str, home_sheet: str) -> dict:
    """Reference classification via the shared parser (`extract_refs`)."""
    refs = restructure.extract_refs(text, home_sheet)
    absolute = relative = mixed = 0
    ranges = single = cross = same = 0
    shapes = Counter()
    row_flags = Counter()
    for ref in refs:
        start, end = ref["start"], ref.get("end")
        col_abs, row_abs = start["abs_column"], start["abs_row"]
        if col_abs and row_abs:
            absolute += 1
        elif col_abs or row_abs:
            mixed += 1
        else:
            relative += 1
        row_flags[(row_abs, ref["range"])] += 1
        if ref["sheet"] is None:
            same += 1
        else:
            cross += 1
        if not ref["range"]:
            single += 1
            shapes["single_cell"] += 1
            continue
        ranges += 1
        if end is None:
            shapes["range_unknown_end"] += 1
            continue
        if start["row"] == end["row"] and start["column"] != end["column"]:
            shapes["horizontal"] += 1
        elif start["column"] == end["column"] and start["row"] != end["row"]:
            shapes["vertical"] += 1
        elif start["column"] == end["column"] and start["row"] == end["row"]:
            shapes["single_cell_range"] += 1
        else:
            shapes["two_dimensional"] += 1
    return {"refs_total": len(refs), "absolute": absolute,
            "relative": relative, "mixed": mixed, "ranges": ranges,
            "single_refs": single, "cross_sheet": cross, "same_sheet": same,
            "shapes": dict(shapes),
            "row_flag_counts": {f"{'abs' if a else 'rel'}_"
                                f"{'range' if r else 'point'}": c
                                for (a, r), c in row_flags.items()}}


def unsupported_reasons(text: str, functions: set[str],
                        profile: dict) -> list[str]:
    """Every reason this formula is refused by the safe machinery (if any)."""
    reasons = []
    masked_text = masked(text)
    supported, reason = expand.classify_formula(text)
    if not supported and reason and reason != "not a formula":
        reasons.append(reason)
    external = EXTERNAL_RE.search(masked_text)
    if external:
        reasons.append("external workbook reference")
    if "[" in masked_text or "]" in masked_text:
        reasons.append("structured reference")
    if "{" in masked_text or "}" in masked_text:
        reasons.append("array or constant formula")
    if "_xlfn." in masked_text.lower():
        reasons.append("future function (_xlfn)")
    if "#REF!" in masked_text:
        reasons.append("broken reference (#REF!)")
    dynamic = sorted(functions & DYNAMIC_FUNCS)
    if dynamic:
        reasons.append("dynamic array function: " + ",".join(dynamic))
    prone = sorted(functions & UNSUPPORTED_PRONE)
    if prone:
        reasons.append("late-bound function: " + ",".join(prone))
    return sorted(set(reasons))


def _defined_names(wb) -> list[dict]:
    out = []
    try:
        items = list(wb.defined_names.items())
    except Exception:  # pragma: no cover - malformed workbook
        return out
    for name, definition in items:
        text = None
        for attr in ("attr_text", "value"):
            candidate = getattr(definition, attr, None)
            if candidate is not None:
                text = str(candidate)
                break
        if text is None:
            continue
        out.append({"name": str(name), "text": text[:240],
                    "is_formula": text.startswith("="),
                    "local_sheet_id": getattr(definition, "localSheetId",
                                              None)})
    return out


def analyse_file(path: Path, *, with_values: bool = True,
                 sample_cap: int = SAMPLE_CAP) -> dict:
    record = {"file": str(path), "available": path.exists()}
    if not record["available"]:
        record["status"] = "MISSING"
        return record
    record["bytes"] = path.stat().st_size
    record["sha_before"] = sha256_of(path)
    t0 = time.perf_counter()
    try:
        wb = load_workbook(path, read_only=True, data_only=False)
    except Exception as exc:  # honest per-file failure, never a crash
        record["status"] = "LOAD_FAILED"
        record["error"] = f"{type(exc).__name__}: {exc}"[:300]
        record["sha_after"] = sha256_of(path)
        return record
    record["load_s"] = round(time.perf_counter() - t0, 2)
    record["defined_names"] = _defined_names(wb)
    sheets = list(wb.sheetnames)

    # pass 1: formula coordinates + per-formula analysis
    formula_coords: dict[str, set] = {}
    samples: dict[str, dict] = {}
    unique_texts: set[str] = set()
    formula_count = 0
    functions_counter: Counter = Counter()
    function_examples: dict[str, dict] = {}
    shape_counter: Counter = Counter()
    row_flag_counter: Counter = Counter()
    unsupported_counter: Counter = Counter()
    unsupported_examples: dict[str, list] = {}
    support_true = support_false = 0
    ref_totals = {"refs_total": 0, "absolute": 0, "relative": 0, "mixed": 0,
                  "ranges": 0, "single_refs": 0, "cross_sheet": 0,
                  "same_sheet": 0}
    with_cross = with_range = with_named = with_structured = 0
    with_external = with_array = with_xlfn = with_ref_error = 0
    t1 = time.perf_counter()
    try:
        for sheet_name in sheets:
            ws = wb[sheet_name]
            coords: set = set()
            formula_coords[sheet_name] = coords
            for row in ws.iter_rows():
                for cell in row:
                    value = cell.value
                    if value is None:
                        continue
                    text = value if isinstance(value, str) else str(value)
                    if not text.startswith("="):
                        continue
                    formula_count += 1
                    coords.add(cell.coordinate)
                    unique_texts.add(text)
                    funcs = function_tokens(text)
                    called = set(funcs)
                    for name in funcs:
                        functions_counter[name] += 1
                        if name not in function_examples:
                            function_examples[name] = {
                                "file": path.name, "sheet": sheet_name,
                                "cell": cell.coordinate, "text": text[:200]}
                    profile = ref_profile(text, sheet_name)
                    for key in ref_totals:
                        ref_totals[key] += profile[key]
                    for key, count in profile["shapes"].items():
                        shape_counter[key] += count
                    for key, count in profile["row_flag_counts"].items():
                        row_flag_counter[key] += count
                    if profile["cross_sheet"]:
                        with_cross += 1
                    if profile["ranges"]:
                        with_range += 1
                    named = named_refs(text, called)
                    if named:
                        with_named += 1
                    reasons = unsupported_reasons(text, called, profile)
                    for reason in reasons:
                        unsupported_counter[reason] += 1
                        bucket = unsupported_examples.setdefault(reason, [])
                        if len(bucket) < EXAMPLE_CAP:
                            bucket.append({"file": path.name,
                                           "sheet": sheet_name,
                                           "cell": cell.coordinate,
                                           "text": text[:240]})
                    masked_text = masked(text)
                    if EXTERNAL_RE.search(masked_text):
                        with_external += 1
                    if "[" in masked_text or "]" in masked_text:
                        with_structured += 1
                    if "{" in masked_text or "}" in masked_text:
                        with_array += 1
                    if "_xlfn." in masked_text.lower():
                        with_xlfn += 1
                    if "#REF!" in masked_text:
                        with_ref_error += 1
                    supported, _reason = expand.classify_formula(text)
                    if supported:
                        support_true += 1
                    else:
                        support_false += 1
                    if len(samples) < sample_cap and text not in samples:
                        samples[text] = {
                            "file": str(path), "sheet": sheet_name,
                            "cell": cell.coordinate, "text": text[:400],
                            "functions": funcs, "named_refs": named,
                            "refs": {k: v for k, v in profile.items()
                                     if k not in ("shapes", "row_flag_counts")},
                            "supported_copy_down": supported,
                            "unsupported": reasons}
    except Exception as exc:
        record["status"] = "SCAN_FAILED"
        record["error"] = f"{type(exc).__name__}: {exc}"[:300]
        record["formulas"] = formula_count
        record["sha_after"] = sha256_of(path)
        return record
    record["analyse_s"] = round(time.perf_counter() - t1, 2)
    try:
        wb.close()
    except Exception:
        pass

    # pass 2: cached value availability (only the formula coordinates)
    cached_total = cached_available = 0
    if with_values and formula_count:
        t2 = time.perf_counter()
        try:
            wb2 = load_workbook(path, read_only=True, data_only=True)
            for sheet_name in sheets:
                wanted = formula_coords.get(sheet_name) or set()
                if not wanted or sheet_name not in wb2.sheetnames:
                    continue
                ws2 = wb2[sheet_name]
                for row in ws2.iter_rows():
                    for cell in row:
                        coord = getattr(cell, "coordinate", None)
                        if coord is not None and coord in wanted:
                            cached_total += 1
                            if cell.value is not None:
                                cached_available += 1
            wb2.close()
        except Exception as exc:
            record["values_error"] = f"{type(exc).__name__}: {exc}"[:200]
        record["values_s"] = round(time.perf_counter() - t2, 2)

    record["_unique_texts"] = unique_texts
    record.update({
        "status": "OK",
        "sheets": len(sheets),
        "formulas": formula_count,
        "formulas_unique_exact": len(unique_texts),
        "unique_truncated": len(unique_texts) > sample_cap,
        "support": {"copy_down_supported": support_true,
                    "copy_down_refused": support_false},
        "functions": dict(functions_counter.most_common(40)),
        "function_examples": function_examples,
        "function_formulas_total": sum(functions_counter.values()),
        "refs": ref_totals,
        "shapes": dict(shape_counter),
        "row_flags": dict(row_flag_counter),
        "formulas_with": {"cross_sheet": with_cross, "range": with_range,
                          "named_ref": with_named,
                          "structured": with_structured,
                          "external": with_external, "array": with_array,
                          "xlfn": with_xlfn, "ref_error": with_ref_error},
        "unsupported": dict(unsupported_counter.most_common()),
        "unsupported_examples": unsupported_examples,
        "cached": {"checked": cached_total, "available": cached_available},
        "defined_names_total": len(record["defined_names"]),
        "defined_names_formula": len([n for n in record["defined_names"]
                                      if n["is_formula"]]),
        "samples": list(samples.values()),
    })
    record["sha_after"] = sha256_of(path)
    record["sha_unchanged"] = record["sha_after"] == record["sha_before"]
    return record


def aggregate(records: list[dict], *, unique_cap: int = 600) -> dict:
    agg = {
        "files": len(records),
        "files_ok": len([r for r in records if r.get("status") == "OK"]),
        "formulas_total": 0, "functions": Counter(), "refs": Counter(),
        "shapes": Counter(), "row_flags": Counter(),
        "unsupported": Counter(), "formulas_with": Counter(),
        "function_examples": {},
        "support": Counter(), "cached": Counter(),
        "defined_names_total": 0, "defined_names_formula": 0,
        "unique_texts": set(), "samples": [], "all_texts": set(),
        "unsupported_examples": {},
    }
    seen_samples: set = set()
    for record in records:
        if record.get("status") != "OK":
            continue
        agg["formulas_total"] += record["formulas"]
        for key, value in record["functions"].items():
            agg["functions"][key] += value
        for key, value in record["refs"].items():
            agg["refs"][key] += value
        agg["shapes"].update(record["shapes"])
        agg["row_flags"].update(record["row_flags"])
        agg["unsupported"].update(record["unsupported"])
        agg["formulas_with"].update(record["formulas_with"])
        agg["support"].update(record["support"])
        for name, example in (record.get("function_examples") or {}).items():
            agg["function_examples"].setdefault(name, example)
        agg["cached"].update(record["cached"])
        agg["all_texts"] |= record.get("_unique_texts") or set()
        agg["defined_names_total"] += record["defined_names_total"]
        agg["defined_names_formula"] += record["defined_names_formula"]
        for sample in record["samples"]:
            text = sample["text"]
            agg["unique_texts"].add(text)
            if len(seen_samples) < unique_cap and text not in seen_samples:
                seen_samples.add(text)
                agg["samples"].append(sample)
        for reason, examples in (record.get("unsupported_examples") or {}).items():
            bucket = agg["unsupported_examples"].setdefault(reason, [])
            for example in examples:
                if len(bucket) < EXAMPLE_CAP:
                    bucket.append(example)
    unique_total = len(agg.pop("unique_texts"))
    true_unique = len(agg.pop("all_texts"))
    out = {k: (dict(v) if isinstance(v, Counter) else v)
           for k, v in agg.items()}
    out["formulas_unique_true"] = true_unique
    out["formulas_unique_exact"] = unique_total
    out["samples_truncated"] = unique_total > len(out["samples"])
    out["samples_count_total"] = unique_total
    out["samples_returned"] = len(out["samples"])
    return out


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--files", nargs="*", default=[])
    parser.add_argument("--sweep", help="sweep_*_results.json with "
                                        "results[].file paths")
    parser.add_argument("--out", required=True)
    parser.add_argument("--no-values", action="store_true")
    parser.add_argument("--limit-files", type=int, default=0)
    parser.add_argument("--sample-cap", type=int, default=SAMPLE_CAP)
    args = parser.parse_args(argv)

    paths = [Path(p) for p in args.files]
    if args.sweep:
        sweep = json.loads(Path(args.sweep).read_text(encoding="utf-8"))
        rows = sweep.get("results") or sweep.get("files") or []
        for row in rows:
            candidate = row.get("file") if isinstance(row, dict) else row
            if candidate:
                paths.append(Path(candidate))
    # determinism: stable order, de-duplicated by resolved path
    seen, ordered = set(), []
    for path in paths:
        key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        ordered.append(path)
    if args.limit_files:
        ordered = ordered[:args.limit_files]

    started = time.perf_counter()
    records = [analyse_file(path, with_values=not args.no_values,
                            sample_cap=args.sample_cap) for path in ordered]
    summary = aggregate(records)
    for record in records:                      # keep the JSON lean
        record.pop("_unique_texts", None)
    payload = {
        "corpus_version": CORPUS_VERSION,
        "generated_by": "xlsx_formula_corpus.py",
        "elapsed_s": round(time.perf_counter() - started, 2),
        "no_write_proof": {
            "sha_unchanged_all": all(r.get("sha_unchanged", True)
                                     for r in records),
            "files_checked": len([r for r in records
                                  if "sha_before" in r]),
        },
        "summary": summary,
        "files": records,
        "caps": {"sample_cap_per_file": args.sample_cap,
                 "example_cap_per_reason": EXAMPLE_CAP},
    }
    Path(args.out).write_text(json.dumps(payload, indent=1, ensure_ascii=False),
                              encoding="utf-8")
    console = {"corpus_version": CORPUS_VERSION,
               "elapsed_s": payload["elapsed_s"],
               "no_write_proof": payload["no_write_proof"],
               "summary": {k: v for k, v in summary.items()
                           if k not in ("samples", "unsupported_examples",
                                         "function_examples")},
               "out": args.out}
    print(json.dumps(console, indent=1, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
