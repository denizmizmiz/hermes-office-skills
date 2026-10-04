#!/usr/bin/env python3
# MIT License. Part of the Hermes xlsx skill.
"""Health-check an .xlsx workbook and report a machine-readable result.

Phase-1 scope: this is a *technical* validation tier, not the Phase-5 QA
layer. It answers "is this package structurally sound and readable?" -- it
never claims that formulas were evaluated (openpyxl does not compute).

Checks performed:
  * file exists, is a regular readable file
  * the file is a valid zip package ([Content_Types].xml present)
  * openpyxl opens it (normal mode) and the sheet list is readable
  * sheet count matches --expect-sheets when given
  * a second read-only pass counts formula texts and cached values;
    a normal pass counts tables / merged ranges / defined names /
    data validations / conditional formats / charts / images per sheet

Honesty rule (architecture 12.4): ``formula_text_ok`` and
``values_recalculated`` are separate fields. This script NEVER sets
``values_recalculated`` to true -- real recalculation only happens via
xlsx_recalc.py (LibreOffice). A missing LibreOffice is NOT a failure here.

Usage:
  xlsx_validate.py file.xlsx
  xlsx_validate.py file.xlsx --expect-sheets "Data,Notes"
  xlsx_validate.py file.xlsx --expect-tables 2

Exit codes: 0 when all checks pass, 1 when any check fails (the JSON still
carries per-check detail). A missing file exits 1 with FILE_NOT_FOUND.
"""
from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_common as common  # noqa: E402


def collect_sheet_counts(ws) -> dict:
    formulas = 0
    for row in ws.iter_rows():
        for cell in row:
            value = cell.value
            if isinstance(value, str) and value.startswith("="):
                formulas += 1
    return {
        "name": ws.title,
        "dimensions": ws.dimensions,
        "max_row": ws.max_row,
        "max_col": ws.max_column,
        "formulas": formulas,
        "tables": len(getattr(ws, "tables", {})),
        "merged": len(ws.merged_cells.ranges),
        "validations": len(ws.data_validations.dataValidation),
        "conditional_formats": len(list(ws.conditional_formatting)),
        "charts": len(getattr(ws, "_charts", [])),
        "images": len(getattr(ws, "_images", [])),
        "protected": bool(ws.protection.sheet),
    }


def validate(path, *, expect_sheets=None, expect_tables=None,
             result: common.Result) -> common.Result:
    target = Path(path)
    checks = []

    def check(check_id, ok, detail=None):
        entry = {"id": check_id, "ok": bool(ok)}
        if detail is not None:
            entry["detail"] = detail
        checks.append(entry)
        if not ok:
            result.set(ok=False)
        return ok

    # -- file level ---------------------------------------------------------
    if not check("file_exists", target.exists(), str(target)):
        raise common.XlsxError("FILE_NOT_FOUND", f"No such file: {target}",
                               recovery="Check the path and try again.",
                               context={"path": str(target)})
    check("file_readable", target.is_file() and target.stat().st_size > 0,
          f"size={target.stat().st_size}")

    package_ok = True
    try:
        with zipfile.ZipFile(target) as zf:
            names = zf.namelist()
            package_ok = "[Content_Types].xml" in names
            if not package_ok:
                check("package_content_types", False,
                      "[Content_Types].xml missing -> not an OOXML package")
            else:
                check("package_content_types", True)
            bad = zf.testzip()
            check("package_members", bad is None, f"first bad member: {bad}")
    except zipfile.BadZipFile as exc:
        package_ok = False
        check("package_readable", False, str(exc))
    except OSError as exc:
        package_ok = False
        check("package_readable", False, str(exc))

    if not package_ok:
        result.set(checks=checks)
        return result

    # -- workbook opens (normal mode) ----------------------------------------
    try:
        loaded = common.load_workbook_safe(target)
        wb = loaded.wb
        check("workbook_opens", True)
    except common.XlsxError as exc:
        check("workbook_opens", False, exc.message)
        result.set(checks=checks)
        return result

    sheets = list(wb.sheetnames)
    check("sheet_list_readable", bool(sheets), f"{len(sheets)} sheet(s)")
    if expect_sheets is not None:
        check("sheet_count_stable",
              sorted(sheets) == sorted(expect_sheets),
              f"expected {sorted(expect_sheets)}, found {sorted(sheets)}")
    else:
        check("sheet_count_reported", True, len(sheets))

    per_sheet = [collect_sheet_counts(ws) for ws in wb.worksheets]
    total_tables = sum(s["tables"] for s in per_sheet)
    if expect_tables is not None:
        check("table_count_stable", total_tables == expect_tables,
              f"expected {expect_tables}, found {total_tables}")

    defined_names = {
        name: dn.attr_text for name, dn in wb.defined_names.items()}
    formula_texts = sum(s["formulas"] for s in per_sheet)

    # -- read-only passes: formula text vs cached values (honest split) ------
    cached_values = 0
    read_ok = True
    ro_formula_cells = []
    try:
        ro = common.load_workbook_safe(target, read_only=True)
        for ws in ro.wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        ro_formula_cells.append((ws.title, cell.coordinate))
        ro.close()
        check("formula_pass_opens", True,
              f"{len(ro_formula_cells)} formula cell(s) in read-only pass")
    except common.XlsxError as exc:
        read_ok = False
        check("formula_pass_opens", False, exc.message)

    try:
        vo = common.load_workbook_safe(target, read_only=True, data_only=True)
        if ro_formula_cells:
            from openpyxl.utils import coordinate_to_tuple  # noqa: PLC0415
            want = {}
            for sheet_name, coord in ro_formula_cells:
                want.setdefault(sheet_name, set()).add(
                    coordinate_to_tuple(coord))
            for sheet_name in vo.sheetnames:
                coords = want.get(sheet_name)
                if not coords:
                    continue
                for row_index, row in enumerate(
                        vo.iter_values(sheet_name), start=1):
                    for col_index, value in enumerate(row, start=1):
                        if value is not None and (row_index, col_index) in coords:
                            cached_values += 1
        vo.close()
        check("cached_pass_opens", True, f"{cached_values} cached value(s)")
    except common.XlsxError as exc:
        check("cached_pass_opens", False, exc.message)

    result.set(
        checks=checks,
        file=str(target),
        file_fingerprint=common.fingerprint_file(target),
        sheets=sheets,
        sheet_count=len(sheets),
        per_sheet=per_sheet,
        defined_names=defined_names,
        defined_name_count=len(defined_names),
        table_count=total_tables,
        formula_text_count=formula_texts,
        formula_text_ok=read_ok,
        cached_value_count=cached_values,
        # Honesty rule: this script never recalculates; only xlsx_recalc.py can.
        values_recalculated=False,
        recalculation_note=(
            "openpyxl never evaluates formulas. Run xlsx_recalc.py to "
            "materialize cached values via LibreOffice; its absence is not a "
            "failure of this validation."),
    )
    if not read_ok:
        result.warn("FEATURE_PARTIAL",
                    "The read-only formula pass failed; formula counts are "
                    "from the normal pass only.")
    return result


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Validate an .xlsx workbook (structural health-check).",
        epilog="Never claims formulas were recalculated; see xlsx_recalc.py.")
    ap.add_argument("file", help="path to .xlsx file")
    ap.add_argument("--expect-sheets", metavar="A,B,C",
                    help="fail when the sheet list differs from this list")
    ap.add_argument("--expect-tables", type=int, metavar="N",
                    help="fail when the workbook has a different table count")
    args = ap.parse_args(argv)

    expect = None
    if args.expect_sheets is not None:
        expect = [s.strip() for s in args.expect_sheets.split(",") if s.strip()]

    result = common.Result(mode="validate", file=args.file)
    try:
        validate(args.file, expect_sheets=expect,
                 expect_tables=args.expect_tables, result=result)
    except common.XlsxError as exc:
        return common.emit_error(exc.code, exc.message, recovery=exc.recovery,
                                 context=exc.context)
    failed = [c for c in result.get("checks", []) if not c["ok"]]
    if failed:
        result.warn("FEATURE_PARTIAL", f"{len(failed)} check(s) failed.")
    return result.emit(exit_code=1 if failed else 0)


if __name__ == "__main__":
    sys.exit(common.guard(main)())
