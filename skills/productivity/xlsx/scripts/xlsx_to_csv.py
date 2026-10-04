#!/usr/bin/env python3
# MIT License. Part of the Hermes xlsx skill.
"""Export one sheet of an .xlsx workbook to CSV.

Dates/datetimes are written in ISO format; None becomes an empty field.
With --data-only, formula cells yield their cached results (present only
if the file was last saved by Excel/LibreOffice; openpyxl never computes).

The workbook is opened read-only through the shared loader; --data-only
uses the values_only handle, which structurally cannot be written back.

Usage:
  xlsx_to_csv.py book.xlsx out.csv
  xlsx_to_csv.py book.xlsx out.csv --sheet Data --encoding utf-8-sig
  xlsx_to_csv.py book.xlsx out.csv --delimiter ';' --data-only
"""
from __future__ import annotations

import argparse
import csv
import sys
from datetime import date, datetime, time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_common as common  # noqa: E402


def to_text(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        # Excel stores pure dates as midnight datetimes; emit a bare date.
        if value.time() == time(0, 0):
            return value.date().isoformat()
        return value.isoformat()
    if isinstance(value, (date, time)):
        return value.isoformat()
    return value


def main(argv=None):
    ap = argparse.ArgumentParser(description=".xlsx sheet -> CSV exporter.")
    ap.add_argument("file", help="input .xlsx path")
    ap.add_argument("output", help="output CSV path")
    ap.add_argument("--sheet", help="sheet name (default: active)")
    ap.add_argument("--encoding", default="utf-8",
                    help="CSV output encoding (default utf-8)")
    ap.add_argument("--delimiter", default=",")
    ap.add_argument("--data-only", action="store_true",
                    help="cached formula results instead of formula strings")
    args = ap.parse_args(argv)

    loaded = common.load_workbook_safe(args.file, data_only=args.data_only)
    if args.sheet:
        if args.sheet not in loaded.sheetnames:
            raise common.XlsxError(
                "SHEET_NOT_FOUND",
                f"Worksheet '{args.sheet}' does not exist.",
                recovery="Use xlsx_read.py --sheets to list sheet names.",
                context={"requested": args.sheet,
                         "available": loaded.sheetnames})
        name = args.sheet
    else:
        name = loaded.active_sheetname

    if loaded.mode == "values_only":
        row_iter = loaded.iter_values(name)
    else:
        row_iter = loaded.wb[name].iter_rows(values_only=True)

    with open(args.output, "w", newline="", encoding=args.encoding) as fh:
        writer = csv.writer(fh, delimiter=args.delimiter)
        count = 0
        for row in row_iter:
            writer.writerow([to_text(v) for v in row])
            count += 1

    result = common.Result(mode="xlsx_to_csv", sheet=name,
                           output=args.output, rows=count)
    return result.emit()


if __name__ == "__main__":
    sys.exit(common.guard(main)())
