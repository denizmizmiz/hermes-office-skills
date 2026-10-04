#!/usr/bin/env python3
# MIT License. Part of the Hermes xlsx skill.
"""Recalculate a workbook's formulas headlessly with LibreOffice.

openpyxl never computes formulas. This script shells out to `soffice`
(LibreOffice) to open the workbook, recalculate, and re-save it, so
cached formula results become available to `xlsx_read.py --data-only`
and `--formulas`.

Behavior:
  * soffice on PATH: converts the file to .xlsx in a temp dir (which
    recalculates all formulas), validates the produced file, backs up
    whatever it replaces, and swaps the result in atomically (or writes
    --out). Prints {"recalculated": true, ...} and exits 0.
  * soffice absent: prints {"recalculated": false, "reason": ...} with
    installation guidance and STILL exits 0 — callers can branch on the
    JSON instead of the exit code.

Note: LibreOffice recalculates .xlsx on load per its default
calculation settings; conversion re-saves with fresh cached values.

Usage:
  xlsx_recalc.py book.xlsx
  xlsx_recalc.py book.xlsx --out recalced.xlsx
  xlsx_recalc.py book.xlsx --timeout 120
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_common as common  # noqa: E402


def count_cached(path):
    """Number of formula cells with a cached value present (Pass A + B)."""
    from openpyxl.utils import coordinate_to_tuple  # noqa: PLC0415

    handle_f = common.load_workbook_safe(path, read_only=True)
    handle_v = common.load_workbook_safe(path, read_only=True, data_only=True)

    cached_coords = set()
    for name in handle_v.sheetnames:
        for row_index, row in enumerate(handle_v.iter_values(name), start=1):
            for col_index, value in enumerate(row, start=1):
                if value is not None:
                    cached_coords.add((name, row_index, col_index))

    formulas = cached = 0
    for name in handle_f.sheetnames:
        ws_f = handle_f.wb[name]
        for row in ws_f.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formulas += 1
                    row_i, col_i = coordinate_to_tuple(cell.coordinate)
                    if (name, row_i, col_i) in cached_coords:
                        cached += 1
    return formulas, cached


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Recalculate .xlsx formulas headlessly via LibreOffice.")
    ap.add_argument("file", help="path to .xlsx file")
    ap.add_argument("--out", help="output path (default: replace input)")
    ap.add_argument("--timeout", type=int, default=180,
                    help="seconds to wait for soffice (default 180)")
    args = ap.parse_args(argv)

    src = Path(args.file).resolve()
    if not src.exists():
        raise common.XlsxError(
            "FILE_NOT_FOUND",
            f"No such file: {src}",
            recovery="Check the path and try again.",
            context={"path": str(src)})

    result = common.Result(mode="recalc", file=str(src))

    soffice = shutil.which("soffice")
    if not soffice:
        result.set(
            recalculated=False,
            reason="LibreOffice (soffice) not found on PATH",
            guidance="Install LibreOffice (e.g. `apt install "
                     "libreoffice-calc` or `brew install --cask "
                     "libreoffice`), or open the file in Excel/"
                     "LibreOffice once and re-save it.")
        return result.emit()

    with tempfile.TemporaryDirectory() as tmp:
        proc = subprocess.run(
            [soffice, "--headless", "--calc", "--convert-to", "xlsx:Calc "
             "MS Excel 2007 XML", "--outdir", tmp, str(src)],
            capture_output=True, text=True, encoding="utf-8",
            timeout=args.timeout,
            env={"HOME": tmp, "PATH": Path(soffice).parent.as_posix()
                 + ":/usr/bin:/bin"})
        produced = Path(tmp) / (src.stem + ".xlsx")
        if proc.returncode != 0 or not produced.exists():
            raise common.XlsxError(
                "WRITE_FAILED",
                "soffice conversion failed",
                recovery="Check that LibreOffice can open this workbook "
                         "manually; context.stderr has the details.",
                context={"stderr": proc.stderr.strip()[-500:]})

        # F: validate the produced file before it replaces anything.
        probe = common.verify_workbook(produced)
        if not probe.get("ok"):
            raise common.XlsxError(
                "VALIDATION_FAILED",
                f"Recalculated output did not validate: {probe.get('reason')}",
                recovery="The original file was not modified.",
                context={"produced": str(produced), "probe": probe})

        formulas, cached = count_cached(produced)

        # F: back up whatever we replace, stage next to the destination,
        # then commit atomically.
        dest = Path(args.out).resolve() if args.out else src
        backups = []
        if dest.exists():
            backups.append(common.backup_workbook(dest, tag="recalc"))
        fd, stage_name = tempfile.mkstemp(
            prefix=f".{dest.stem}.recalc-", suffix=".xlsx",
            dir=str(dest.parent))
        os.close(fd)
        shutil.copyfile(produced, stage_name)
        placed = common.atomic_replace(stage_name, dest)

    result.set(recalculated=True, output=str(dest),
               formula_cells=formulas, with_cached_values=cached,
               atomic=placed["atomic"],
               backup=backups[0]["backup"] if backups else None)
    if backups:
        result.diagnose(
            f"previous file backed up to {backups[0]['backup']}")
    return result.emit()


if __name__ == "__main__":
    sys.exit(common.guard(main)())
