#!/usr/bin/env python3
"""FAZ 3C -- performance benchmark (spec 20, AC-3C23).

Measures the four stages separately (semantics -> plan -> dry-run ->
execute) on real workbooks of increasing size, plus the marginal cost of
row expansion. Prints a JSON table; nothing is written to the source
files (every measurement runs on a copy in the scratch area).

Usage:  python scripts/bench_pipeline_3c.py [--json out.json]
"""
from __future__ import annotations

import argparse
import json
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
import xlsx_common as common  # noqa: E402
import xlsx_execute as ex  # noqa: E402

PYTHON = sys.executable

import os  # noqa: E402

# Point XLSX_BENCH_DIR at a folder with your own workbooks (any size
# mix); missing files are reported as unavailable instead of failing.
_BENCH_DIR = os.environ.get("XLSX_BENCH_DIR", "")
REAL_FILES = [
    ("small", os.path.join(_BENCH_DIR, "small.xlsx")),
    ("medium", os.path.join(_BENCH_DIR, "medium.xlsx")),
    ("large", os.path.join(_BENCH_DIR, "large.xlsx")),
] if _BENCH_DIR else []


def timed(args, timeout=1800):
    started = time.perf_counter()
    proc = subprocess.run([str(PYTHON)] + [str(a) for a in args],
                          capture_output=True, text=True, encoding="utf-8",
                          timeout=timeout)
    return {"s": round(time.perf_counter() - started, 2),
            "rc": proc.returncode,
            "error": None if proc.returncode == 0
            else (proc.stderr or proc.stdout).strip()[-200:]}


def first_json_line(text):
    for line in reversed((text or "").strip().splitlines()):
        if line.strip().startswith("{"):
            try:
                return json.loads(line)
            except json.JSONDecodeError:
                return None
    return None


def execute_in_process(fill_plan, profile, source, target, out, manifest,
                       token):
    import contextlib
    import io
    buf = io.StringIO()
    payload, code = None, None
    try:
        with contextlib.redirect_stdout(buf):
            code = ex.main(["--fill-plan", str(fill_plan), "--profile",
                            str(profile), "--source", str(source), "--target",
                            str(target), "--out", str(out),
                            "--approve-token", str(token), "--manifest-dir",
                            str(manifest)])
        payload = first_json_line(buf.getvalue())
    except common.XlsxError as exc:
        payload = {"error_code": exc.code, "context": exc.context}
    return code, payload


def first_writable_key(profile, src_all, work, target):
    plan = work / "probe_plan.json"
    dry = work / "probe_dry.json"
    timed([SCRIPTS / "xlsx_mapping.py", "--profile", profile, "--source",
           src_all, "--emit", "plan", "--out", plan])
    timed([SCRIPTS / "xlsx_mapping.py", "--profile", profile, "--source",
           src_all, "--plan", plan, "--target", target, "--emit", "dry-run",
           "--out", dry])
    if not dry.exists():
        return None
    data = json.loads(dry.read_text(encoding="utf-8"))
    for entry in data.get("would_write") or []:
        key = str(entry.get("source", "")).split("[")[0]
        if key:
            return key
    return None


def measure_file(name, path, repeat=1):
    original = Path(path)
    record = {"file": name, "bytes": None, "available": original.exists()}
    if not record["available"]:
        return record
    work = Path(tempfile.mkdtemp(prefix=f"p3c_bench_{name}_"))
    copy = work / original.name
    shutil.copy2(original, copy)
    record["bytes"] = copy.stat().st_size
    profile = work / "profile.json"
    record["semantics"] = timed([SCRIPTS / "xlsx_semantics.py", copy,
                                 "--emit", "profile", "--out", profile])
    if record["semantics"]["rc"] != 0:
        record["result"] = "SEMANTICS_FAILED"
        return record
    prof = json.loads(profile.read_text(encoding="utf-8"))
    headers = []
    for sheet in prof.get("sheets", []):
        for column in sheet.get("columns", []):
            header = column.get("header") or column.get("name")
            if header and header not in headers:
                headers.append(header)
    src_all = work / "all.csv"
    src_all.write_text(",".join(headers) + "\n"
                       + ",".join("DEGER" for _ in headers) + "\n",
                       encoding="utf-8")
    key = first_writable_key(profile, src_all, work, copy)
    record["writable_key"] = key
    if not key:
        record["result"] = "NO_WRITABLE_KEY"
        return record
    src = work / "one.csv"
    src.write_text(f"{key}\nDEGER\n", encoding="utf-8")
    plan = work / "plan.json"
    dry = work / "dry.json"
    record["plan"] = timed([SCRIPTS / "xlsx_mapping.py", "--profile", profile,
                            "--source", src, "--emit", "plan", "--out", plan])
    record["dry_run"] = timed([SCRIPTS / "xlsx_mapping.py", "--profile",
                               profile, "--source", src, "--plan", plan,
                               "--target", copy, "--emit", "dry-run",
                               "--out", dry])
    if not dry.exists():
        record["result"] = "DRY_RUN_FAILED"
        return record
    data = json.loads(dry.read_text(encoding="utf-8"))
    token = str(data.get("plan_id"))[:12]
    times = []
    for i in range(repeat):
        out = work / f"out_{i}.xlsx"
        started = time.perf_counter()
        code, payload = execute_in_process(dry, profile, src, copy, out,
                                           work / f"mf_{i}", token)
        times.append(round(time.perf_counter() - started, 2))
        if payload is None:
            record["result"] = "EXECUTE_NO_PAYLOAD"
            return record
        record["execute_last"] = {
            "rc": code, "status": payload.get("status"),
            "write_count": payload.get("write_count"),
            "timings_s": payload.get("timings_s"),
            "error": payload.get("error_code"),
        }
    record["execute"] = {"s": times, "median_s": round(statistics.median(times), 2)}
    record["result"] = ("OK" if record["execute_last"]["rc"] == 0
                        else record["execute_last"].get("error"))
    return record


def measure_expansion(rows=(25, 100, 250)):
    """Marginal cost of inserting N rows.

    Uses the expansion-capable synthetic template (record key + template
    formula + below-block aggregate) so the run really applies an expansion
    -- a keyless template is refused by the 3B gate and would measure nothing.
    """
    sys.path.insert(0, str(SCRIPTS.parent / "tests"))
    try:
        import test_xlsx_3c_hardening as H
    except ImportError:
        return []
    out = []
    for count in rows:
        root = Path(tempfile.mkdtemp(prefix=f"p3c_bench_exp{count}_"))
        book = H.build_defter(root / "defter.xlsx")
        src = root / "src.csv"
        src.write_text("Kod,Ad,Adet,Fiyat,Durum\n" + "\n".join(
            f"K-{i:03d},Kalem {i},{i},10,OK" for i in range(1, count + 1))
            + "\n", encoding="utf-8")
        prof = root / "profile.json"
        pol = root / "policy.json"
        pol.write_text(json.dumps({"allow_expansion": True}), encoding="utf-8")
        plan = root / "plan.json"
        dry = root / "dry.json"
        timed([SCRIPTS / "xlsx_semantics.py", book, "--emit", "profile",
               "--out", prof])
        timed([SCRIPTS / "xlsx_mapping.py", "--profile", prof, "--source",
               src, "--emit", "plan", "--out", plan, "--policy-file", pol])
        timed([SCRIPTS / "xlsx_mapping.py", "--profile", prof, "--source",
               src, "--plan", plan, "--target", book, "--emit", "dry-run",
               "--out", dry, "--policy-file", pol])
        if not dry.exists():
            out.append({"rows_requested": count, "s": None,
                        "error": "DRY_RUN_FAILED"})
            continue
        data = json.loads(dry.read_text(encoding="utf-8"))
        token = str(data.get("plan_id"))[:12]
        started = time.perf_counter()
        code, payload = execute_in_process(dry, prof, src, book,
                                           root / "out.xlsx", root / "mf",
                                           token)
        payload = payload or {}
        expansion = payload.get("expansion") or {}
        out.append({"rows_requested": count,
                    "s": round(time.perf_counter() - started, 2),
                    "rc": code,
                    "status": payload.get("status"),
                    "error": payload.get("error_code"),
                    "rows_added": expansion.get("rows_added")
                    or expansion.get("actual_rows_added"),
                    "timings_s": payload.get("timings_s"),
                    "writes": payload.get("write_count")})
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", metavar="PATH")
    args = parser.parse_args()
    report = {"files": [measure_file(n, p) for n, p in REAL_FILES],
              "expansion_marginal": measure_expansion(),
              "python": sys.version.split()[0],
              "notes": ("subprocess stages include interpreter start-up; "
                        "execute timings are measured in-process")}
    text = json.dumps(report, indent=2, ensure_ascii=False)
    if args.json:
        Path(args.json).write_text(text, encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
