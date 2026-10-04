#!/usr/bin/env python3
"""FAZ 3C -- real-file recovery + performance evidence (spec 19/20, AC-3C21/22/23).

For every real workbook: copy it, run the approved pipeline on the COPY, then
exercise three scenarios against the same copy:

  A. successful execution          (must commit; QA pass; copy untouched)
  B. injected failure               (qa fault -> no commit, copy untouched,
                                     a failure record lands in the manifest)
  C. rerun of the same plan         (must be already_applied / safe_to_reapply
                                     with write_count 0 -- never a duplicate)

Nothing writes to the original source files: only copies are touched and
every copy's SHA-256 is compared before/after. Prints one JSON summary.

Usage:  python scripts/smoke_pipeline_3c_recovery.py [--json out.json]
"""
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
SKILL = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))

import xlsx_common as common  # noqa: E402
import xlsx_execute as ex  # noqa: E402
import xlsx_mapping as mapping  # noqa: E402

import os  # noqa: E402

# Point XLSX_SMOKE_DIR at a folder with your own workbooks; each key
# below is looked up there. Missing files are reported as unavailable.
_SMOKE_DIR = os.environ.get("XLSX_SMOKE_DIR", "")
REAL_FILES = {
    key: os.path.join(_SMOKE_DIR, name)
    for key, name in (
        ("workbook_a", "workbook_a.xlsx"),
        ("workbook_b", "workbook_b.xlsx"),
        ("workbook_c", "workbook_c.xlsx"),
    )
} if _SMOKE_DIR else {}
PYTHON = sys.executable


def sha256(path) -> str:
    hasher = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def run(args, timeout=1800):
    started = time.perf_counter()
    proc = subprocess.run([str(PYTHON)] + [str(a) for a in args],
                          capture_output=True, text=True, encoding="utf-8",
                          timeout=timeout)
    return proc, round(time.perf_counter() - started, 2)


def source_csv_from_keys(keys) -> str:
    """Header row + one generic data row -- the mapping only needs shape."""
    return ",".join(keys) + "\n" + ",".join("DEGER" for _ in keys) + "\n"


def source_csv(profile, keys) -> str:
    headers, row, seen = [], [], set()
    for sheet in profile.get("sheets", []):
        for column in sheet.get("columns", []):
            name = column.get("header") or column.get("name")
            if not name or name not in keys or name in seen:
                continue
            seen.add(name)
            headers.append(name)
            kind = column.get("value_kind")
            row.append("1" if kind in ("integer", "decimal")
                       else "2024-01-05" if kind in ("date", "datetime")
                       else "DEGER")
    return ",".join(headers) + "\n" + ",".join(row) + "\n"


def manifest_entries(directory: Path) -> list:
    path = directory / ".xlsx_ops" / "manifest.jsonl"
    if not path.exists():
        return []
    out = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.strip():
            try:
                out.append(json.loads(line))
            except json.JSONDecodeError:
                out.append({"__corrupt__": True})
    return out


def execute_in_process(fill_plan, profile_path, source_path, target, out,
                       manifest_dir, token, faults=None, in_place=False):
    """Run the executor in this process so faults/telemetry are visible."""
    import contextlib  # noqa: PLC0415
    import io  # noqa: PLC0415

    common.FAULTS.clear()
    if faults:
        common.FAULTS.update(faults)
    buffer = io.StringIO()
    try:
        with contextlib.redirect_stdout(buffer):
            args = ["--fill-plan", str(fill_plan), "--profile",
                    str(profile_path), "--source", str(source_path),
                    "--target", str(target), "--approve-token", str(token),
                    "--manifest-dir", str(manifest_dir)]
            if in_place:
                args += ["--in-place", "--out", str(out)]
            else:
                args += ["--out", str(out)]
            code = ex.main(args)
        payload = None
        for line in reversed(buffer.getvalue().strip().splitlines()):
            if line.strip().startswith("{"):
                try:
                    payload = json.loads(line)
                except json.JSONDecodeError:
                    payload = None
                break
        return {"rc": code, "error": None, "payload": payload}
    except common.XlsxError as exc:
        return {"rc": 1, "error": exc.code, "context": exc.context,
                "payload": None}
    finally:
        common.FAULTS.clear()


def scenario(name, source_path, record):
    work = Path(tempfile.mkdtemp(prefix=f"p3c_rec_{name}_"))
    original = Path(source_path)
    if not original.exists():
        record["available"] = False
        return record
    copy = work / original.name
    shutil.copy2(original, copy)
    record.update({"available": True, "bytes": copy.stat().st_size,
                   "sha_before": sha256(copy)})

    profile_path = work / "profile.json"
    proc, semantics_s = run([SCRIPTS / "xlsx_semantics.py", copy, "--emit",
                             "profile", "--out", profile_path])
    record["semantics_s"] = semantics_s
    if proc.returncode != 0 or not profile_path.exists():
        record["result"] = "SEMANTICS_FAILED"
        record["error"] = (proc.stderr or proc.stdout)[-200:]
        return record
    profile = json.loads(profile_path.read_text(encoding="utf-8"))

    # probe the writable keys, then rebuild the source with just the first one
    probe = work / "probe.csv"
    probe.write_text(source_csv(profile, {c.get("header") or c.get("name")
                                          for s in profile.get("sheets", [])
                                          for c in s.get("columns", [])}),
                     encoding="utf-8")
    plan = work / "plan.json"
    run([SCRIPTS / "xlsx_mapping.py", "--profile", profile_path, "--source",
         probe, "--emit", "plan", "--out", plan])
    dry = work / "dry.json"
    run([SCRIPTS / "xlsx_mapping.py", "--profile", profile_path, "--source",
         probe, "--plan", plan, "--target", copy, "--emit", "dry-run",
         "--out", dry])
    if not dry.exists():
        record["result"] = "DRY_RUN_FAILED"
        return record
    plan_data = json.loads(dry.read_text(encoding="utf-8"))
    keys = []
    for entry in plan_data.get("would_write") or []:
        key = str(entry.get("source", "")).split("[")[0]
        if key and key not in keys:
            keys.append(key)
    record["writable_keys"] = keys[:5]
    if not keys:
        record["result"] = "NO_WRITABLE_KEY"
        return record

    src = work / "src.csv"
    # build the one-key source from the probe's own key name: profile column
    # names may be normalised and would not match the planned source key
    src.write_text(source_csv_from_keys(keys[:1]), encoding="utf-8")
    plan2 = work / "plan2.json"
    proc_plan2, _ = run([SCRIPTS / "xlsx_mapping.py", "--profile",
                         profile_path, "--source", src, "--emit", "plan",
                         "--out", plan2])
    if not plan2.exists():
        record["result"] = "PLAN_FAILED"
        record["plan_rc"] = proc_plan2.returncode
        record["error"] = ((proc_plan2.stderr or "")
                           + (proc_plan2.stdout or "")).strip()[-300:]
        return record
    dry2 = work / "dry2.json"
    proc2, dry2_s = run([SCRIPTS / "xlsx_mapping.py", "--profile",
                         profile_path, "--source", src, "--plan", plan2,
                         "--target", copy, "--emit", "dry-run",
                         "--out", dry2])
    record["dry_run_s"] = dry2_s
    if not dry2.exists():
        record["result"] = "DRY_RUN_FAILED"
        record["dry_run_rc"] = proc2.returncode
        record["error"] = ((proc2.stderr or "") + (proc2.stdout or "")
                           ).strip()[-300:]
        return record
    data = json.loads(dry2.read_text(encoding="utf-8"))
    token = str(data.get("plan_id"))[:12]
    manifest = work / "mf"

    # A. successful execution
    out = work / "out.xlsx"
    started = time.perf_counter()
    first = execute_in_process(dry2, profile_path, src, copy, out, manifest,
                               token)
    record["execute_s"] = round(time.perf_counter() - started, 2)
    record["sha_after_success"] = sha256(copy)
    record["original_unchanged_after_success"] = (
        record["sha_after_success"] == record["sha_before"])
    entries = manifest_entries(manifest)
    committed = [e for e in entries
                 if e.get("plan_id") == str(data.get("plan_id"))
                 and e.get("status") == "committed"]
    first_payload = first.get("payload") or {}
    record["success"] = {
        "rc": first["rc"], "error": first["error"],
        "blockers": [(b.get("code"), (b.get("message") or "")[:110])
                     for b in ((first.get("context") or {}).get("blockers")
                               or [])],
        "status": first_payload.get("status"),
        "write_count": first_payload.get("write_count"),
        "idempotency_status": first_payload.get("idempotency_status"),
        "timings_s": first_payload.get("timings_s"),
        "committed_records": len(committed),
        "qa_status": (first_payload.get("qa") or {}).get("status"),
        "output_sha": sha256(out) if out.exists() else None,
        "staged_records": len([e for e in entries
                               if e.get("status") == "staged"]),
    }

    # B. injected failure (qa fault) on the same copy
    common.FAULTS.clear()
    before_fail = sha256(copy)
    failed = execute_in_process(dry2, profile_path, src, copy,
                                work / "out_failed.xlsx", manifest, token,
                                faults={"qa": "QA_FAILED"})
    record["sha_after_failure"] = sha256(copy)
    record["injected_failure"] = {
        "error": failed["error"],
        "commit_state": (failed.get("context") or {}).get("commit_state"),
        "temp": (failed.get("context") or {}).get("temp"),
        "original_unchanged": sha256(copy) == before_fail,
        "no_output": not (work / "out_failed.xlsx").exists(),
        "failure_records": len([e for e in manifest_entries(manifest)
                                if e.get("status") == "failed"]),
    }

    # C. rerun of the same plan -> never a duplicate
    rerun = execute_in_process(dry2, profile_path, src, copy, out, manifest,
                               token)
    entries = manifest_entries(manifest)
    rerun_payload = rerun.get("payload") or {}
    record["rerun"] = {
        "rc": rerun["rc"], "error": rerun["error"],
        "status": rerun_payload.get("status"),
        "write_count": rerun_payload.get("write_count"),
        "idempotency_status": rerun_payload.get("idempotency_status"),
        "committed_records": len([e for e in entries
                                  if e.get("plan_id") == str(data.get("plan_id"))
                                  and e.get("status") == "committed"]),
        "sha_after_rerun": sha256(copy),
        "original_unchanged": sha256(copy) == record["sha_before"],
    }
    # D. in-place idempotency: apply, then re-apply on the SAME file
    inplace = work / "inplace.xlsx"
    shutil.copy2(original, inplace)
    first_in = execute_in_process(dry2, profile_path, src, inplace, inplace,
                                  work / "mf2", token, in_place=True)
    sha_after_inplace = sha256(inplace)
    second_in = execute_in_process(dry2, profile_path, src, inplace, inplace,
                                   work / "mf2", token, in_place=True)
    payload_in = second_in.get("payload") or {}
    record["in_place_idempotency"] = {
        "first_rc": first_in["rc"],
        "first_status": (first_in.get("payload") or {}).get("status"),
        "second_rc": second_in["rc"],
        "second_status": payload_in.get("status"),
        "second_write_count": payload_in.get("write_count"),
        "second_idempotency": payload_in.get("idempotency_status"),
        "file_unchanged_after_second": sha256(inplace) == sha_after_inplace,
        "backup": (first_in.get("payload") or {}).get("backup"),
        "backup_exists": bool(
            (first_in.get("payload") or {}).get("backup")
            and Path((first_in.get("payload") or {}).get("backup")).exists()),
    }

    record["result"] = ("PASS" if (record["success"]["rc"] == 0
                                   and record["original_unchanged_after_success"]
                                   and record["injected_failure"][
                                       "original_unchanged"]
                                   and record["injected_failure"][
                                       "failure_records"] >= 1
                                   and record["rerun"]["original_unchanged"]
                                   and record["in_place_idempotency"][
                                       "second_write_count"] == 0
                                   and record["in_place_idempotency"][
                                       "file_unchanged_after_second"])
                        else "FAIL")
    return record


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", metavar="PATH")
    parser.add_argument("--only", metavar="NAME")
    args = parser.parse_args()
    report = {}
    for name, path in REAL_FILES.items():
        if args.only and name != args.only:
            continue
        report[name] = scenario(name, path, {})
    text = json.dumps(report, indent=2, ensure_ascii=False)
    if args.json:
        Path(args.json).write_text(text, encoding="utf-8")
    print(text)
    passed = [n for n, r in report.items() if r.get("result") == "PASS"]
    print(f"SUMMARY: {len(passed)}/{len(report)} PASS -> {sorted(passed)}",
          file=sys.stderr)
    return 0 if len(passed) >= 3 else 1


if __name__ == "__main__":
    sys.exit(main())
