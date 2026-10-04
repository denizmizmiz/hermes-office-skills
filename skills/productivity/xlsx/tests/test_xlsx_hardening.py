"""Phase-1 hardening tests for the xlsx skill.

Additive companion to test_xlsx_skill.py (which stays untouched as the
regression suite): safety invariants (I1-I5), structured diagnostics,
D4/D5 bug fixes, fingerprints and the operation manifest.

Run:  pytest tests/ -q
"""
from __future__ import annotations

import ast
import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"


def run(script, *args, expect_ok=True, env=None):
    base_env = dict(os.environ, LC_ALL="C", LANG="C")
    base_env.pop("PYTHONIOENCODING", None)
    if env:
        base_env.update(env)
    proc = subprocess.run(
        [sys.executable, str(SCRIPTS / script), *map(str, args)],
        capture_output=True, text=True, env=base_env, encoding="utf-8")
    if expect_ok:
        assert proc.returncode == 0, f"{script} failed: {proc.stderr}"
    return proc


def run_common(code, *args):
    """Run a snippet with scripts/ on sys.path; args land in sys.argv[1:]."""
    snippet = ("import sys\nsys.path.insert(0, '%s')\n%s"
               % (SCRIPTS.as_posix(), code))
    return subprocess.run(
        [sys.executable, "-c", snippet, *map(str, args)],
        capture_output=True, text=True, encoding="utf-8")


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def make_probe(tmp_path):
    """Small workbook; native table T2 lives on the SECOND sheet (D4)."""
    spec = {
        "sheets": [
            {"name": "Ozet", "rows": [["Toplam"]],
             "cells": {"B1": {"formula": "SUM('Urun Listesi'!B2:B4)"}}},
            {"name": "Urun Listesi",
             "rows": [["Ad", "Adet"], ["a", 1], ["b", 2], ["c", 3]],
             "tables": [{"name": "T2", "range": "A1:B4"}]},
        ]
    }
    spec_path = tmp_path / "probe_spec.json"
    spec_path.write_text(json.dumps(spec), encoding="utf-8")
    probe = tmp_path / "probe.xlsx"
    run("xlsx_create.py", spec_path, probe)
    return probe, spec_path


# ---------------------------------------------------------------------------
# A. dependency hardening
# ---------------------------------------------------------------------------

def test_dependency_present_and_pinned():
    proc = run_common(
        "import json, xlsx_common as c\n"
        "print(json.dumps(c.check_dependency()))")
    assert proc.returncode == 0
    d = json.loads(proc.stdout)
    assert d["ok"] is True
    assert d["openpyxl"].startswith("3.1")


def test_dependency_missing_is_structured(tmp_path):
    shim = tmp_path / "shim"
    shim.mkdir()
    (shim / "openpyxl.py").write_text(
        'raise ImportError("simulated: openpyxl unavailable")\n',
        encoding="utf-8")
    env = {"PYTHONPATH": str(shim)}

    # guard path: script has no top-level openpyxl imports
    probe, _ = make_probe(tmp_path)
    proc = run("xlsx_read.py", probe, "--sheets", expect_ok=False, env=env)
    assert proc.returncode == 1
    d = json.loads(proc.stderr)
    assert d["ok"] is False and d["error_code"] == "DEPENDENCY_MISSING"
    assert "uv run" in d.get("recovery", "")

    # import-guard path: script imports openpyxl at module top
    spec = tmp_path / "s.json"
    spec.write_text('{"sheets": []}', encoding="utf-8")
    proc = run("xlsx_create.py", spec, tmp_path / "x.xlsx",
               expect_ok=False, env=env)
    assert proc.returncode == 1
    assert json.loads(proc.stderr)["error_code"] == "DEPENDENCY_MISSING"


# ---------------------------------------------------------------------------
# B. structured diagnostics / spec validation (D2/D3)
# ---------------------------------------------------------------------------

def test_missing_required_spec_key(tmp_path):
    spec = tmp_path / "s.json"
    spec.write_text('{"defined_names": {}}', encoding="utf-8")  # no sheets
    out = tmp_path / "o.xlsx"
    proc = run("xlsx_create.py", spec, out, expect_ok=False)
    d = json.loads(proc.stderr)
    assert d["error_code"] == "SPEC_INVALID"
    assert d["context"]["path"] == "sheets"
    assert not out.exists()


def test_unknown_key_strict_then_lenient(tmp_path):
    spec = {"sheets": [{"name": "S", "rows": [[1]],
                        "column_widhts": {"A": 5}}]}
    spec_path = tmp_path / "s.json"
    spec_path.write_text(json.dumps(spec), encoding="utf-8")
    out = tmp_path / "o.xlsx"

    proc = run("xlsx_create.py", spec_path, out, expect_ok=False)
    d = json.loads(proc.stderr)
    assert d["error_code"] == "SPEC_UNKNOWN_KEY"
    assert d["context"]["unknown_keys"] == ["sheets[0].column_widhts"]
    assert d["context"].get("nearest") == "column_widths"
    assert not out.exists()

    proc = run("xlsx_create.py", spec_path, out, "--allow-unknown-keys")
    d = json.loads(proc.stdout)
    assert d["ok"] is True
    assert [w["code"] for w in d["warnings"]] == ["UNKNOWN_SPEC_KEY"]
    assert out.exists()


def test_type_mismatch_reports_path(tmp_path):
    spec = {"sheets": [{"name": "S", "rows": [[1]], "merges": "A1:B1"}]}
    spec_path = tmp_path / "s.json"
    spec_path.write_text(json.dumps(spec), encoding="utf-8")
    proc = run("xlsx_create.py", spec_path, tmp_path / "o.xlsx",
               expect_ok=False)
    d = json.loads(proc.stderr)
    assert d["error_code"] == "SPEC_INVALID"
    assert d["context"]["path"] == "sheets[0].merges"
    assert d["context"]["expected"] == "list"


def test_table_missing_range_is_spec_error(tmp_path):
    spec = {"sheets": [{"name": "S", "rows": [[1]],
                        "tables": [{"name": "T"}]}]}
    spec_path = tmp_path / "s.json"
    spec_path.write_text(json.dumps(spec), encoding="utf-8")
    proc = run("xlsx_create.py", spec_path, tmp_path / "o.xlsx",
               expect_ok=False)
    d = json.loads(proc.stderr)
    assert d["error_code"] == "SPEC_INVALID"
    assert d["context"]["path"] == "sheets[0].tables[0].range"


def test_missing_file_and_bad_sheet_are_structured(tmp_path):
    proc = run("xlsx_read.py", tmp_path / "none.xlsx", "--sheets",
               expect_ok=False)
    assert json.loads(proc.stderr)["error_code"] == "FILE_NOT_FOUND"

    proc = run("xlsx_recalc.py", tmp_path / "none.xlsx", expect_ok=False)
    assert json.loads(proc.stderr)["error_code"] == "FILE_NOT_FOUND"

    probe, _ = make_probe(tmp_path)
    proc = run("xlsx_read.py", probe, "--json", "--sheet", "Yok",
               expect_ok=False)
    d = json.loads(proc.stderr)
    assert d["error_code"] == "SHEET_NOT_FOUND"
    assert d["context"]["requested"] == "Yok"

    proc = run("xlsx_edit.py", probe, "--sheet", "Yok", "--set", "A1=1",
               expect_ok=False)
    assert json.loads(proc.stderr)["error_code"] == "SHEET_NOT_FOUND"


# ---------------------------------------------------------------------------
# C. safety invariants (I1/I2) + write chain (F)
# ---------------------------------------------------------------------------

def test_data_only_handle_write_forbidden(tmp_path):
    probe, _ = make_probe(tmp_path)
    before = sha256(probe)
    proc = run_common(
        "import json, xlsx_common as c\n"
        "probe = sys.argv[1]\n"
        "h = c.load_workbook_safe(probe, data_only=True)\n"
        "res = {'mode': h.mode}\n"
        "try:\n"
        "    h.wb\n"
        "    res['wb_access'] = 'allowed'\n"
        "except c.XlsxError as e:\n"
        "    res['wb_access'] = e.code\n"
        "try:\n"
        "    c.save_workbook_safe(h, probe + '.out.xlsx')\n"
        "    res['save'] = 'allowed'\n"
        "except c.XlsxError as e:\n"
        "    res['save'] = e.code\n"
        "print(json.dumps(res))", probe)
    d = json.loads(proc.stdout)
    assert d["mode"] == "values_only"
    assert d["wb_access"] == "DATA_ONLY_WRITE_FORBIDDEN"
    assert d["save"] == "DATA_ONLY_WRITE_FORBIDDEN"
    assert not Path(str(probe) + ".out.xlsx").exists()
    assert sha256(probe) == before


def test_read_only_handle_write_forbidden(tmp_path):
    probe, _ = make_probe(tmp_path)
    before = sha256(probe)
    proc = run_common(
        "import json, xlsx_common as c\n"
        "probe = sys.argv[1]\n"
        "h = c.load_workbook_safe(probe, read_only=True)\n"
        "try:\n"
        "    c.save_workbook_safe(h, probe + '.out.xlsx')\n"
        "    print(json.dumps({'save': 'allowed'}))\n"
        "except c.XlsxError as e:\n"
        "    print(json.dumps({'save': e.code}))", probe)
    assert json.loads(proc.stdout)["save"] == "READ_ONLY_FILE"
    assert sha256(probe) == before


def test_backup_failure_aborts_write(tmp_path):
    probe, _ = make_probe(tmp_path)
    before = sha256(probe)
    proc = run_common(
        "import json, xlsx_common as c\n"
        "probe = sys.argv[1]\n"
        "def boom(p, tag='prewrite'):\n"
        "    raise c.XlsxError('BACKUP_FAILED', 'simulated backup failure')\n"
        "c.backup_workbook = boom\n"
        "h = c.load_workbook_safe(probe)\n"
        "try:\n"
        "    c.save_workbook_safe(h, probe, in_place=True)\n"
        "    print(json.dumps({'result': 'allowed'}))\n"
        "except c.XlsxError as e:\n"
        "    print(json.dumps({'result': e.code}))", probe)
    assert json.loads(proc.stdout)["result"] == "BACKUP_FAILED"
    assert sha256(probe) == before


def test_temp_validation_failure_keeps_original(tmp_path):
    probe, _ = make_probe(tmp_path)
    before = sha256(probe)
    target = tmp_path / "victim.xlsx"
    target.write_bytes(b"placeholder that must not be destroyed")
    t_before = sha256(target)
    proc = run_common(
        "import json, os, xlsx_common as c\n"
        "probe, target = sys.argv[1], sys.argv[2]\n"
        "c.verify_workbook = lambda p, **kw: {'ok': False, "
        "'reason': 'simulated validation failure'}\n"
        "h = c.load_workbook_safe(probe)\n"
        "try:\n"
        "    c.save_workbook_safe(h, target, backup=True)\n"
        "    out = {'result': 'allowed'}\n"
        "except c.XlsxError as e:\n"
        "    tmp = [n for n in os.listdir(os.path.dirname(target))\n"
        "           if '.tmp-' in n]\n"
        "    out = {'result': e.code, 'tmp_kept': tmp,\n"
        "           'target_intact': os.path.exists(target)}\n"
        "print(json.dumps(out))", probe, target)
    d = json.loads(proc.stdout)
    assert d["result"] == "VALIDATION_FAILED"
    assert d["tmp_kept"], "temporary output should be kept for diagnostics"
    assert d["target_intact"] is True
    assert sha256(probe) == before
    assert sha256(target) == t_before


def test_atomic_commit_backup_and_restore_bytes(tmp_path):
    probe, _ = make_probe(tmp_path)
    before = sha256(probe)
    proc = run("xlsx_edit.py", probe, "--set", "A9=1")
    d = json.loads(proc.stdout)
    assert d["ok"] is True and d["atomic"] is True
    assert d["approval"] == "caller-explicit"
    assert d["plan_hash"].startswith("sha256:")
    assert d["backup"] and Path(d["backup"]).exists()
    assert sha256(d["backup"]) == before      # backup holds pre-write bytes
    assert sha256(probe) != before            # the file was replaced

    r = json.loads(run("xlsx_read.py", probe, "--json", "--sheet",
                       "Ozet").stdout)
    assert any(row and row[0] == 1 for row in r["rows"])


def test_unacknowledged_in_place_refused(tmp_path):
    probe, _ = make_probe(tmp_path)
    proc = run_common(
        "import json, xlsx_common as c\n"
        "probe = sys.argv[1]\n"
        "h = c.load_workbook_safe(probe)\n"
        "try:\n"
        "    c.save_workbook_safe(h, probe)  # no in_place=True\n"
        "    print(json.dumps({'r': 'allowed'}))\n"
        "except c.XlsxError as e:\n"
        "    print(json.dumps({'r': e.code}))", probe)
    assert json.loads(proc.stdout)["r"] == "READ_ONLY_FILE"


def test_require_approval_gate(tmp_path):
    probe, _ = make_probe(tmp_path)
    proc = run_common(
        "import json, xlsx_common as c\n"
        "probe = sys.argv[1]\n"
        "h = c.load_workbook_safe(probe)\n"
        "out = {}\n"
        "try:\n"
        "    c.save_workbook_safe(h, probe, in_place=True, "
        "require_approval=True)\n"
        "    out['no_token'] = 'allowed'\n"
        "except c.XlsxError as e:\n"
        "    out['no_token'] = e.code\n"
        "info = c.save_workbook_safe(h, probe, in_place=True, "
        "require_approval=True, approve_token='abc123abc123')\n"
        "out['with_token'] = info['approval']\n"
        "print(json.dumps(out))", probe)
    d = json.loads(proc.stdout)
    assert d["no_token"] == "PLAN_CONFLICT"
    assert d["with_token"] == "token:abc123abc123"


# ---------------------------------------------------------------------------
# D. raw-save lint (AST)  --  I5 structural guard
# ---------------------------------------------------------------------------

def test_no_raw_workbook_save_outside_common():
    offenders = []
    for script in sorted(SCRIPTS.glob("*.py")):
        if script.name == "xlsx_common.py":
            continue
        tree = ast.parse(script.read_text(encoding="utf-8"),
                         filename=str(script))
        for node in ast.walk(tree):
            if (isinstance(node, ast.Call)
                    and isinstance(node.func, ast.Attribute)
                    and node.func.attr == "save"):
                offenders.append(f"{script.name}:{node.lineno}")
    assert offenders == [], f"forbidden raw save calls: {offenders}"


# ---------------------------------------------------------------------------
# E. D4 / D5 bug fixes
# ---------------------------------------------------------------------------

def test_d4_table_append_without_sheet(tmp_path):
    probe, _ = make_probe(tmp_path)
    proc = run("xlsx_edit.py", probe, "--table-append", 'T2=["d", 4]')
    d = json.loads(proc.stdout)
    assert d["ok"] is True
    assert d["changes"] == ["table_append T2 -> A1:B5"]
    r = json.loads(run("xlsx_read.py", probe, "--json", "--sheet",
                       "Urun Listesi").stdout)
    assert r["rows"][-1] == ["d", 4]


def test_d4_table_not_found_has_context(tmp_path):
    probe, _ = make_probe(tmp_path)
    proc = run("xlsx_edit.py", probe, "--table-append", "NOPE=[1]",
               expect_ok=False)
    d = json.loads(proc.stderr)
    assert d["error_code"] == "TABLE_NOT_FOUND"
    assert d["context"]["requested_table"] == "NOPE"
    assert "Ozet" in d["context"]["searched_sheets"]
    assert "Urun Listesi" in d["context"]["searched_sheets"]


def test_d4_sheet_scoping(tmp_path):
    probe, _ = make_probe(tmp_path)
    # explicit --sheet searches only that sheet
    proc = run("xlsx_edit.py", probe, "--sheet", "Ozet",
               "--table-append", "T2=[1]", expect_ok=False)
    d = json.loads(proc.stderr)
    assert d["error_code"] == "TABLE_NOT_FOUND"
    assert d["context"]["searched_sheets"] == ["Ozet"]
    # the sheet that owns the table still works
    proc = run("xlsx_edit.py", probe, "--sheet", "Urun Listesi",
               "--table-append", "T2=[9, 9]")
    assert json.loads(proc.stdout)["ok"] is True


def test_d5_copy_then_rename_single_call(tmp_path):
    probe, _ = make_probe(tmp_path)
    proc = run("xlsx_edit.py", probe,
               "--copy-sheet", "Ozet:Kopya", "--rename-sheet", "Ozet:Ana")
    d = json.loads(proc.stdout)
    assert d["ok"] is True
    assert d["changes"] == ["copy Ozet->Kopya", "rename Ozet->Ana"]
    sheets = json.loads(run("xlsx_read.py", probe, "--sheets").stdout)
    assert [s["name"] for s in sheets["sheets"]] == \
        ["Ana", "Urun Listesi", "Kopya"]


def test_d5_rename_only_still_works(tmp_path):
    probe, _ = make_probe(tmp_path)
    proc = run("xlsx_edit.py", probe, "--rename-sheet", "Ozet:Yeni")
    d = json.loads(proc.stdout)
    assert d["changes"] == ["rename Ozet->Yeni"]
    assert d["sheet"] == "Yeni"


# ---------------------------------------------------------------------------
# F. fingerprints
# ---------------------------------------------------------------------------

def test_fingerprints_deterministic(tmp_path):
    probe, _ = make_probe(tmp_path)
    proc = run_common(
        "import json, xlsx_common as c\n"
        "probe = sys.argv[1]\n"
        "f1, f2 = c.fingerprint_file(probe), c.fingerprint_file(probe)\n"
        "h = c.load_workbook_safe(probe, read_only=True)\n"
        "ws = h.wb['Urun Listesi']\n"
        "r1, r2 = (c.fingerprint_region(ws, 'A1:B4'),\n"
        "          c.fingerprint_region(ws, 'A1:B4'))\n"
        "other = c.fingerprint_region(h.wb['Ozet'], 'A1:A1')\n"
        "hn = c.load_workbook_safe(probe)\n"
        "s1 = c.fingerprint_structure(hn.wb)\n"
        "s2 = c.fingerprint_structure(hn.wb)\n"
        "print(json.dumps({'f_eq': f1 == f2, 'f_prefix': "
        "f1.startswith('sha256:'), 'r_eq': r1 == r2, "
        "'r_diff_sheet': r1 != other, 's_eq': s1 == s2, "
        "'s_prefix': s1.startswith('sha256:')}))", probe)
    d = json.loads(proc.stdout)
    assert d["f_eq"] and d["f_prefix"] and d["r_eq"] and d["r_diff_sheet"]
    assert d["s_eq"] and d["s_prefix"]


def test_region_fingerprint_change_sensitivity(tmp_path):
    probe, _ = make_probe(tmp_path)
    proc = run_common(
        "import json, xlsx_common as c\n"
        "probe = sys.argv[1]\n"
        "h = c.load_workbook_safe(probe, read_only=True)\n"
        "ws = h.wb['Urun Listesi']\n"
        "before_a = c.fingerprint_region(ws, 'A1:A4')\n"
        "before_b = c.fingerprint_region(ws, 'B1:B4')\n"
        "h2 = c.load_workbook_safe(probe)\n"
        "h2.wb['Urun Listesi']['B2'] = 111\n"
        "ws2 = h2.wb['Urun Listesi']\n"
        "after_a = c.fingerprint_region(ws2, 'A1:A4')\n"
        "after_b = c.fingerprint_region(ws2, 'B1:B4')\n"
        "print(json.dumps({'unrelated_same': before_a == after_a,\n"
        "                  'target_changed': before_b != after_b}))", probe)
    d = json.loads(proc.stdout)
    assert d["unrelated_same"] is True
    assert d["target_changed"] is True


# ---------------------------------------------------------------------------
# G. manifest / idempotency foundation
# ---------------------------------------------------------------------------

def test_manifest_roundtrip_and_corruption(tmp_path):
    proc = run_common(
        "import json, xlsx_common as c\n"
        "base = sys.argv[1]\n"
        "a1 = c.manifest_append({'operation_id': 'op_test_1',\n"
        "    'plan_hash': 'sha256:abc', 'region_fingerprint': 'sha256:def',\n"
        "    'status': 'committed', 'output': 'x.xlsx'}, skill_dir=base)\n"
        "a2 = c.manifest_append({'operation_id': 'op_test_2',\n"
        "    'plan_hash': 'sha256:abc', 'region_fingerprint': 'sha256:ghi',\n"
        "    'status': 'committed'}, skill_dir=base)\n"
        "r = c.manifest_read(skill_dir=base)\n"
        "f = c.manifest_find('sha256:abc', 'sha256:def', skill_dir=base)\n"
        "d1 = c.idempotency_decision('sha256:abc', 'sha256:def', "
        "skill_dir=base)\n"
        "d2 = c.idempotency_decision('sha256:abc', 'sha256:zzz', "
        "skill_dir=base)\n"
        "d3 = c.idempotency_decision('sha256:none', 'sha256:def', "
        "skill_dir=base)\n"
        "mp = c.manifest_path(skill_dir=base)\n"
        "with open(mp, 'a', encoding='utf-8') as fh:\n"
        "    fh.write('{broken json\\n')\n"
        "r2 = c.manifest_read(skill_dir=base)\n"
        "print(json.dumps({'a': a1['ok'] and a2['ok'],\n"
        "                  'n': len(r['entries']),\n"
        "                  'find': (f or {}).get('operation_id'),\n"
        "                  'd1': d1['status'], 'd2': d2['status'],\n"
        "                  'd3': d3['status'],\n"
        "                  'corrupt': len(r2.get('corrupt_lines', []))}))",
        tmp_path)
    d = json.loads(proc.stdout)
    assert d["a"] is True and d["n"] == 2
    assert d["find"] == "op_test_1"
    assert d["d1"] == "already_applied"
    assert d["d2"] == "conflict"
    assert d["d3"] == "first_run"
    assert d["corrupt"] == 1


def test_save_records_manifest_entry(tmp_path):
    """save_workbook_safe stores plan_hash + fingerprints + region fp (I4)."""
    probe, _ = make_probe(tmp_path)
    out = tmp_path / "out.xlsx"
    proc = run_common(
        "import json, xlsx_common as c\n"
        "probe, out, base = sys.argv[1], sys.argv[2], sys.argv[3]\n"
        "h_ro = c.load_workbook_safe(probe, read_only=True)\n"
        "fp = c.fingerprint_region(h_ro.wb['Urun Listesi'], 'A1:B4')\n"
        "h = c.load_workbook_safe(probe)\n"
        "info = c.save_workbook_safe(h, out, plan={'action': 'test'},\n"
        "                            region_fingerprint=fp,\n"
        "                            manifest_dir=base)\n"
        "r = c.manifest_read(skill_dir=base)\n"
        "entry = r['entries'][-1]\n"
        "print(json.dumps({'saved_fp': fp,\n"
        "                  'stored_fp': entry.get("
        "'target_region_fingerprint'),\n"
        "                  'plan_hash': entry.get('plan_hash'),\n"
        "                  'status': entry.get('status'),\n"
        "                  'op_id': entry.get('operation_id'),\n"
        "                  'manifest_ok': info['manifest']['ok']}))",
        probe, out, tmp_path)
    d = json.loads(proc.stdout)
    assert d["stored_fp"] == d["saved_fp"]
    assert d["plan_hash"].startswith("sha256:")
    assert d["status"] == "committed" and d["op_id"]
    assert d["manifest_ok"] is True


# ---------------------------------------------------------------------------
# H. envelope sweep, validator, recalc branches
# ---------------------------------------------------------------------------

def test_all_scripts_emit_envelope(tmp_path):
    probe, spec_path = make_probe(tmp_path)
    csv_in = tmp_path / "in.csv"
    csv_in.write_text("a,b\n1,2\n", encoding="utf-8")

    payloads = {
        "read": json.loads(run("xlsx_read.py", probe, "--sheets").stdout),
        "create": json.loads(
            run("xlsx_create.py", spec_path, tmp_path / "c2.xlsx").stdout),
        "edit": json.loads(
            run("xlsx_edit.py", probe, "--set", "C1=1",
                "--out", tmp_path / "e2.xlsx").stdout),
        "restructure": json.loads(
            run("xlsx_restructure.py", probe, "--sheet", "Ozet",
                "--insert-rows", "2", "--out", tmp_path / "r2.xlsx").stdout),
        "csv_to_xlsx": json.loads(
            run("csv_to_xlsx.py", csv_in, tmp_path / "x2.xlsx").stdout),
        "xlsx_to_csv": json.loads(
            run("xlsx_to_csv.py", probe, tmp_path / "o.csv").stdout),
    }
    env = dict(os.environ, LC_ALL="C", LANG="C", PATH=str(tmp_path))
    proc = subprocess.run(
        [sys.executable, str(SCRIPTS / "xlsx_recalc.py"), str(probe)],
        capture_output=True, text=True, env=env, encoding="utf-8")
    payloads["recalc"] = json.loads(proc.stdout)

    for name, payload in payloads.items():
        assert payload["ok"] is True, name
        for key in ("warnings", "unsupported", "diagnostics"):
            assert isinstance(payload.get(key), list), f"{name}.{key}"


def test_validate_script_reports_honestly(tmp_path):
    probe, _ = make_probe(tmp_path)
    proc = run("xlsx_validate.py", probe)
    d = json.loads(proc.stdout)
    assert d["ok"] is True
    assert d["formula_text_ok"] is True
    assert d["values_recalculated"] is False
    assert d["formula_text_count"] >= 1
    assert any(c["id"] == "workbook_opens" and c["ok"] for c in d["checks"])

    proc = run("xlsx_validate.py", probe, "--expect-sheets", "Nope",
               expect_ok=False)
    d = json.loads(proc.stdout)
    assert proc.returncode == 1 and d["ok"] is False
    assert any(c["id"] == "sheet_count_stable" and not c["ok"]
               for c in d["checks"])


def test_dependency_error_alias_kept():
    """Old parsers read .error; new contract keeps it as an alias."""
    proc = run("xlsx_read.py", "C:/definitely/not/here.xlsx", "--sheets",
               expect_ok=False)
    d = json.loads(proc.stderr)
    assert d["error"] == d["message"]
    assert d["error_code"] == "FILE_NOT_FOUND"
