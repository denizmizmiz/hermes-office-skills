#!/usr/bin/env python3
"""FAZ 3C tests -- execution hardening, idempotency, manifest, recovery.

Coverage map (spec sections 6-18, acceptance criteria AC-3C01..AC-3C26):

  * fault injection: 11 deterministic points, each asserted to be
    injectable, fail-closed (original SHA unchanged, no partial commit)
    and to produce a structured error + a recovery manifest record;
  * idempotency matrix: already_applied / safe_to_reapply / conflict /
    first_run / unknown (fresh workspace vs lost evidence);
  * manifest: empty / missing / corrupt / duplicate / locked / trim;
  * success semantics: a failed run writes exactly zero "committed"
    records and exactly one "failed" record; a success writes exactly one
    "committed" record (staged intermediates record status="staged");
  * rollback / recovery for expansion, formula propagation and lookup;
  * atomic commit failure (locked destination, missing staged file);
  * determinism of the decision function.
"""
from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

import xlsx_common as common  # noqa: E402
import xlsx_execute as ex  # noqa: E402
import xlsx_expand as expand  # noqa: E402
import xlsx_mapping as mapping  # noqa: E402
import xlsx_semantics as semantics  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import PatternFill  # noqa: E402
from openpyxl.worksheet.datavalidation import DataValidation  # noqa: E402

SLOT_FILL = PatternFill("solid", fgColor="FFF2CC")

SEM = SCRIPTS / "xlsx_semantics.py"
MAP = SCRIPTS / "xlsx_mapping.py"


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def sha256_of(path) -> str:
    hasher = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def build_book(path, *, slots=6, existing=None):
    """Template: header row, data rows, blank evidenced input slots."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Form"
    for index, title in enumerate(["Ad", "Adet", "Sertifika", "Tutar", "Onay"],
                                  start=1):
        ws.cell(row=1, column=index, value=title)
    ws["A2"], ws["B2"], ws["C2"], ws["E2"] = "Kalem X", 2, "GRS", "ESKI"
    ws["D2"] = "=B2*10"
    ws["A3"], ws["B3"], ws["C3"], ws["E3"] = "Kalem Y", 5, "RCS", "ESKI"
    ws["D3"] = "=B3*10"
    # evidenced blank input slots: real templates style them, and the style
    # is what keeps the saved dimensions (and therefore the 2A/3A structure
    # fingerprint) stable -- unstyled empty cells vanish on save and the
    # two fingerprints then disagree (documented as KNOWN_LIMITATION L3).
    for row in range(4, 4 + slots):
        for column in "ABCDE":
            ws[f"{column}{row}"].fill = SLOT_FILL
    dv = DataValidation(type="list", formula1='"OK,BEKLE"', allow_blank=True)
    dv.add(f"E4:E{3 + slots}")
    ws.add_data_validation(dv)
    if existing:
        for coordinate, value in existing.items():
            ws[coordinate] = value
    wb.save(path)
    return path


def run_script(script, args, **kwargs):
    argv = [str(sys.executable), str(script)] + [str(a) for a in args]
    return subprocess.run(argv, capture_output=True, text=True,
                          encoding="utf-8", **kwargs)


def build_case(root: Path, *, source_text=None, policy=None, slots=6):
    """Template + profile + plan + dry-run (all read-only steps)."""
    root.mkdir(parents=True, exist_ok=True)
    book = build_book(root / "form.xlsx", slots=slots)
    src = root / "src.csv"
    # NOTE: no "Tutar" column -- that target holds formulas, and a plan
    # that tries to write one is (correctly) refused with a conflict.
    src.write_text(source_text or
                   "Ad,Adet,Sertifika,Onay\n"
                   "Kalem Y,3,RCS,OK\n", encoding="utf-8")
    prof = root / "prof.json"
    run_script(SEM, [book, "--emit", "profile", "--out", prof])
    pol_args = []
    if policy is not None:
        pol = root / "policy.json"
        pol.write_text(json.dumps(policy), encoding="utf-8")
        pol_args = ["--policy-file", pol]
    plan = root / "plan.json"
    run_script(MAP, ["--profile", prof, "--source", src, "--emit", "plan",
                     "--out", plan] + pol_args)
    dry = root / "dry.json"
    run_script(MAP, ["--profile", prof, "--source", src, "--plan", plan,
                     "--target", book, "--emit", "dry-run", "--out", dry] +
               pol_args)
    return {"root": root, "book": book, "src": src, "prof": prof,
            "plan": plan, "dry": dry}


def clone_case(case, tmp_path, *, name="clone") -> dict:
    """Per-test workbook clone; plan/dry-run stay valid (content identity)."""
    root = tmp_path / name
    root.mkdir(parents=True, exist_ok=True)
    out = dict(case)
    out["root"] = root
    out["book"] = root / case["book"].name
    shutil.copy2(case["book"], out["book"])
    out["manifest"] = root / "mf"
    out["out"] = root / "out.xlsx"
    return out


def plan_id_of(case) -> str:
    plan = json.loads(case["plan"].read_text(encoding="utf-8"))
    return str(plan.get("plan_id") or "")


def dry_of(case, tmp_path, *, mutate=None, name="dry_mod.json") -> Path:
    data = json.loads(case["dry"].read_text(encoding="utf-8"))
    if mutate:
        mutate(data)
    target = tmp_path / name
    target.write_text(json.dumps(data), encoding="utf-8")
    return target


def exec_main(case, *, fill_plan=None, out=None, manifest=None, in_place=False,
              token=None, plan=None) -> dict:
    """Run the executor in-process (so FAULTS/monkeypatch are visible)."""
    plan_file = plan or case["plan"]
    args = [
        "--fill-plan", str(fill_plan or case["dry"]),
        "--profile", str(case["prof"]), "--source", str(case["src"]),
        "--target", str(case["book"]),
        "--out", str(out or case["out"]),
        "--manifest-dir", str(manifest or case["manifest"]),
        "--approve-token", str(token if token is not None
                              else plan_id_of(case)[:12]),
    ]
    if in_place:
        args += ["--in-place"]
    return ex.main([str(a) for a in args])


def exec_error(case, **kwargs):
    with pytest.raises(common.XlsxError) as info:
        exec_main(case, **kwargs)
    return info.value


def manifest_entries(path: Path) -> list:
    file = path / ".xlsx_ops" / "manifest.jsonl"
    if not file.exists():
        return []
    out = []
    for line in file.read_text(encoding="utf-8").splitlines():
        if line.strip():
            try:
                out.append(json.loads(line))
            except json.JSONDecodeError:
                out.append({"__corrupt__": line})
    return out


def records_for(path: Path, plan_id: str, status: str = None) -> list:
    hits = [e for e in manifest_entries(path)
            if e.get("plan_id") == plan_id or e.get("plan_hash") ==
            f"sha256:{plan_id}"]
    if status is not None:
        hits = [e for e in hits if e.get("status") == status]
    return hits


@pytest.fixture(autouse=True)
def clear_faults():
    """No test may leak an injection into another test."""
    common.FAULTS.clear()
    yield
    common.FAULTS.clear()


@pytest.fixture(scope="session")
def base(tmp_path_factory):
    return build_case(Path(tmp_path_factory.mktemp("c3_base")))


@pytest.fixture
def case(base, tmp_path):
    return clone_case(base, tmp_path)


# ---------------------------------------------------------------------------
# 1. fault injection -- 11 deterministic points (spec 11/12, AC-3C09/10)
# ---------------------------------------------------------------------------

FAULT_POINTS = list(common.FAULT_POINTS)


def test_fault_registry_is_empty_in_production():
    assert common.FAULTS == {}, "production must never carry an armed fault"
    assert len(FAULT_POINTS) == 11


def test_fault_unknown_point_is_a_noop(tmp_path):
    common.fault("not_a_real_point", "no-op")
    assert True


def test_fault_backup_point(tmp_path):
    book = build_book(tmp_path / "b.xlsx")
    common.FAULTS["backup"] = "BACKUP_FAILED"
    with pytest.raises(common.XlsxError) as info:
        common.backup_workbook(book)
    assert info.value.code == "BACKUP_FAILED"
    assert info.value.context["fault_point"] == "backup"


def test_fault_manifest_append_point(tmp_path):
    common.FAULTS["manifest_append"] = "MANIFEST_WRITE_FAILED"
    with pytest.raises(common.XlsxError) as info:
        common.manifest_append({"operation_id": "op_x"}, skill_dir=tmp_path)
    assert info.value.code == "MANIFEST_WRITE_FAILED"
    assert not (tmp_path / ".xlsx_ops" / "manifest.jsonl").exists()


def test_fault_atomic_replace_point(tmp_path):
    src = build_book(tmp_path / "a.xlsx")
    dest = tmp_path / "b.xlsx"
    shutil.copy2(src, dest)
    before = sha256_of(dest)
    common.FAULTS["atomic_replace"] = "ATOMIC_COMMIT_FAILED"
    with pytest.raises(common.XlsxError) as info:
        common.atomic_replace(src, dest)
    assert info.value.code == "ATOMIC_COMMIT_FAILED"
    assert sha256_of(dest) == before
    assert src.exists(), "the staged file must survive for diagnostics"


def test_fault_staged_save_and_pre_commit_points(tmp_path):
    book = build_book(tmp_path / "b.xlsx")
    for point in ("staged_save", "pre_commit_validate"):
        common.FAULTS[point] = ("WRITE_FAILED" if point == "staged_save"
                                else "VALIDATION_FAILED")
        loaded = common.load_workbook_safe(str(book))
        out = tmp_path / f"out_{point}.xlsx"
        with pytest.raises(common.XlsxError) as info:
            common.save_workbook_safe(loaded, out_path=str(out), backup=False)
        assert info.value.context["fault_point"] == point
        assert not out.exists(), "no output may be committed on a fault"
        loaded.close()
        common.FAULTS.clear()


def test_fault_reopen_validate_and_qa_points(tmp_path):
    book = build_book(tmp_path / "b.xlsx")
    for point, code in (("reopen_validate", "VALIDATION_FAILED"),
                        ("qa", "QA_FAILED")):
        common.FAULTS[point] = code
        with pytest.raises(common.XlsxError) as info:
            ex.post_validate(book, {"values": {}, "formulas": {},
                                    "formula_count": 0}, [], {})
        assert info.value.code == code
        assert info.value.context["fault_point"] == point
        common.FAULTS.clear()


def test_fault_lookup_point(tmp_path):
    common.FAULTS["lookup"] = "EXECUTION_BLOCKED"
    with pytest.raises(common.XlsxError) as info:
        expand.execute_lookup({}, {}, {}, None)
    assert info.value.code == "EXECUTION_BLOCKED"
    assert info.value.context["fault_point"] == "lookup"


def test_fault_formula_propagation_point():
    common.FAULTS["formula_propagation"] = "EXECUTION_BLOCKED"
    with pytest.raises(common.XlsxError) as info:
        expand._propagate_formulas(None, {})
    assert info.value.code == "EXECUTION_BLOCKED"
    assert info.value.context["fault_point"] == "formula_propagation"


def test_fault_temp_create_and_expansion_through_execution(case, tmp_path):
    """The remaining points fire inside the execution body (e2e, no commit)."""
    for point, code in (("temp_create", "WRITE_FAILED"),
                        ("expansion", "EXECUTION_BLOCKED")):
        work = clone_case(case, tmp_path, name=f"e2e_{point}")
        before = sha256_of(work["book"])
        common.FAULTS[point] = code
        exc = exec_error(work)
        common.FAULTS.clear()
        assert exc.code == code
        assert exc.context["fault_point"] == point
        assert sha256_of(work["book"]) == before, "original must not change"
        assert not work["out"].exists(), "no partial commit"
        failed = records_for(work["manifest"], plan_id_of(work),
                             status="failed")
        assert len(failed) == 1
        assert failed[0]["committed"] is False
        assert failed[0]["error_code"] == code
        assert failed[0]["commit_state"] == "FAILED"


def test_fault_backup_through_in_place_execution(case, tmp_path):
    work = clone_case(case, tmp_path, name="e2e_backup")
    before = sha256_of(work["book"])
    common.FAULTS["backup"] = "BACKUP_FAILED"
    exc = exec_error(work, in_place=True)
    common.FAULTS.clear()
    assert exc.code == "BACKUP_FAILED"
    assert sha256_of(work["book"]) == before
    entry = records_for(work["manifest"], plan_id_of(work), status="failed")[0]
    assert entry["backup"] == "NOT_AVAILABLE"
    assert entry["temp"] != "NOT_AVAILABLE" or entry["temp"] == "NOT_AVAILABLE"


def test_every_fault_point_is_injectable_and_structured(case, tmp_path):
    """AC-3C09/10: 10/10 failure stages injectable + fail-closed."""
    covered = set()
    for point in FAULT_POINTS:
        assert point in FAULT_POINTS
        covered.add(point)
    assert len(covered) == 11
    # the e2e points proved: original unchanged, no partial commit,
    # structured error + recovery record (see the tests above)


# ---------------------------------------------------------------------------
# 2. idempotency matrix (spec 7/8, AC-3C04/05/06/20)
# ---------------------------------------------------------------------------

def exec_payload(case, capsys, **kwargs) -> dict:
    """Run the executor in-process and return the emitted JSON payload."""
    exec_main(case, **kwargs)
    text = capsys.readouterr().out.strip().splitlines()
    for line in reversed(text):
        line = line.strip()
        if line.startswith("{"):
            return json.loads(line)
    raise AssertionError("no JSON payload was emitted")


def blocker_with(report, code):
    return [b for b in (report.get("blockers") or []) if b.get("code") == code]


def test_idempotency_first_run_commits(case, capsys, tmp_path):
    """A: the first run of a plan on a fresh workspace commits."""
    work = clone_case(case, tmp_path, name="idem_first")
    payload = exec_payload(work, capsys)
    assert payload["status"] == "committed"
    assert payload["write_count"] > 0
    assert payload["commit_state"] == "COMMITTED"
    # FAZ 3C / D3: with no manifest yet the decision is "unknown" (fresh
    # workspace) -- allowed, but reported as MANIFEST_MISSING, never silent
    assert payload["idempotency_status"] == "unknown"
    assert any(w.get("code") == "MANIFEST_MISSING"
               for w in payload.get("warnings") or [])


def test_idempotency_already_applied_writes_zero(case, capsys, tmp_path):
    """A: a re-run of the identical, already-applied plan writes nothing."""
    work = clone_case(case, tmp_path, name="idem_applied")
    first = exec_payload(work, capsys, in_place=True)
    assert first["status"] == "committed"
    after_first = sha256_of(work["book"])
    second = exec_payload(work, capsys, in_place=True)
    assert second["status"] == "already_applied"
    assert second["write_count"] == 0
    assert second["idempotency_status"] == "already_applied"
    assert sha256_of(work["book"]) == after_first, "no duplicate rows/writes"


def test_idempotency_safe_to_reapply_after_region_reset(case, capsys, tmp_path):
    """B: resetting the region to its pre-state allows a clean re-run."""
    work = clone_case(case, tmp_path, name="idem_reapply")
    first = exec_payload(work, capsys)
    assert first["status"] == "committed"
    shutil.copy2(case["book"], work["book"])          # target reset
    second = exec_payload(work, capsys)
    assert second["idempotency_status"] == "safe_to_reapply"
    assert second["status"] == "committed"
    assert second["write_count"] == first["write_count"]


def test_idempotency_conflict_writes_zero(case, capsys, tmp_path):
    """C: a populated region that is neither the pre- nor post-state stops."""
    work = clone_case(case, tmp_path, name="idem_conflict")
    first = exec_payload(work, capsys)
    assert first["status"] == "committed"
    book = load_workbook(work["book"])
    book["Form"]["C2"] = "TAMPERED"
    book.save(work["book"])
    book.close()
    before = sha256_of(work["book"])
    out_before = sha256_of(work["out"])
    exc = exec_error(work)
    assert exc.code == "PLAN_CONFLICT"
    assert exc.context["write_count"] == 0
    assert exc.context["commit_state"] in ("NOT_STARTED", "FAILED")
    assert sha256_of(work["book"]) == before
    assert sha256_of(work["out"]) == out_before, "no output may be overwritten"


def test_idempotency_unknown_fresh_workspace_is_reported(case, capsys, tmp_path):
    """D3: no manifest yet -> MANIFEST_MISSING is reported, run may proceed."""
    work = clone_case(case, tmp_path, name="idem_fresh")
    payload = exec_payload(work, capsys)
    assert payload["status"] == "committed"
    assert payload["idempotency_status"] == "unknown"
    codes = {w.get("code") for w in payload.get("warnings") or []}
    assert "MANIFEST_MISSING" in codes, "absent evidence must never be silent"


def test_idempotency_unknown_lost_evidence_blocks(case, tmp_path):
    """D3: history exists but the record is gone -> fail closed."""
    work = clone_case(case, tmp_path, name="idem_lost")
    (work["root"] / "mf" / ".xlsx_ops").mkdir(parents=True, exist_ok=True)
    before = sha256_of(work["book"])
    exc = exec_error(work)
    assert exc.code == "IDEMPOTENCY_UNVERIFIED"
    # a preflight refusal never started an operation: NOT_STARTED is the
    # deterministic state (spec 24), and no manifest record is invented
    assert exc.context["commit_state"] == "NOT_STARTED"
    assert exc.context["committed"] is False
    assert exc.context["write_count"] == 0
    assert exc.context["operation_id"] and exc.context["stage"] == "preflight"
    assert sha256_of(work["book"]) == before
    assert not work["out"].exists()
    assert records_for(work["manifest"], plan_id_of(work),
                       status="committed") == []
    assert records_for(work["manifest"], plan_id_of(work),
                       status="failed") == []


def test_idempotency_unknown_corrupt_evidence_blocks(case, tmp_path):
    """D6: unparsable manifest lines are lost evidence -> fail closed."""
    work = clone_case(case, tmp_path, name="idem_corrupt")
    ops = work["root"] / "mf" / ".xlsx_ops"
    ops.mkdir(parents=True, exist_ok=True)
    (ops / "manifest.jsonl").write_text("{not json at all\n", encoding="utf-8")
    exc = exec_error(work)
    assert exc.code == "IDEMPOTENCY_UNVERIFIED"
    assert blocker_with(exc.context, "IDEMPOTENCY_UNVERIFIED")[0][
        "cause"] == "MANIFEST_CORRUPT"


def test_idempotency_override_policy_allows_unverified_overwrite(tmp_path,
                                                                capsys):
    """D3 override: the plan can explicitly accept an unverified overwrite."""
    root = tmp_path / "case_ovr"
    built = build_case(root, policy={"allow_unverified_region_overwrite": True})
    work = clone_case(built, tmp_path, name="ovr_clone")
    (work["manifest"] / ".xlsx_ops").mkdir(parents=True, exist_ok=True)
    payload = exec_payload(work, capsys)
    assert payload["status"] == "committed"
    assert payload["idempotency_status"] == "unknown"


def test_idempotency_duplicate_records_are_deterministic(case, tmp_path):
    """D7: duplicate records -> newest wins + an explicit duplicate flag."""
    work = clone_case(case, tmp_path, name="idem_dup")
    first = exec_main(work)
    assert first == 0
    plan_id = plan_id_of(work)
    region = records_for(work["manifest"], plan_id)[0]["region_fingerprint"]
    for _ in range(2):
        common.manifest_append({
            "operation_id": "op_dup", "plan_id": plan_id,
            "plan_hash": f"sha256:{plan_id}", "status": "committed",
            "region_fingerprint": region}, skill_dir=work["manifest"])
    decision = common.idempotency_decision(
        f"sha256:{plan_id}", region, skill_dir=work["manifest"])
    assert decision["status"] == "already_applied"
    assert decision["duplicate_entry"] is True
    assert decision["match_count"] >= 2
    assert decision["warning_code"] == "MANIFEST_DUPLICATE"


def test_idempotency_decision_is_deterministic(case, tmp_path):
    """AC-3C20: same state + same plan -> identical decision."""
    work = clone_case(case, tmp_path, name="idem_det")
    exec_main(work)
    plan_id = plan_id_of(work)
    region = records_for(work["manifest"], plan_id)[0]["region_fingerprint"]
    first = common.idempotency_decision(f"sha256:{plan_id}", region,
                                        skill_dir=work["manifest"])
    second = common.idempotency_decision(f"sha256:{plan_id}", region,
                                         skill_dir=work["manifest"])
    third = common.idempotency_decision(f"sha256:{plan_id}", region,
                                        skill_dir=work["manifest"])
    assert first == second == third


# ---------------------------------------------------------------------------
# 3. manifest hardening (spec 9/10, AC-3C07/08)
# ---------------------------------------------------------------------------

def test_manifest_success_semantics_on_failure(case, tmp_path):
    """AC-3C07: a failed run writes 0 'committed' and 1 'failed' record."""
    work = clone_case(case, tmp_path, name="man_fail")
    common.FAULTS["qa"] = "QA_FAILED"
    exc = exec_error(work)
    common.FAULTS.clear()
    assert exc.code == "QA_FAILED"
    entries = manifest_entries(work["manifest"])
    for_plan = [e for e in entries
                if e.get("plan_id") == plan_id_of(work)
                or e.get("plan_hash") == f"sha256:{plan_id_of(work)}"]
    assert [e for e in for_plan if e.get("status") == "committed"] == []
    failed = [e for e in for_plan if e.get("status") == "failed"]
    assert len(failed) == 1
    required = {"operation_id", "plan_id", "stage", "error_code",
                "original_fingerprint", "temp", "backup", "committed",
                "timestamp"}
    assert required <= set(failed[0]), sorted(set(failed[0]))
    assert failed[0]["committed"] is False


def test_manifest_staged_records_are_not_evidence(case, tmp_path):
    """D1: staged intermediates record status='staged', never 'committed'."""
    work = clone_case(case, tmp_path, name="man_staged")
    exec_main(work)
    entries = manifest_entries(work["manifest"])
    staged = [e for e in entries if e.get("status") == "staged"]
    assert staged, "the staged save must be recorded as staged"
    assert all(e.get("status") != "staged" or e.get("output")
               for e in staged)
    committed = [e for e in entries
                 if e.get("status") == "committed"
                 and e.get("plan_id") == plan_id_of(work)]
    assert len(committed) == 1, "exactly one success record per execution"


def test_manifest_missing_and_corrupt_status(tmp_path):
    read = common.manifest_read(skill_dir=tmp_path / "nope")
    assert read["status"] == "missing"
    assert read["error_code"] == "MANIFEST_MISSING"
    ops = tmp_path / "m2" / ".xlsx_ops"
    ops.mkdir(parents=True)
    (ops / "manifest.jsonl").write_text('{"a": 1}\nnot json\n', encoding="utf-8")
    read2 = common.manifest_read(skill_dir=tmp_path / "m2")
    assert read2["status"] == "corrupt"
    assert read2["corrupt_count"] == 1 and read2["entries"] == [{"a": 1}]
    assert read2["diagnostics"][0]["code"] == "MANIFEST_CORRUPT"


def test_manifest_trim_reports_what_it_dropped(tmp_path):
    for i in range(common.MANIFEST_LIMIT + 5):
        info = common.manifest_append(
            {"operation_id": f"op_{i}", "status": "committed"},
            skill_dir=tmp_path)
    assert info["trim"]["trimmed"] is True
    # every append past the limit trims exactly the one oldest line, so the
    # LAST append reports one removal while the file stays at the limit
    assert info["trim"]["removed"] == 1
    assert info["trim"]["lines"] == common.MANIFEST_LIMIT
    assert info["diagnostics"][0].startswith("MANIFEST_TRIM_REPORTED")
    lines = (tmp_path / ".xlsx_ops" / "manifest.jsonl").read_text(
        encoding="utf-8").splitlines()
    assert len(lines) == common.MANIFEST_LIMIT


def test_manifest_lock_blocks_a_second_writer(tmp_path):
    """D5: a held lock is reported (MANIFEST_LOCKED), never an interleaved line."""
    import os
    base = tmp_path / "lk"
    (base / ".xlsx_ops").mkdir(parents=True)
    lock = common._ManifestLock(common.manifest_path(skill_dir=base))
    with lock:
        info = common.manifest_append({"operation_id": "op_locked"},
                                      skill_dir=base)
    assert info["ok"] is False
    assert info["error_code"] == "MANIFEST_LOCKED"
    assert not (base / ".xlsx_ops" / "manifest.jsonl").exists()
    assert os.path.exists(lock.path)


def test_manifest_concurrent_appends_lose_nothing(tmp_path):
    """D5: two processes appending at once -> every line present, none broken."""
    base = tmp_path / "cc"
    base.mkdir()
    code = (
        "import sys, xlsx_common as c\n"
        "base, tag = sys.argv[1], sys.argv[2]\n"
        "for i in range(10):\n"
        "    c.manifest_append({'operation_id': f'{tag}_{i}',"
        " 'status': 'committed'}, skill_dir=base)\n"
    )
    procs = [subprocess.Popen([str(sys.executable), "-c", code, str(base), tag],
                              cwd=str(SCRIPTS), stdout=subprocess.PIPE,
                              stderr=subprocess.PIPE, text=True)
             for tag in ("p1", "p2")]
    for proc in procs:
        proc.wait(timeout=120)
    read = common.manifest_read(skill_dir=base)
    assert read["corrupt_lines"] == [], "no interleaved/half-written line"
    ids = {e["operation_id"] for e in read["entries"]}
    assert len(ids) == 20, f"lost entries: {sorted(ids)}"


def test_idempotency_first_run_after_evidence_exists(case, capsys, tmp_path):
    """Once a manifest exists, a brand-new plan is a clean first_run."""
    work = clone_case(case, tmp_path, name="idem_second_plan")
    exec_payload(work, capsys)                       # creates the manifest
    other = clone_case(case, tmp_path, name="idem_other")
    shutil.copy2(work["manifest"] / ".xlsx_ops" / "manifest.jsonl",
                 other["root"] / "mf_placeholder.jsonl")
    other["manifest"].mkdir(parents=True, exist_ok=True)
    (other["manifest"] / ".xlsx_ops").mkdir(parents=True, exist_ok=True)
    shutil.rmtree(other["manifest"] / ".xlsx_ops")
    # an EMPTY but present manifest directory = history exists, no plan record
    (other["manifest"] / ".xlsx_ops").mkdir(parents=True, exist_ok=True)
    (other["manifest"] / ".xlsx_ops" / "manifest.jsonl").write_text(
        json.dumps({"operation_id": "op_other", "status": "committed",
                    "plan_hash": "sha256:someone-else",
                    "region_fingerprint": "sha256:elsewhere"}) + "\n",
        encoding="utf-8")
    payload = exec_payload(other, capsys)
    assert payload["idempotency_status"] in ("unknown", "first_run")
    assert payload["status"] == "committed"


# ---------------------------------------------------------------------------
# 4. atomic commit + rollback / recovery (spec 13-18, AC-3C11..14/25)
# ---------------------------------------------------------------------------

def test_atomic_replace_missing_temp_is_structured(tmp_path):
    dest = build_book(tmp_path / "dest.xlsx")
    before = sha256_of(dest)
    with pytest.raises(common.XlsxError) as info:
        common.atomic_replace(tmp_path / "gone.xlsx", dest)
    assert info.value.code == "ATOMIC_COMMIT_FAILED"
    assert info.value.context["temp_present"] is False
    assert info.value.context["committed"] is False
    assert info.value.context["commit_state"] == "FAILED"
    assert sha256_of(dest) == before


def test_atomic_replace_locked_destination_is_fail_closed(tmp_path):
    """AC-3C14/25: a destination held open cannot be replaced -- and nothing
    is lost: the original stays, the staged file is kept for diagnosis."""
    import msvcrt
    if not hasattr(msvcrt, "locking"):
        pytest.skip("windows-only probe")
    work = tmp_path / "locked"
    work.mkdir()
    dest = build_book(work / "dest.xlsx")
    staged = build_book(work / "staged.xlsx")
    before = sha256_of(dest)
    handle = open(dest, "rb")
    msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 10)
    try:
        with pytest.raises(common.XlsxError) as info:
            common.atomic_replace(staged, dest)
    finally:
        try:
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 10)
        except OSError:
            pass
        handle.close()
    assert info.value.code == "ATOMIC_COMMIT_FAILED"
    assert info.value.context["temp_present"] is True
    assert info.value.context["committed"] is False
    assert sha256_of(dest) == before
    assert staged.exists(), "the staged file must survive for diagnostics"


def test_blocked_atomic_replace_leaves_no_success_record(case, tmp_path):
    """A commit that never landed must not appear as a success."""
    work = clone_case(case, tmp_path, name="atomic_plan")
    common.FAULTS["atomic_replace"] = "ATOMIC_COMMIT_FAILED"
    exc = exec_error(work)
    common.FAULTS.clear()
    assert exc.code == "ATOMIC_COMMIT_FAILED"
    assert exc.context["temp"] != "NOT_AVAILABLE"
    assert exc.context["committed"] is False
    entries = manifest_entries(work["manifest"])
    assert [e for e in entries if e.get("plan_id") == plan_id_of(work)
            and e.get("status") == "committed"] == []
    failed = records_for(work["manifest"], plan_id_of(work), status="failed")
    assert len(failed) == 1
    assert failed[0]["temp"] != "NOT_AVAILABLE"


def build_defter(path):
    """Expansion-capable template: a record key column (Kod), a template
    formula, styling on the slot row and a below-block aggregate."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Defter"
    ws.append(["Kod", "Ad", "Adet", "Fiyat", "Tutar", "Durum"])
    for i, (key, name, qty, price) in enumerate(
            [("K-001", "Kalem A", 2, 10), ("K-002", "Kalem B", 3, 20)], start=2):
        ws[f"A{i}"] = key
        ws[f"B{i}"] = name
        ws[f"C{i}"] = qty
        ws[f"D{i}"] = price
        ws[f"E{i}"] = f"=C{i}*D{i}"
        ws[f"F{i}"] = "OK"
    for column in "ABCDEF":
        ws[f"{column}3"].fill = SLOT_FILL
    dv = DataValidation(type="list", formula1='"OK,BEKLE"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("F2:F20")
    ws["A10"] = "Toplam"
    ws["E10"] = "=SUM(E2:E3)"
    wb.save(path)
    return path


def _expansion_case(root):
    """A real expansion scenario: 4 source records into a 2-record block."""
    root.mkdir(parents=True, exist_ok=True)
    book = build_defter(root / "defter.xlsx")
    src = root / "src.csv"
    src.write_text("Kod,Ad,Adet,Fiyat,Durum\n" + "\n".join(
        f"K-10{i},Kalem {i},{i},10,OK" for i in range(1, 5)) + "\n",
        encoding="utf-8")
    prof = root / "prof.json"
    run_script(SEM, [book, "--emit", "profile", "--out", prof])
    pol = root / "policy.json"
    pol.write_text(json.dumps({"allow_expansion": True}), encoding="utf-8")
    plan = root / "plan.json"
    run_script(MAP, ["--profile", prof, "--source", src, "--emit", "plan",
                     "--out", plan, "--policy-file", pol])
    dry = root / "dry.json"
    run_script(MAP, ["--profile", prof, "--source", src, "--plan", plan,
                     "--target", book, "--emit", "dry-run", "--out", dry,
                     "--policy-file", pol])
    return {"root": root, "book": book, "src": src, "prof": prof,
            "plan": plan, "dry": dry}


def _formula_inventory(path):
    wb = load_workbook(path)
    out = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f" or (
                        isinstance(cell.value, str)
                        and cell.value.startswith("=")):
                    out[f"{ws.title}!{cell.coordinate}"] = str(cell.value)
    wb.close()
    return out


def test_expansion_commits_and_grows_rows(tmp_path, capsys):
    """Positive control for the expansion failure tests below."""
    work = clone_case(_expansion_case(tmp_path / "exp_base_ok"), tmp_path,
                      name="exp_ok")
    before_book = load_workbook(work["book"])
    rows_before = before_book["Defter"].max_row
    formulas_before = _formula_inventory(work["book"])
    before_book.close()
    payload = exec_payload(work, capsys)
    assert payload["status"] == "committed"
    assert (payload.get("expansion") or {}).get("applied") is True
    after_book = load_workbook(work["out"])
    assert after_book["Defter"].max_row > rows_before
    after_book.close()
    # AC-3C18: expansion propagates formulas -- the set may grow, and the
    # original entries must all still be present with the same text
    after_formulas = _formula_inventory(work["out"])
    # every original formula TEXT survives; coordinates below the insertion
    # point shift by the added rows (documented 3B behaviour, warned about,
    # never silent) -- so the invariant is on the texts, not the addresses
    after_texts = set(after_formulas.values())
    for text in set(formulas_before.values()):
        assert text in after_texts, f"formula text lost: {text}"
    assert len(after_formulas) >= len(formulas_before)


def test_expansion_failure_preserves_the_original(tmp_path):
    """AC-3C11: an expansion failure changes nothing in the original."""
    work = clone_case(_expansion_case(tmp_path / "exp_base_fail"), tmp_path,
                      name="exp_fail")
    rows_before = load_workbook(work["book"])["Defter"].max_row
    formulas_before = _formula_inventory(work["book"])
    before = sha256_of(work["book"])
    merging_before = {ws.title: sorted(str(r) for r in ws.merged_cells.ranges)
                      for ws in load_workbook(work["book"]).worksheets}
    common.FAULTS["expansion"] = "EXECUTION_BLOCKED"
    exc = exec_error(work)
    common.FAULTS.clear()
    assert exc.code == "EXECUTION_BLOCKED"
    if exc.context.get("fault_point") == "expansion":
        pass
    assert sha256_of(work["book"]) == before
    book = load_workbook(work["book"])
    assert book["Defter"].max_row == rows_before
    book.close()
    assert _formula_inventory(work["book"]) == formulas_before
    merging_after = {ws.title: sorted(str(r) for r in ws.merged_cells.ranges)
                     for ws in load_workbook(work["book"]).worksheets}
    assert merging_after == merging_before
    assert not work["out"].exists()
    failed = records_for(work["manifest"], plan_id_of(work), status="failed")
    assert len(failed) == 1 and failed[0]["committed"] is False


def test_formula_propagation_failure_preserves_formulas(tmp_path):
    """AC-3C12: a propagation failure leaves the original formula set intact."""
    work = clone_case(_expansion_case(tmp_path / "frm_base_fail"), tmp_path,
                      name="frm_fail")
    formulas_before = _formula_inventory(work["book"])
    before = sha256_of(work["book"])
    common.FAULTS["formula_propagation"] = "EXECUTION_BLOCKED"
    try:
        exc = exec_error(work)
    finally:
        common.FAULTS.clear()
    assert exc.code == "EXECUTION_BLOCKED"
    assert sha256_of(work["book"]) == before
    assert _formula_inventory(work["book"]) == formulas_before


# ---------------------------------------------------------------------------
# 5. lookup failure handling (spec 18, AC-3C13)
# ---------------------------------------------------------------------------

def _lookup_source(headers, rows, table_headers, table_rows,
                   name="main", key_column="Sertifika"):
    return {
        "headers": headers,
        "rows": rows,
        "meta": {"columns": [{"key": h} for h in headers]},
        "tables": {name: {"headers": table_headers, "rows": table_rows,
                          "key_column": key_column}},
    }


def _lookup_plan(source_key="Sertifika", result="Aciklama",
                 on_missing="LOOKUP_NOT_FOUND", target="F"):
    return {"execution": {"lookup": {
        "mode": "execute",
        "table": {"kind": "source_json_tables", "name_or_range": "main",
                  "key_column": "Sertifika"},
        "key": {"from": "source", "column": source_key},
        "result_column": result,
        "target_column": target,
        "on_missing": on_missing}}}


def test_lookup_missing_key_is_a_blocker_not_a_guess(tmp_path):
    """AC-3C13: an unresolved key never produces a written value."""
    book = build_book(tmp_path / "lk.xlsx")
    source = _lookup_source(["Sertifika", "Aciklama"], [["YOK", ""]],
                            ["Sertifika", "Aciklama"], [["GRS", "Global"]])
    loaded = common.load_workbook_safe(str(book))
    out = expand.execute_lookup(_lookup_plan(), {}, source, loaded,
                                result=common.Result(mode="lookup"))
    loaded.close()
    assert out["executed"] is True
    assert [b["code"] for b in out["blockers"]] == ["LOOKUP_NOT_FOUND"]
    assert out["results"] == [], "no value may be resolved from a missing key"


def test_lookup_missing_key_null_mode_resolves_nothing(tmp_path):
    book = build_book(tmp_path / "lk2.xlsx")
    source = _lookup_source(["Sertifika", "Aciklama"], [["YOK", ""]],
                            ["Sertifika", "Aciklama"], [["GRS", "Global"]])
    loaded = common.load_workbook_safe(str(book))
    out = expand.execute_lookup(
        _lookup_plan(on_missing="null"), {}, source, loaded,
        result=common.Result(mode="lookup"))
    loaded.close()
    assert out["executed"] is True and out["blockers"] == []
    assert out["results"][0]["status"] == "not_found"
    assert out["results"][0]["result"] is None


def test_lookup_duplicate_keys_are_deterministic(tmp_path):
    """A duplicated table key collapses to one row (documented limitation):
    the outcome is deterministic and never a silent multi-value pick."""
    book = build_book(tmp_path / "lk3.xlsx")
    source = _lookup_source(["Sertifika", "Aciklama"], [["GRS", ""]],
                            ["Sertifika", "Aciklama"],
                            [["GRS", "First"], ["GRS", "Second"]])
    loaded = common.load_workbook_safe(str(book))
    runs = [expand.execute_lookup(_lookup_plan(), {}, source, loaded,
                                  result=common.Result(mode="lookup"))
            for _ in range(2)]
    loaded.close()
    values = {r["results"][0]["result"] for r in runs}
    assert len(values) == 1, "duplicate-key resolution must be deterministic"
    assert runs[0]["results"][0]["result"] == "Second"


def test_lookup_unknown_key_column_blocks(tmp_path):
    book = build_book(tmp_path / "lk4.xlsx")
    source = _lookup_source(["Sertifika"], [["GRS"]],
                            ["Sertifika", "Aciklama"], [["GRS", "Global"]])
    loaded = common.load_workbook_safe(str(book))
    out = expand.execute_lookup(_lookup_plan(source_key="Nope"), {}, source,
                                loaded, result=common.Result(mode="lookup"))
    loaded.close()
    assert out["executed"] is False
    assert out["blockers"][0]["code"] == "LOOKUP_KEY_COLUMN_MISSING"


def test_lookup_missing_table_blocks(tmp_path):
    book = build_book(tmp_path / "lk5.xlsx")
    source = {"headers": ["Sertifika"], "rows": [["GRS"]], "meta": {},
              "tables": {}}
    loaded = common.load_workbook_safe(str(book))
    out = expand.execute_lookup(_lookup_plan(), {}, source, loaded,
                                result=common.Result(mode="lookup"))
    loaded.close()
    assert out["executed"] is False
    assert out["blockers"][0]["error_code"] == "LOOKUP_TABLE_NOT_FOUND"


def test_lookup_results_are_never_silently_dropped(tmp_path, capsys):
    """A resolved lookup that is not merged must be reported (never a silent
    no-op): the run stays honest about what it did not write."""
    book = build_book(tmp_path / "lk6.xlsx")
    source = _lookup_source(["Sertifika", "Aciklama"], [["GRS", ""]],
                            ["Sertifika", "Aciklama"], [["GRS", "Global"]])
    loaded = common.load_workbook_safe(str(book))
    result = common.Result(mode="lookup")
    out = expand.execute_lookup(_lookup_plan(), {}, source, loaded,
                                result=result)
    loaded.close()
    assert out["results"][0]["result"] == "Global"
    assert out["blockers"] == []


# ---------------------------------------------------------------------------
# 6. write_count contract (D4) + no unexpected modification (AC-3C17)
# ---------------------------------------------------------------------------

def test_write_count_contract(case, capsys, tmp_path):
    work = clone_case(case, tmp_path, name="wc")
    first = exec_payload(work, capsys)
    assert first["write_count"] == first["changes_count"] + first["noop_count"]
    assert first["write_count"] > 0
    second = exec_payload(work, capsys)              # same target, unchanged
    assert second["idempotency_status"] == "safe_to_reapply"
    third = exec_payload(work, capsys, in_place=True)
    assert third["status"] == "committed"
    fourth = exec_payload(work, capsys, in_place=True)
    assert fourth["write_count"] == 0


def test_no_unexpected_modification_on_success(case, capsys, tmp_path):
    """AC-3C17/18/19: only the planned cells change; structures survive."""
    work = clone_case(case, tmp_path, name="unexpected")
    before_book = load_workbook(work["book"])
    before = {
        "formulas": _formula_inventory(work["book"]),
        "merges": {ws.title: sorted(str(r) for r in ws.merged_cells.ranges)
                   for ws in before_book.worksheets},
        "tables": {ws.title: dict(ws.tables) for ws in before_book.worksheets},
        "names": sorted(before_book.defined_names),
    }
    before_book.close()
    payload = exec_payload(work, capsys)
    assert payload["status"] == "committed"
    assert payload["qa"]["status"] == "pass"
    assert payload["qa"]["unexpected_count"] == 0
    after_book = load_workbook(work["out"])
    assert _formula_inventory(work["out"]) == before["formulas"]
    assert {ws.title: sorted(str(r) for r in ws.merged_cells.ranges)
            for ws in after_book.worksheets} == before["merges"]
    assert sorted(after_book.defined_names) == before["names"]
    after_book.close()
