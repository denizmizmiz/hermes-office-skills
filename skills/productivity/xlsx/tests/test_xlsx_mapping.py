# -*- coding: utf-8 -*-
"""FAZ 2B -- xlsx_mapping.py plan + dry-run tests (no-write by design)."""
import hashlib
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

from test_xlsx_semantics import (MAPPER, SEM, make_book, parse, run_map,
                                run_sem, sha)

SKILL = SEM.parent.parent
LABELS_DIR = Path(__file__).resolve().parent / "fixtures" / "labels"


def make_source(path, rows, headers=("Ad", "Adet", "Birim Fiyat", "Tutar",
                                     "Notlar"), delimiter=";"):
    if path.suffix == ".json":
        payload = [dict(zip(headers, row)) for row in rows]
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2),
                        encoding="utf-8")
        return path
    lines = [delimiter.join(headers)]
    for row in rows:
        lines.append(delimiter.join("" if cell is None else str(cell)
                                    for cell in row))
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


ROWS = [
    ("Kalem X", 5, "1250,50", None, "ilk not"),
    ("Kalem Y", 3, "999,90", None, "ikinci not"),
    ("Kalem Z", 7, "100,00", None, "ucuncu not"),
    ("Kalem W", 2, "50,00", None, "dorduncu not"),
    ("Kalem V", 9, "10,00", None, "besinci not"),
]


@pytest.fixture(scope="module")
def workspace(tmp_path_factory):
    directory = tmp_path_factory.mktemp("map2b")
    books = {name: make_book(directory / f"{name}.xlsx", variant=name)
             for name in ("clean", "lookup")}
    profiles = {}
    for name, book in books.items():
        out = directory / f"{name}.profile.json"
        run_sem(book, "--emit", "profile", "--out", out)
        profiles[name] = out
    csv = make_source(directory / "src.csv", ROWS)
    plan = directory / "plan.json"
    run_map("--profile", profiles["clean"], "--source", csv, "--emit", "plan",
            "--out", plan)
    return {"dir": directory, "books": books, "profiles": profiles,
            "csv": csv, "plan": plan}


def profile_of(workspace, name="clean"):
    return workspace["profiles"][name]


# ---------------------------------------------------------------------------
# B5 -- mapping plan
# ---------------------------------------------------------------------------

def test_plan_contract(workspace):
    doc = parse(run_map("--profile", profile_of(workspace),
                        "--source", workspace["csv"], "--emit", "plan"))
    assert doc["ok"] is True and doc["command"] == "plan"
    for key in ("warnings", "unsupported", "diagnostics", "plan"):
        assert key in doc
    plan = doc["plan"]
    for key in ("plan_version", "plan_id", "profile_id",
                "source_content_fingerprint", "target_structure_fingerprint",
                "mappings", "unresolved", "conflicts", "policy", "complete"):
        assert key in plan, key
    assert plan["policy"]["on_low_confidence"] == "report_only"
    for mapping in plan["mappings"]:
        for key in ("source_key", "target", "mode", "confidence", "level",
                    "evidence", "ambiguity", "preconditions"):
            assert key in mapping, key
        assert mapping["evidence"]


def test_plan_is_complete_for_a_matching_source(workspace):
    plan = parse(run_map("--profile", profile_of(workspace),
                         "--source", workspace["csv"], "--emit", "plan"))["plan"]
    keys = {m["source_key"] for m in plan["mappings"]}
    assert {"Ad", "Adet", "Birim Fiyat", "Notlar"} <= keys
    notlar = next(m for m in plan["mappings"] if m["source_key"] == "Notlar")
    assert notlar["mode"] == "row_block"
    assert notlar["target"]["anchor"].startswith("E")
    assert all(m["confidence"] >= 30 for m in plan["mappings"])
    assert plan["summary"]["unresolved"] == len(plan["unresolved"])


def test_plan_never_invents_unmatched_source(workspace):
    directory = workspace["dir"]
    weird = make_source(directory / "weird.csv", ROWS,
                        headers=("Ad", "Adet", "Birim Fiyat", "Tutar",
                                 "Zzz Qqq"))
    plan = parse(run_map("--profile", profile_of(workspace), "--source", weird,
                         "--emit", "plan"))["plan"]
    unresolved = [u.get("source_key") for u in plan["unresolved"]]
    assert "Zzz Qqq" in unresolved
    assert all(m["source_key"] != "Zzz Qqq" for m in plan["mappings"])


def test_plan_type_mismatch_is_a_conflict(workspace):
    directory = workspace["dir"]
    text_source = make_source(directory / "text_tutar.csv",
                             [("A", 1, "1,0", "metin", "n")], 
                             headers=("Ad", "Adet", "Birim Fiyat", "Tutar",
                                      "Notlar"))
    plan = parse(run_map("--profile", profile_of(workspace),
                         "--source", text_source, "--emit", "plan"))["plan"]
    kinds = {c["kind"] for c in plan["conflicts"]}
    assert "value_kind_mismatch" in kinds
    assert plan["complete"] is False


def test_plan_determinism_and_identity(workspace):
    first = run_map("--profile", profile_of(workspace),
                    "--source", workspace["csv"], "--emit", "plan")
    second = run_map("--profile", profile_of(workspace),
                     "--source", workspace["csv"], "--emit", "plan")
    assert first.stdout == second.stdout
    plan_id = parse(first)["plan"]["plan_id"]
    left = workspace["dir"] / "left"
    right = workspace["dir"] / "right"
    left.mkdir(exist_ok=True)
    right.mkdir(exist_ok=True)
    a = make_source(left / "same.csv", ROWS)
    b = make_source(right / "same.csv", ROWS)
    os.utime(a, (1_600_000_000, 1_600_000_000))
    os.utime(b, (1_700_000_000, 1_700_000_000))
    plan_a = parse(run_map("--profile", profile_of(workspace), "--source", a,
                           "--emit", "plan"))["plan"]
    plan_b = parse(run_map("--profile", profile_of(workspace), "--source", b,
                           "--emit", "plan"))["plan"]
    assert plan_a["plan_id"] == plan_b["plan_id"] == plan_id
    assert plan_a["source"]["content_fingerprint"] == \
        plan_b["source"]["content_fingerprint"]


def test_plan_json_source_supported(workspace):
    directory = workspace["dir"]
    js = make_source(directory / "src.json", ROWS)
    plan = parse(run_map("--profile", profile_of(workspace), "--source", js,
                         "--emit", "plan"))["plan"]
    assert plan["mappings"]
    assert plan["source"]["kind"] == "json"


def test_plan_policy_validation(workspace, tmp_path):
    bad = tmp_path / "policy.json"
    bad.write_text('{"nope": 1}', encoding="utf-8")
    proc = run_map("--profile", profile_of(workspace),
                   "--source", workspace["csv"], "--policy-file", bad,
                   "--emit", "plan", expect_ok=False)
    assert json.loads(proc.stderr)["error_code"] == "SPEC_UNKNOWN_KEY"
    bad.write_text('{"on_conflict": "explode"}', encoding="utf-8")
    proc = run_map("--profile", profile_of(workspace),
                   "--source", workspace["csv"], "--policy-file", bad,
                   "--emit", "plan", expect_ok=False)
    assert json.loads(proc.stderr)["error_code"] == "SPEC_INVALID"


# ---------------------------------------------------------------------------
# B6 -- dry-run
# ---------------------------------------------------------------------------

def dry_run(workspace, book="clean", source=None, plan=None):
    return parse(run_map("--profile", profile_of(workspace, book),
                         "--plan", plan or workspace["plan"],
                         "--source", source or workspace["csv"],
                         "--emit", "dry-run",
                         "--target", workspace["books"][book]))


def test_dry_run_contract_and_partition(workspace):
    doc = dry_run(workspace)
    assert doc["command"] == "dry-run"
    fill = doc["fill_plan"]
    for key in ("dry_run_version", "plan_id", "would_write", "blocked",
                "preserved", "unresolved", "unsupported", "row_expansion",
                "stats", "policy"):
        assert key in fill, key
    for entry in fill["would_write"]:
        for key in ("sheet", "cell", "source", "preview", "type", "checks"):
            assert key in entry, key
        assert set(("validation", "merged", "format_compatible",
                    "role_compatible")) <= set(entry["checks"])
    classified = (len(fill["would_write"]) + len(fill["blocked"]) +
                  len(fill["preserved"]))
    assert classified >= 1
    assert fill["stats"]["cells"] == len(fill["would_write"])
    assert "actual_rows_added" not in ()


def test_dry_run_preserves_formula_cells(workspace):
    fill = dry_run(workspace)["fill_plan"]
    assert fill["preserved"], "formula target column must be preserved"
    for entry in fill["preserved"]:
        assert entry["write_policy"] == "preserve_formula"
        assert entry["cell"].startswith("D")
    assert fill["stats"]["preserved_formula_cells"] == len(fill["preserved"])


def test_dry_run_writes_only_into_evidenced_anchors(workspace):
    fill = dry_run(workspace)["fill_plan"]
    cells = {entry["cell"] for entry in fill["would_write"]}
    assert cells, "expected previewable writes"
    rows = {int("".join(ch for ch in cell if ch.isdigit())) for cell in cells}
    assert min(rows) >= 2


def test_row_expansion_is_planned_only(workspace):
    fill = dry_run(workspace)["fill_plan"]
    expansion = fill["row_expansion"]
    assert expansion["status"] == "planned_only"
    assert expansion["actual_rows_added"] == 0
    assert expansion["needed"] is True
    assert expansion["requested_rows"] > (expansion["available_rows"] or 0)


def test_dry_run_blocks_empty_required_values(workspace):
    directory = workspace["dir"]
    partial = make_source(directory / "partial.csv",
                          [("Kalem X", 5, "1,0", None, "")],
                          headers=("Ad", "Adet", "Birim Fiyat", "Tutar",
                                   "Notlar"))
    doc = dry_run(workspace, source=partial)
    reasons = " ".join(entry.get("reason", "") for entry in doc["fill_plan"]["blocked"])
    assert "not invented" in reasons or "empty" in reasons


def test_dry_run_reports_lookup_validation_as_unsupported(workspace):
    profile = profile_of(workspace, "lookup")
    doc = parse(run_map("--profile", profile, "--plan", workspace["plan"],
                        "--source", workspace["csv"], "--emit", "dry-run",
                        "--target", workspace["books"]["lookup"]))
    messages = " ".join(item["message"] for item in doc["unsupported"])
    assert "lookup/list" in messages or "reference-style" in messages


def test_dry_run_determinism(workspace):
    first = run_map("--profile", profile_of(workspace), "--plan",
                    workspace["plan"], "--source", workspace["csv"],
                    "--emit", "dry-run", "--target",
                    workspace["books"]["clean"])
    second = run_map("--profile", profile_of(workspace), "--plan",
                     workspace["plan"], "--source", workspace["csv"],
                     "--emit", "dry-run", "--target",
                     workspace["books"]["clean"])
    assert first.stdout == second.stdout


def test_dry_run_requires_plan_and_source(workspace):
    proc = run_map("--profile", profile_of(workspace),
                   "--source", workspace["csv"], "--emit", "dry-run",
                   expect_ok=False)
    assert json.loads(proc.stderr)["error_code"] == "SPEC_INVALID"
    proc = run_map("--profile", profile_of(workspace), "--plan",
                   workspace["plan"], "--emit", "dry-run", expect_ok=False)
    assert json.loads(proc.stderr)["error_code"] == "SPEC_INVALID"


def test_dry_run_never_touches_the_target(workspace):
    before = {name: sha(path) for name, path in workspace["books"].items()}
    dry_run(workspace)
    after = {name: sha(path) for name, path in workspace["books"].items()}
    assert before == after


def test_mapping_module_has_no_write_calls():
    text = MAPPER.read_text(encoding="utf-8")
    assert "import openpyxl" not in text
    import ast  # noqa: PLC0415
    tree = ast.parse(text)
    for node in ast.walk(tree):
        if isinstance(node, ast.Call):
            func = node.func
            name = getattr(func, "attr", None) or getattr(func, "id", None)
            assert name not in ("save", "save_workbook_safe", "ZipFile"), \
                f"forbidden call {name} at line {node.lineno}"


def test_profile_from_doc_map_file_works_for_mapping(workspace, tmp_path):
    doc_map = tmp_path / "map.json"
    proc = subprocess.run([sys.executable, str(SEM.parent / "xlsx_understand.py"),
                           str(workspace["books"]["clean"])],
                          capture_output=True, text=True, encoding="utf-8",
                          timeout=300)
    doc_map.write_text(proc.stdout, encoding="utf-8")
    plan = parse(run_map("--profile", profile_of(workspace),
                         "--source", workspace["csv"], "--emit", "plan",
                         "--doc-map", doc_map))["plan"]
    assert plan["target_structure_fingerprint"]
    assert plan["mappings"]
