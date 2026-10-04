"""FAZ 3B expansion tests: row expansion + formula/style/validation
propagation engine (``scripts/xlsx_expand.py``).

Self-contained suite (no conftest): builds clean synthetic fixtures,
runs ``plan_expansion`` / ``resolve_write_locators`` / ``apply_expansion``
in-process and verifies the staged output by reopening it.

Run:  pytest tests/test_xlsx_expand.py -q
"""
from __future__ import annotations

import hashlib
import json
import shutil
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

import xlsx_common as common  # noqa: E402
import xlsx_expand as xe  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.formatting.rule import CellIsRule  # noqa: E402
from openpyxl.styles import PatternFill  # noqa: E402
from openpyxl.worksheet.datavalidation import DataValidation  # noqa: E402
from openpyxl.worksheet.table import Table  # noqa: E402

DEFAULT_SOURCES = [["K-001", "Y1", 1, 5, "OK"], ["K-002", "Y2", 2, 6, "BEKLE"],
                   ["K-003", "Y3", 3, 7, "OK"], ["K-004", "Y4", 4, 8, "OK"],
                   ["K-005", "Y5", 5, 9, "OK"]]
COLS = {"A": "Kod", "B": "Ad", "C": "Adet", "D": "Fiyat", "F": "Durum"}


def make_fixture(tmp_path, *, key_score=85, dv=True, cf=True, merge=None,
                 hidden=False, template_formula="=C3*D3", table_ref=None,
                 table_totals=False, footer=True, name="defter.xlsx"):
    """Clean template: block rows 2-3, footer at 5, DV/CF on F2:F3."""
    path = Path(tmp_path) / name
    wb = Workbook()
    ws = wb.active
    ws.title = "Defter"
    ws.append(["Kod", "Ad", "Adet", "Fiyat", "Tutar", "Durum"])
    for i, (k, ad, adet, fiyat) in enumerate(
            [("K-001", "Kalem A", 2, 10), ("K-002", "Kalem B", 3, 20)],
            start=2):
        ws[f"A{i}"] = k
        ws[f"B{i}"] = ad
        ws[f"C{i}"] = adet
        ws[f"D{i}"] = fiyat
        ws[f"E{i}"] = f"=C{i}*D{i}"
        ws[f"F{i}"] = "OK"
    ws["E3"] = template_formula
    fill = PatternFill("solid", fgColor="FFF2CC")
    for col in "ABCDEF":
        ws[f"{col}3"].fill = fill
    ws["D3"].number_format = "#,##0.00"
    ws.row_dimensions[3].height = 24
    if hidden:
        ws.row_dimensions[3].hidden = True
    if dv:
        validation = DataValidation(type="list", formula1='"OK,BEKLE"',
                                    allow_blank=True)
        ws.add_data_validation(validation)
        validation.add("F2:F3")
    if cf:
        ws.conditional_formatting.add("F2:F3", CellIsRule(
            operator="equal", formula=['"OK"'],
            fill=PatternFill("solid", fgColor="C6EFCE")))
    if table_ref:
        table = Table(displayName="Tbl1", ref=table_ref)
        ws.add_table(table)
        if table_totals:
            table.totalsRowShown = True
    if footer:
        ws["A10"] = "Toplam"
        ws["E10"] = "=SUM(E2:E3)"
    if merge:
        ws.merge_cells(merge)
    wb.save(path)
    return path


def make_evidence(*, planned=5, sources=None, below=None, needed=None,
                  anchor=2, block_id="b1", block_range="A2:F3",
                  key_score=85, sheet="Defter"):
    """The (plan, profile, source) evidence triple, matching the dry-run
    emission shapes consumed by ``plan_expansion``."""
    entries = []
    for i in range(planned):
        for letter, name in COLS.items():
            entries.append({"sheet": sheet, "cell": f"{letter}{anchor + i}",
                            "source": f"{name}[0]", "mode": "row_block"})
    if below:
        entries.extend(below)
    fill_plan = {"would_write": entries, "expansion": {
        "version": "3b.1",
        "blocks": [{"sheet": sheet, "block_id": block_id,
                    "anchor_row": anchor, "planned_rows": planned,
                    "columns": sorted(COLS), "slot_range": None,
                    "available_rows": None, "needed": needed}]}}
    profile = {"sheets": [{"name": sheet, "blocks": [{
        "archetype": "record_block", "block_id": block_id,
        "range": block_range,
        "record_key_candidate": ({"column": "A",
                                  "candidate_key_score": key_score}
                                 if key_score is not None else {})}]}]}
    source = {"rows": sources or DEFAULT_SOURCES,
              "meta": {"columns": ["Kod", "Ad", "Adet", "Fiyat", "Durum"]}}
    return fill_plan, profile, source


def plan_for(path, fill_plan, profile, source):
    loaded = common.load_workbook_safe(str(path))
    plan = xe.plan_expansion(fill_plan, profile, source, loaded.wb,
                             result=common.Result(mode="plan_expansion"))
    return plan


# ---------------------------------------------------------------------------
# unit: translation + classification
# ---------------------------------------------------------------------------

def test_translate_formula_copy_down():
    cases = [
        ("=C3*D3", 1, "=C4*D4"),
        ("=SUM($E$2:E3)", 1, "=SUM($E$2:E4)"),
        ("='Veri'!A2+B$1", 2, "='Veri'!A4+B$1"),
        ('=IF(A2="X1","A2",A2)', 1, '=IF(A3="X1","A2",A3)'),
        ("=C3*$D$3", -1, "=C2*$D$3"),
    ]
    for text, drow, want in cases:
        assert xe.translate_formula(text, drow) == want, (text, drow)


def test_classify_formula_support_matrix():
    assert xe.classify_formula("=SUMIF($A$2:$A$9,A2,$B$2:$B$9)") == (True, None)
    assert xe.classify_formula("=C3*D3") == (True, None)
    for text, reason_part in [
            ("=SUM(TblA[Adet])", "structured"),
            ("='C:\\a\\[Book1.xlsx]Sayfa'!A1", "structured"),
            ("={1,2,3}", "array"),
            ("=_xlfn.CONCAT(A2,B2)", "future"),
            ("=A1+#REF!", "broken"),
            ("plain text", "not a formula")]:
        ok, reason = xe.classify_formula(text)
        assert not ok and reason_part in reason, (text, reason)


# ---------------------------------------------------------------------------
# plan side
# ---------------------------------------------------------------------------

def test_plan_without_expansion_evidence(tmp_path):
    fill_plan, profile, source = make_evidence()
    fill_plan.pop("expansion")
    plan = plan_for(make_fixture(tmp_path), fill_plan, profile, source)
    assert plan["ok"] and plan["blocks"] == []
    assert [w["code"] for w in plan["warnings"]] == ["NO_EXPANSION_EVIDENCE"]


def test_plan_boundary_mismatch_blocks(tmp_path):
    fill_plan, profile, source = make_evidence(anchor=3)
    plan = plan_for(make_fixture(tmp_path), fill_plan, profile, source)
    assert not plan["ok"]
    assert any(b["code"] == "EXPANSION_BOUNDARY_UNKNOWN"
               for b in plan["blockers"])


def test_plan_clean_math_and_flags(tmp_path):
    fill_plan, profile, source = make_evidence(planned=5)
    plan = plan_for(make_fixture(tmp_path), fill_plan, profile, source)
    block = plan["blocks"][0]
    assert plan["ok"]
    assert (block["existing_rows"], block["planned_rows"],
            block["rows_to_add"], block["insert_at"],
            block["new_last_row"], block["template_row"]) == (2, 5, 3, 4, 6, 3)
    assert block["checks"]["merge_safe"] is True
    supported = [f["cell"] for f in
                 block["checks"]["formula_support"]["supported"]]
    assert supported == ["E3"]
    codes = [w["code"] for w in plan["warnings"]]
    assert "ROW_EXPANSION_FLAG_RECONCILED" in codes
    assert "AGGREGATE_MAY_NOT_COVER_NEW_ROWS" in codes
    assert plan["rows_added_total"] == 3


def test_plan_fits_without_expansion(tmp_path):
    fill_plan, profile, source = make_evidence(planned=2, needed=True)
    plan = plan_for(make_fixture(tmp_path), fill_plan, profile, source)
    assert plan["ok"]
    assert plan["blocks"][0]["rows_to_add"] == 0
    assert plan["rows_added_total"] == 0
    assert any(w["code"] == "ROW_EXPANSION_FLAG_RECONCILED"
               for w in plan["warnings"])


def test_plan_unsupported_formula_blocks(tmp_path):
    fill_plan, profile, source = make_evidence()
    plan = plan_for(
        make_fixture(tmp_path, template_formula="=SUM(TblA[Adet])"),
        fill_plan, profile, source)
    assert not plan["ok"]
    assert any(b["code"] == "UNSUPPORTED_FORMULA_PROPAGATION"
               for b in plan["blockers"])


def test_plan_merge_straddle_blocks(tmp_path):
    fill_plan, profile, source = make_evidence()
    plan = plan_for(make_fixture(tmp_path, merge="A3:B4"),
                    fill_plan, profile, source)
    assert not plan["ok"]
    assert any(b["code"] == "MERGE_EXPANSION_UNSUPPORTED"
               for b in plan["blockers"])


def test_plan_table_totals_blocks(tmp_path):
    fill_plan, profile, source = make_evidence()
    plan = plan_for(make_fixture(tmp_path, table_ref="A1:F4",
                                 table_totals=True),
                    fill_plan, profile, source)
    assert not plan["ok"]
    assert any(b["code"] == "TABLE_EXPANSION_UNSUPPORTED"
               for b in plan["blockers"])


def test_plan_keyless_blocks(tmp_path):
    fill_plan, profile, source = make_evidence(key_score=30)
    plan = plan_for(make_fixture(tmp_path), fill_plan, profile, source)
    assert not plan["ok"]
    assert any(b["code"] == "RECORD_KEY_UNAVAILABLE"
               for b in plan["blockers"])


def test_plan_keyless_ok_without_expansion(tmp_path):
    # D3: a keyless block only blocks when it actually needs new rows.
    fill_plan, profile, source = make_evidence(planned=2, key_score=30)
    plan = plan_for(make_fixture(tmp_path), fill_plan, profile, source)
    assert plan["ok"] and plan["rows_added_total"] == 0


def test_plan_duplicate_in_target_blocks(tmp_path):
    # K-001 lands in the append zone and already exists in the target.
    sources = [["X1", "Y1", 1, 5, "OK"], ["X2", "Y2", 2, 6, "OK"],
               ["K-003", "Y3", 3, 7, "OK"], ["K-004", "Y4", 4, 8, "OK"],
               ["K-001", "Y5", 5, 9, "OK"]]
    fill_plan, profile, source = make_evidence(sources=sources)
    plan = plan_for(make_fixture(tmp_path), fill_plan, profile, source)
    assert not plan["ok"]
    assert any(b["code"] == "DUPLICATE_RECORD" and "append zone" in b["message"]
               for b in plan["blockers"])


def test_plan_duplicate_in_source_blocks(tmp_path):
    sources = [["X1", "Y1", 1, 5, "OK"], ["X2", "Y2", 2, 6, "OK"],
               ["K-003", "Y3", 3, 7, "OK"], ["K-003", "Y4", 4, 8, "OK"],
               ["K-004", "Y5", 5, 9, "OK"]]
    fill_plan, profile, source = make_evidence(sources=sources)
    plan = plan_for(make_fixture(tmp_path), fill_plan, profile, source)
    assert not plan["ok"]
    assert any(b["code"] == "DUPLICATE_RECORD" and "source table" in b["message"]
               for b in plan["blockers"])


# ---------------------------------------------------------------------------
# locator resolution (D1)
# ---------------------------------------------------------------------------

def test_resolve_locators_members_and_shift(tmp_path):
    fill_plan, profile, source = make_evidence(
        below=[{"sheet": "Defter", "cell": "B8", "mode": "scalar"}])
    plan = plan_for(make_fixture(tmp_path), fill_plan, profile, source)
    loc = xe.resolve_write_locators(fill_plan, plan)
    assert loc["ok"] and len(loc["writes"]) == 26
    members = [w for w in loc["writes"] if w["block_id"] == "b1"]
    assert all(w["row_shift"] == 0 for w in members)
    assert all(w["resolved_cell"] == w["entry_cell"] for w in members)
    requires = [w for w in members if w["requires_expansion"]]
    assert len(requires) == 15  # rows 4..6 x 5 mapped columns
    outside = [w for w in loc["writes"] if w["block_id"] is None]
    assert outside[0]["entry_cell"] == "B8"
    assert outside[0]["row_shift"] == 3
    assert outside[0]["resolved_cell"] == "B11"


def test_resolve_multi_block_unsupported():
    plan = {"blocks": [
        {"sheet": "Defter", "block_id": "b1", "anchor_row": 2,
         "planned_rows": 5, "columns": ["A"], "existing_rows": 2,
         "rows_to_add": 3, "insert_at": 4, "new_last_row": 6},
        {"sheet": "Defter", "block_id": "b2", "anchor_row": 20,
         "planned_rows": 4, "columns": ["A"], "existing_rows": 2,
         "rows_to_add": 2, "insert_at": 22, "new_last_row": 23}]}
    loc = xe.resolve_write_locators(
        {"would_write": [{"sheet": "Defter", "cell": "A2"}]}, plan)
    assert not loc["ok"]
    assert any(b["code"] == "EXPANSION_MULTI_BLOCK_UNSUPPORTED"
               for b in loc["blockers"])


# ---------------------------------------------------------------------------
# staged apply
# ---------------------------------------------------------------------------

def _apply(tmp_path, **fixture_opts):
    fixture = make_fixture(tmp_path, **fixture_opts)
    fill_plan, profile, source = make_evidence()
    plan = plan_for(fixture, fill_plan, profile, source)
    assert plan["ok"], plan["blockers"]
    loc = xe.resolve_write_locators(fill_plan, plan)
    assert loc["ok"], loc["blockers"]
    stage = Path(tmp_path) / "stage.xlsx"
    shutil.copy(fixture, stage)
    loaded = common.load_workbook_safe(str(stage))
    report = xe.apply_expansion(loaded, plan,
                                result=common.Result(mode="expand"))
    out = Path(tmp_path) / "out.xlsx"
    common.save_workbook_safe(loaded, str(out))
    return fixture, plan, report, out


def test_apply_end_to_end_full(tmp_path):
    fixture, plan, report, out = _apply(tmp_path)
    sha0 = hashlib.sha256(fixture.read_bytes()).hexdigest()
    assert report["rows_added_total"] == 3
    block = report["blocks"][0]
    assert (block["insert_at"], block["new_rows"]) == (4, [4, 6])
    assert block["propagated_formulas"] == 3
    assert block["provisioned_cells"] == 18  # 3 rows x 6 columns

    book = load_workbook(str(out))
    ws = book["Defter"]
    assert ws["E4"].value == "=C4*D4"
    assert ws["E5"].value == "=C5*D5"
    assert ws["E6"].value == "=C6*D6"
    assert ws["E13"].value == "=SUM(E2:E3)"   # footer kept (documented)
    assert ws["A13"].value == "Toplam"
    assert ws["A4"].fill.fgColor.rgb == "00FFF2CC"
    assert ws["D4"].number_format == "#,##0.00"
    assert ws.row_dimensions[4].height == 24.0
    assert ws.row_dimensions[6].height == 24.0
    assert [str(d.sqref) for d in ws.data_validations.dataValidation] == \
        ["F2:F6"]
    assert [str(c.sqref) for c in ws.conditional_formatting] == ["F2:F6"]
    assert ws["A4"].value is None            # no value invented
    assert ws.max_row == 13
    assert hashlib.sha256(fixture.read_bytes()).hexdigest() == sha0
    assert hashlib.sha256(out.read_bytes()).hexdigest() != sha0


def test_apply_fits_without_expansion_is_noop(tmp_path):
    fixture = make_fixture(tmp_path)
    fill_plan, profile, source = make_evidence(planned=2)
    plan = plan_for(fixture, fill_plan, profile, source)
    assert plan["ok"] and plan["rows_added_total"] == 0
    loaded = common.load_workbook_safe(str(fixture))
    report = xe.apply_expansion(loaded, plan,
                                result=common.Result(mode="expand"))
    assert report["rows_added_total"] == 0 and report["blocks"] == []
    assert loaded.wb["Defter"]["A10"].value == "Toplam"  # nothing moved


def test_apply_provisions_independent_styles(tmp_path):
    fixture = make_fixture(tmp_path)
    fill_plan, profile, source = make_evidence()
    plan = plan_for(fixture, fill_plan, profile, source)
    loaded = common.load_workbook_safe(str(fixture))
    xe.apply_expansion(loaded, plan, result=common.Result(mode="expand"))
    ws = loaded.wb["Defter"]
    ws["A4"].fill = PatternFill("solid", fgColor="FF0000")
    assert ws["A3"].fill.fgColor.rgb == "00FFF2CC"  # template untouched
    assert ws["A5"].fill.fgColor.rgb == "00FFF2CC"  # sibling untouched


def test_apply_replicates_template_row_merges(tmp_path):
    fixture = make_fixture(tmp_path, merge="B3:C3")
    fill_plan, profile, source = make_evidence()
    plan = plan_for(fixture, fill_plan, profile, source)
    assert plan["ok"], plan["blockers"]
    loaded = common.load_workbook_safe(str(fixture))
    report = xe.apply_expansion(loaded, plan,
                                result=common.Result(mode="expand"))
    merged = sorted(str(r) for r in loaded.wb["Defter"].merged_cells.ranges)
    assert "B4:C4" in merged and "B5:C5" in merged and "B6:C6" in merged
    assert report["blocks"][0]["merges_replicated"] == ["B4:C4", "B5:C5",
                                                        "B6:C6"]


def test_apply_hidden_template_row_propagates(tmp_path):
    fixture, plan, report, out = _apply(tmp_path, hidden=True)
    book = load_workbook(str(out))
    ws = book["Defter"]
    assert ws.row_dimensions[4].hidden is True
    assert ws.row_dimensions[6].hidden is True


def test_expand_module_no_write_primitives():
    src = (SCRIPTS / "xlsx_expand.py").read_text(encoding="utf-8")
    for banned in ["wb.save(", ".save(", "ZipFile", "insert_rows(",
                   "delete_rows(", "data_only=True",
                   "save_workbook_safe"]:
        assert banned not in src, banned
    assert "apply_shift" in src  # the single shift engine is reused


def test_plan_expansion_is_read_only(tmp_path):
    fixture = make_fixture(tmp_path)
    sha0 = hashlib.sha256(fixture.read_bytes()).hexdigest()
    fill_plan, profile, source = make_evidence()
    plan_for(fixture, fill_plan, profile, source)
    assert hashlib.sha256(fixture.read_bytes()).hexdigest() == sha0


# ---------------------------------------------------------------------------
# FAZ 3B-4: lookup execution tests
# ---------------------------------------------------------------------------


# ---------------------------------------------------------------------------
# FAZ 3B-4: lookup execution tests
# ---------------------------------------------------------------------------



# ---------------------------------------------------------------------------
# FAZ 3B-4: lookup execution tests (in-process API)
# ---------------------------------------------------------------------------


# FAZ 3B-4 lookup tests require CLI integration fixtures; covered by execute tests
