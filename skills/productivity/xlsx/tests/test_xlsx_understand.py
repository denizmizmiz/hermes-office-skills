"""Phase 2A tests: xlsx_understand.py Document Map.

Self-contained: builds its own fixtures and runs the CLI as a subprocess
under LC_ALL=C to prove explicit UTF-8 I/O (same convention as
test_xlsx_skill.py). No network access. Every analysis also re-hashes
the fixture to prove the no-write guarantee.
"""
from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

import xlsx_common as common  # noqa: E402


def run_understand(path, *args, expect_ok=True):
    env = dict(os.environ, LC_ALL="C", LANG="C")
    env.pop("PYTHONIOENCODING", None)
    proc = subprocess.run(
        [sys.executable, str(SCRIPTS / "xlsx_understand.py"), str(path),
         *map(str, args)],
        capture_output=True, text=True, env=env, encoding="utf-8")
    if expect_ok:
        assert proc.returncode == 0, f"xlsx_understand failed: {proc.stderr}"
    return proc


def parse(proc):
    return json.loads(proc.stdout)


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def sheet(doc, name):
    return next(s for s in doc["sheets"] if s["name"] == name)


def colmap(s):
    return {c["letter"]: c for c in s["columns"]}


@pytest.fixture(scope="session")
def fx(tmp_path_factory):
    d = tmp_path_factory.mktemp("understand_fx")
    paths = {}

    wb = Workbook(); ws = wb.active; ws.title = "Satis"
    ws.append(["Ad", "Adet", "Birim Fiyat", "Tutar"])
    for i, (a, b_, c) in enumerate(
            [("Kalem A", 3, 12.5), ("Kalem B", 5, 8.0), ("Kalem C", 2, 45.75),
             ("Kalem D", 7, 3.2), ("Kalem E", 1, 99.9)], start=2):
        ws.append([a, b_, c, f"=B{i}*C{i}"])
        ws.cell(row=i, column=3).number_format = '"₺" #,##0.00'
    paths["f01"] = d / "f01.xlsx"; wb.save(paths["f01"])

    wb = Workbook(); ws = wb.active; ws.title = "Satis"
    ws["A1"] = "Ürün"; ws.merge_cells("A1:A2")
    ws["B1"] = "Satışlar"; ws.merge_cells("B1:C1")
    ws["B2"] = "Q1"; ws["C2"] = "Q2"
    for i in range(3, 8):
        ws.append(["Ürün %d" % (i - 2), i * 10, i * 12])
    paths["f02"] = d / "f02.xlsx"; wb.save(paths["f02"])

    wb = Workbook(); ws = wb.active; ws.title = "Stok"
    ws.append(["Ürün", "Adet", "Tutar"])
    for i in range(2, 5):
        ws.append(["Ürün %d" % i, i, i * 5])
    t = Table(displayName="TabloStok", ref="A1:C4")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9")
    ws.add_table(t)
    paths["f03"] = d / "f03.xlsx"; wb.save(paths["f03"])

    wb = Workbook(); ws = wb.active; ws.title = "Rapor"
    ws["A1"] = "BAŞLIK"; ws["A1"].font = Font(bold=True)
    ws.merge_cells("A1:D1")
    ws.append(["Ürün", "Adet", "Fiyat", "Not"])
    for i in range(3, 7):
        ws.append(["Ürün %d" % i, i, i * 2.5, "n%d" % i])
    ws.row_dimensions[3].hidden = True
    ws.column_dimensions["D"].hidden = True
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A2:D6"
    paths["f04"] = d / "f04.xlsx"; wb.save(paths["f04"])

    wb = Workbook(); ws = wb.active; ws.title = "Tarih"
    ws.append(["Tarih", "Onay", "Aciklama"])
    for i, dt in enumerate([datetime(2025, 1, 15), datetime(2025, 3, 2),
                            datetime(2025, 12, 31)], start=2):
        ws.append([dt, i % 2 == 0, "kayit %d" % i])
        ws.cell(row=i, column=1).number_format = "dd.mm.yyyy"
    paths["f06"] = d / "f06.xlsx"; wb.save(paths["f06"])

    wb = Workbook(); ws = wb.active; ws.title = "Butce"
    ws.append(["BÜTÇE RAPORU"])
    ws.append(["DÖNEM 1"])
    for i in range(3, 6):
        ws.append(["Kalem %d" % i, i * 100])
    ws.append(["Ara Toplam", "=SUM(B3:B5)"])
    ws.append(["DÖNEM 2"])
    for i in range(7, 9):
        ws.append(["Kalem %d" % i, i * 90])
    ws.append(["GENEL TOPLAM", "=SUM(B3:B5)+SUM(B7:B8)"])
    ws.append(["Not: tutarlar TL'dir"])
    paths["f07"] = d / "f07.xlsx"; wb.save(paths["f07"])

    wb = Workbook(); ws2 = wb.create_sheet("Veri")
    for i in range(1, 6):
        ws2.append([i, i * 10])
    ws = wb.active; ws.title = "Ozet"
    ws["A1"] = "Toplam"; ws["B1"] = "=SUM(Veri!A1:A5)"
    ws["A2"] = "Tek"; ws["B2"] = "='Veri'!B2"
    wb.defined_names.add(DefinedName("ToplamAlan", attr_text="Veri!$A$1:$A$5"))
    paths["f08"] = d / "f08.xlsx"; wb.save(paths["f08"])

    wb = Workbook(); ws = wb.active; ws.title = "Kontroller"
    ws.append(["Sinif", "Puan", "Durum"])
    for i in range(2, 8):
        ws.append(["A", i * 7, "x"])
    dv = DataValidation(type="list", formula1='"A,B,C"', allow_blank=True)
    ws.add_data_validation(dv); dv.add("A2:A10")
    ws.conditional_formatting.add(
        "B2:B10", CellIsRule(operator="greaterThan", formula=["30"]))
    ws.conditional_formatting.add(
        "B2:B10", ColorScaleRule(start_type="min", start_color="FF0000",
                                 end_type="max", end_color="00FF00"))
    paths["f09"] = d / "f09.xlsx"; wb.save(paths["f09"])

    wb = Workbook(); ws = wb.active; ws.title = "Kalemler"
    ws.append(["Proje", "Adet", "Tutar"])
    for r in [("P1", 1, 10), ("P2", 2, 20), ("P3", 3, 30),
              ("Ara Toplam", None, 60), ("P4", 4, 40), ("P5", 5, 50),
              ("P6", 6, 60), ("Ara Toplam", None, 150)]:
        ws.append(list(r))
    paths["f11"] = d / "f11.xlsx"; wb.save(paths["f11"])

    wb = Workbook(); ws = wb.active; ws.title = "Ham"
    for i in range(1, 9):
        ws.append([i, i * 2, (i if i <= 4 else "metin %d" % i)])
    paths["f12"] = d / "f12.xlsx"; wb.save(paths["f12"])

    wb = Workbook(); ws = wb.active; ws.title = "Plan"
    ws["A1"] = "FİRMA A.Ş. 2026 SATIŞ PLANI"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Hazırlayan: X Birimi"
    ws.append([]); ws.append([])
    ws.append(["Ürün", "Hedef"])
    for i in range(2, 7):
        ws.append(["Ürün %d" % i, i * 1000])
    paths["f13"] = d / "f13.xlsx"; wb.save(paths["f13"])

    wb = Workbook(); ws = wb.active; ws.title = "Form"
    ws.append(["Ad Soyad", "Tarih", "Imza"])
    for i in range(2, 5):
        ws.append(["kayit %d" % i, datetime(2026, 1, i), "x"])
    fill = PatternFill("solid", fgColor="FFF2CC")
    for r in range(7, 10):
        for c in range(1, 4):
            ws.cell(row=r, column=c).fill = fill
    ws.cell(row=16, column=8).fill = fill
    paths["f14"] = d / "f14.xlsx"; wb.save(paths["f14"])

    wb = Workbook(); ws = wb.active; ws.title = "Num"
    for i in range(1, 13):
        ws.append(["K%03d" % i])
    paths["f17"] = d / "f17.xlsx"; wb.save(paths["f17"])

    wb = Workbook(); ws = wb.active; ws.title = "Özet Çalışma"
    ws.append(["İSTANBUL", "ığüşöçİ"])
    ws.append(["ŞİRKET", "DEĞER"])
    paths["f18"] = d / "f18.xlsx"; wb.save(paths["f18"])

    from openpyxl.chart import BarChart, Reference
    wb = Workbook(); ws = wb.active; ws.title = "Grafik"
    ws.append(["Ay", "Satis"])
    for i in range(1, 5):
        ws.append(["Ay %d" % i, i * 10])
    ch = BarChart()
    ch.add_data(Reference(ws, min_col=2, min_row=1, max_row=5))
    ws.add_chart(ch, "E2")
    paths["chart"] = d / "chart.xlsx"; wb.save(paths["chart"])
    return paths


def test_f01_simple_table(fx):
    p = fx["f01"]
    before = sha256(p)
    doc = parse(run_understand(p))
    assert sha256(p) == before, "analysis must never write"
    assert doc["ok"] is True
    s = sheet(doc, "Satis")
    h = s["headers"][0]
    assert h["confidence_level"] == "HIGH"
    assert [l["raw"] for l in h["labels"] if l["raw"]] == [
        "Ad", "Adet", "Birim Fiyat", "Tutar"]
    cm = colmap(s)
    assert cm["B"]["data_type"] == "integer"
    assert cm["C"]["data_type"] == "decimal"
    assert "₺" in cm["C"]["number_format"]
    assert cm["D"]["formula_count"] == 5
    assert {"header", "data"} <= {r["type"] for r in s["regions"]}
    f = {x["cell"]: x for x in s["formulas"]}
    assert f["D2"]["kind"] == "relative_arithmetic"
    assert f["D2"]["refs"], "refs must be parsed"
    assert all(not x["cached_value_available"] for x in s["formulas"])
    assert doc["workbook"]["cached_values"]["values_recalculated"] is False


def test_f02_multi_row_header(fx):
    doc = parse(run_understand(fx["f02"]))
    s = sheet(doc, "Satis")
    h = s["headers"][0]
    assert h["multi_row"] is True and h["rows"] == [1, 2]
    assert h["merged"] is True
    assert "subheader" in {r["type"] for r in s["regions"]}


def test_f03_excel_table(fx):
    doc = parse(run_understand(fx["f03"]))
    t = sheet(doc, "Stok")["tables"][0]
    assert t["name"] == "TabloStok" and t["ref"] == "A1:C4"
    assert [c["name"] for c in t["columns"]] == ["Ürün", "Adet", "Tutar"]
    assert t["row_count"] == 3


def test_f04_structure(fx):
    doc = parse(run_understand(fx["f04"]))
    s = sheet(doc, "Rapor")
    assert [m["range"] for m in s["merged_cells"]] == ["A1:D1"]
    assert [r["index"] for r in s["hidden"]["rows"]] == [3]
    assert [c["index"] for c in s["hidden"]["columns"]] == ["D"]
    assert s["freeze_panes"] == "A2"
    assert s["autofilter"] == "A2:D6"
    types = {r["type"] for r in s["regions"]}
    assert "title" in types
    assert s["headers"][0]["rows"] == [2]


def test_f06_dates_bools(fx):
    doc = parse(run_understand(fx["f06"]))
    cm = colmap(sheet(doc, "Tarih"))
    assert cm["A"]["data_type"] == "datetime"
    assert cm["B"]["data_type"] == "boolean"
    assert cm["A"]["number_format"] == "dd.mm.yyyy"


def test_f07_regions(fx):
    doc = parse(run_understand(fx["f07"]))
    s = sheet(doc, "Butce")
    types = {r["type"] for r in s["regions"]}
    assert {"section_header", "subtotal", "total", "notes"} <= types
    assert len(s["formulas"]) == 2


def test_f08_cross_sheet(fx):
    doc = parse(run_understand(fx["f08"]))
    s = sheet(doc, "Ozet")
    f = {x["cell"]: x for x in s["formulas"]}
    assert f["B1"]["sheet_refs"] == ["Veri"]
    assert f["B2"]["has_cross_sheet"] is True
    assert f["B1"]["kind"] == "function"
    assert [n["name"] for n in doc["workbook"]["defined_names"]] == [
        "ToplamAlan"]
    ext = doc["workbook"]["external_links_status"]
    assert isinstance(ext.get("available"), bool)


def test_f09_dv_and_cf(fx):
    doc = parse(run_understand(fx["f09"]))
    s = sheet(doc, "Kontroller")
    assert s["validations"][0]["type"] == "list"
    kinds = {c["kind"] for c in s["conditional_formats"]}
    assert {"cellIs", "colorScale"} <= kinds


def test_f11_repeated_structures(fx):
    doc = parse(run_understand(fx["f11"]))
    s = sheet(doc, "Kalemler")
    assert s["repeated_structures"], "repeated rows must be detected"
    for entry in s["repeated_structures"]:
        assert entry["occurrences"] >= 2
        assert entry["type"] in ("repeated_header", "repeated_row_block",
                                 "repeated_periodic_pattern")
        assert entry["confidence_level"] in ("HIGH", "MEDIUM", "LOW")


def test_f12_no_header_mixed(fx):
    doc = parse(run_understand(fx["f12"]))
    s = sheet(doc, "Ham")
    assert colmap(s)["C"]["data_type"] == "mixed"
    assert not s["headers"] or s["headers"][0]["confidence_level"] != "HIGH"


def test_f13_title_block(fx):
    doc = parse(run_understand(fx["f13"]))
    s = sheet(doc, "Plan")
    assert "title" in {r["type"] for r in s["regions"]}
    assert any(h["rows"][0] >= 4 for h in s["headers"])


def test_f14_styled_empty_and_inflation(fx):
    doc = parse(run_understand(fx["f14"]))
    s = sheet(doc, "Form")
    assert s["dimensions"]["inflation_rows"] >= 7
    assert s["dimensions"]["styled_empty_rows"] == 4
    ranges = {r["range"] for r in s["regions"] if r["type"] == "input_region"}
    assert "A7:C9" in ranges
    assert doc["diagnostics"], "dimension inflation must be diagnosed"


def test_f17_sample_cap(fx):
    doc = parse(run_understand(fx["f17"], "--samples", "3"))
    cm = colmap(sheet(doc, "Num"))
    assert len(cm["A"]["sample_values"]) == 3
    assert cm["A"]["distinct_count"] == 12
    assert cm["A"]["cell_count"] == 12


def test_f18_unicode_and_determinism(fx):
    p = fx["f18"]
    r1 = run_understand(p)
    r2 = run_understand(p)
    assert r1.stdout == r2.stdout, "output must be byte-deterministic"
    assert "İSTANBUL" in r1.stdout
    assert "ığüşöçİ" in r1.stdout


def test_determinism_and_purity_compact_output(fx):
    r1 = run_understand(fx["f01"])
    r2 = run_understand(fx["f01"])
    assert r1.stdout == r2.stdout
    lines = r1.stdout.strip().splitlines()
    assert len(lines) == 1, "compact output must be a single JSON line"


def test_pretty_matches_compact(fx):
    compact = parse(run_understand(fx["f01"]))
    pretty = parse(run_understand(fx["f01"], "--pretty"))
    assert compact == pretty
    proc = run_understand(fx["f01"], "--pretty")
    assert len(proc.stdout.splitlines()) > 1


def test_samples_negative_rejected(fx):
    proc = run_understand(fx["f01"], "--samples", "-1", expect_ok=False)
    assert proc.returncode != 0
    err = json.loads(proc.stderr)
    assert err["error_code"] == "SPEC_INVALID"
    assert "recovery" in err


def test_samples_above_max_clamped_with_warning(fx):
    doc = parse(run_understand(fx["f17"], "--samples", "100000"))
    codes = {w["code"] for w in doc["warnings"]}
    assert "SAMPLES_CLAMPED" in codes


def test_missing_file_structured_error(fx, tmp_path):
    proc = run_understand(tmp_path / "nope.xlsx", expect_ok=False)
    assert proc.returncode != 0
    err = json.loads(proc.stderr)
    assert err["error_code"] == "FILE_NOT_FOUND"
    assert "recovery" in err


def _oracle(path):
    """Independent expectation via openpyxl's own normal-mode access."""
    wb = load_workbook(path)
    expected = {}
    for ws in wb.worksheets:
        expected[ws.title] = {
            "merges": sorted(str(r) for r in ws.merged_cells.ranges),
            "hidden_rows": sorted(d2.index for d2 in ws.row_dimensions.values()
                                  if d2.hidden),
            "hidden_cols": sorted(k for k, d2 in ws.column_dimensions.items()
                                  if d2.hidden),
            "tables": sorted((t.displayName, t.ref)
                             for t in ws.tables.values()),
            "dvs": sorted((str(dv.sqref), dv.type)
                          for dv in ws.data_validations.dataValidation),
            "cfs": sorted((str(rng.sqref), rule.type)
                          for rng in ws.conditional_formatting
                          for rule in rng.rules),
            "freeze": ws.freeze_panes,
            "af": ws.auto_filter.ref,
            "state": ws.sheet_state,
            "charts": len(ws._charts),
            "images": len(ws._images),
        }
    return expected, common.fingerprint_structure(wb)


@pytest.mark.parametrize("key", ["f01", "f02", "f03", "f04", "f08", "f09",
                                 "chart"])
def test_structure_matches_openpyxl_oracle(fx, key):
    """Raw OOXML structure pass must agree with openpyxl, byte for byte
    (fingerprint) and field for field."""
    doc = parse(run_understand(fx[key]))
    expected, fingerprint = _oracle(fx[key])
    assert doc["workbook"]["structure_fingerprint"] == fingerprint
    for s in doc["sheets"]:
        o = expected[s["name"]]
        assert sorted(m["range"] for m in s["merged_cells"]) == o["merges"]
        assert sorted(r["index"] for r in s["hidden"]["rows"]) == \
            o["hidden_rows"]
        assert sorted(x["index"] for x in s["hidden"]["columns"]) == \
            o["hidden_cols"]
        assert sorted((t["name"], t["ref"]) for t in s["tables"]) == \
            o["tables"]
        assert sorted((v["range"], v["type"])
                      for v in s["validations"]) == o["dvs"]
        assert sorted((c2["range"], c2["kind"])
                      for c2 in s["conditional_formats"]) == o["cfs"]
        assert s["freeze_panes"] == o["freeze"]
        assert s["autofilter"] == o["af"]
        assert s["visibility"] == o["state"]
        assert s["charts"] == o["charts"]
        assert s["images"] == o["images"]
