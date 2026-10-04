#!/usr/bin/env python3
"""FAZ 4 / step 1 tests -- tokenizer, normalizer, signature.

Coverage map (spec sections 5-7, acceptance criteria AC-4.02, AC-4.03,
AC-4.04, AC-4.05, AC-4.10, AC-4.11 and the D2/D8 decisions):

  * reference extraction + R1C1 normalisation (relative / absolute / mixed);
  * string-literal protection and sheet-qualified references;
  * whole-column / whole-row spans (the parser gap measured in preflight);
  * closed unsupported taxonomy: every code in UNSUPPORTED_CODES, no
    synthesis eligibility for any unsupported construct;
  * determinism: same formula + same cell == byte-identical signature;
  * parity with `extract_refs` (the shared parser) -- no second parser;
  * named references are reported, never silently treated as values.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

import xlsx_formula_parse as fp  # noqa: E402


def parse(text, cell="H10", sheet="CAP"):
    return fp.parse_formula(text, sheet=sheet, cell=cell)


# --- references and normalisation (AC-4.02, AC-4.03, AC-4.04) -------------

def test_relative_refs_normalise_to_offsets():
    result = parse("=F10*G10")
    assert result["signature"] == "=RC[-2]*RC[-1]"
    assert result["refs_total"] == 2
    assert result["relative_refs"] == 2
    assert result["absolute_refs"] == 0 and result["mixed_refs"] == 0


def test_absolute_refs_are_preserved():
    result = parse("=$A$1+$B$2")
    assert result["absolute_refs"] == 2
    assert "R1C1" in result["signature"] and "R2C2" in result["signature"]


def test_mixed_refs_are_classified_as_mixed():
    result = parse("=$A1+A$1")
    assert result["mixed_refs"] == 2
    assert result["absolute_refs"] == 0


def test_range_normalisation_and_shape():
    result = parse("=SUM($A$1:A10)")
    assert result["ranges"] == 1
    assert result["range_shapes"] == ["vertical_range"]
    assert result["signature"].startswith("=SUM(R1C1:RC[")


def test_two_dimensional_range_shape():
    assert parse("=SUM(A1:B2)")["range_shapes"] == ["two_dimensional"]


def test_string_literals_are_protected():
    result = parse('=IF(D10="GRS","Yes","No")')
    assert result["refs_total"] == 1
    for literal in ('"GRS"', '"Yes"', '"No"'):
        assert literal in result["signature"]
    assert result["functions"] == ["IF"]


def test_cross_sheet_reference_is_reported():
    result = parse("='Ürün Listesi'!F2")
    assert result["cross_sheet"] == ["Ürün Listesi"]
    assert result["signature"].startswith("='ÜRÜN LISTESI'!")


def test_same_sheet_qualified_reference_is_not_cross_sheet():
    result = parse("=CAP!F2", sheet="CAP")
    assert result["cross_sheet"] == []
    assert result["refs_total"] == 1


# --- whole-column / whole-row (preflight parser gap) ----------------------

def test_whole_column_reference_is_tokenised():
    result = parse("=SUMIF('OUTPUT PRODUCT TRACK'!R:R,'INPUT TC REGISTRY'!P3,"
                   "'OUTPUT PRODUCT TRACK'!E:E)")
    assert result["range_shapes"] == ["whole_column", "single_cell",
                                     "whole_column"]
    assert "#COL" in result["signature"]
    assert "R?" not in result["signature"]
    assert result["unsupported"] == []


def test_absolute_whole_column_token():
    assert "#COL18" in parse("=SUM($R:$R)")["signature"]


def test_whole_row_reference_is_tokenised():
    result = parse("=SUM(3:3)")
    assert result["range_shapes"] == ["whole_row"]
    assert "#ROW" in result["signature"]


# --- closed unsupported taxonomy (AC-4.10, AC-4.11, D8) -------------------

@pytest.mark.parametrize("text,code", [
    ("=SUM(Table_5[Volume])", "UNSUPPORTED_STRUCTURED_FORMULA"),
    ("=[1]Report!E8", "UNSUPPORTED_EXTERNAL_REFERENCE"),
    ("=SUM({1,2;3,4})", "UNSUPPORTED_ARRAY_CONSTANT"),
    ("=_xlfn.XLOOKUP(A1,B:B,C:C)", "UNSUPPORTED_FUTURE_FUNCTION"),
    ("=#REF!+A1", "UNSUPPORTED_BROKEN_REFERENCE"),
    ('=INDIRECT("A"&B1)', "UNSUPPORTED_LATE_BOUND"),
    ("=OFFSET(A1,1,1)", "UNSUPPORTED_LATE_BOUND"),
    ("=TEXTSPLIT(A1,\",\")", "UNSUPPORTED_DYNAMIC_ARRAY"),
])
def test_unsupported_code_and_blocked_synthesis(text, code):
    result = parse(text)
    assert code in result["unsupported"]
    assert result["synthesis_eligible"] is False
    assert set(result["unsupported"]) <= set(fp.UNSUPPORTED_CODES)


def test_structured_and_external_codes_do_not_bleed_into_each_other():
    assert parse("=SUM(Table_5[Volume])")["unsupported"] == [
        "UNSUPPORTED_STRUCTURED_FORMULA"]
    assert parse("=[1]Report!E8")["unsupported"] == [
        "UNSUPPORTED_EXTERNAL_REFERENCE"]


def test_supported_formula_is_synthesis_eligible():
    result = parse("=F10*G10")
    assert result["unsupported"] == []
    assert result["synthesis_eligible"] is True


def test_taxonomy_is_closed_and_free_of_duplicates():
    assert len(fp.UNSUPPORTED_CODES) == len(set(fp.UNSUPPORTED_CODES))
    assert all(code.isupper() for code in fp.UNSUPPORTED_CODES)


# --- named references (spec section 17) -----------------------------------

def test_named_reference_is_reported_not_guessed():
    result = parse("=TaxRate*A10")
    assert result["named_refs"] == ["TaxRate"]
    assert result["unsupported"] == ["UNKNOWN_NAMED_REFERENCE"]
    assert result["synthesis_eligible"] is False


def test_table_name_is_not_a_named_reference():
    assert parse("=SUM(Table_5[Volume])")["named_refs"] == []


def test_sheet_name_and_whole_column_are_not_named_references():
    assert parse("=Listas!$AE$1")["named_refs"] == []
    assert parse("=SUM(R:R)")["named_refs"] == []


# --- determinism (AC-4.05, AC-4.22) --------------------------------------

def test_signature_is_deterministic():
    first = parse("=SUMIF('OUTPUT PRODUCT TRACK'!R:R,C3,D:D)")
    second = parse("=SUMIF('OUTPUT PRODUCT TRACK'!R:R,C3,D:D)")
    assert first == second
    assert json.dumps(first, sort_keys=True) == json.dumps(second,
                                                           sort_keys=True)


def test_same_formula_in_different_cells_differs_only_by_offsets():
    here = parse("=F10*G10", cell="H10")
    there = parse("=F11*G11", cell="H11")
    assert here["signature"] == there["signature"]


def test_signature_carries_its_version():
    assert parse("=A1")["sig_version"] == fp.SIG_VERSION


# --- parser parity (no second parser, spec section 5) --------------------

def test_ref_count_parity_with_extract_refs():
    import xlsx_restructure as restructure
    formulas = ["=F10*G10", "=SUM(A1:B2)", "=IF(D10=\"GRS\",F10,G10)",
                "=SUM('Ürün Listesi'!F2:F21)", "=$A$1+$B$2",
                "=Listas!$AE$1"]
    for formula in formulas:
        result = fp.parse_formula(formula, sheet="CAP", cell="H10")
        assert result["refs_total"] == len(
            restructure.extract_refs(formula, "CAP")), formula


def test_non_formula_input_is_reported_as_such():
    assert parse("hello")["is_formula"] is False
    assert fp.parse_formula(None)["is_formula"] is False


def test_formula_without_home_cell_reports_unknown_offsets():
    result = fp.parse_formula("=F10*G10")
    assert result["base_cell_known"] is False
    assert "R?" in result["signature"] or "C?" in result["signature"]
