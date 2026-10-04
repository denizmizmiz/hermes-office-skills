#!/usr/bin/env python3
"""FAZ 4 / step 4 tests -- intent inference (D5, spec sections 18-19).

Rules under test:
  * `INTENT_CLASSES` is the taxonomy source of truth and matches the spec's
    minimum class list exactly (drift guard -- documents never restate a
    count, they reference the constant);
  * intent comes from evidence: a function name alone never decides
    ("=IF(D10,1,2)" without a comparison does NOT become conditional);
  * `unknown` is a real result and is never silently coerced into another
    class;
  * confidence is HIGH/MEDIUM/LOW with evidence[]/ambiguity[], downgraded
    when competing candidates exist.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

import xlsx_formula_family as fam  # noqa: E402

# spec section 18 (minimum class list) + the explicit unknown default
SPEC_INTENT_CLASSES = {
    "calculation", "aggregation", "lookup", "conditional", "date_logic",
    "validation_logic", "reference", "summary", "ratio", "percentage",
    "unknown",
}


def intent_of(text, cell="H10"):
    return fam.classify_intent(fam.fp.parse_formula(text, sheet="CAP",
                                                    cell=cell))


# --- taxonomy drift guard (D5) --------------------------------------------

def test_intent_classes_match_spec_list():
    assert set(fam.INTENT_CLASSES) == SPEC_INTENT_CLASSES
    assert len(fam.INTENT_CLASSES) == len(set(fam.INTENT_CLASSES))


def test_every_class_is_reachable():
    fixtures = {
        "calculation": "=F10+G10",
        "aggregation": "=SUM(F10:G10)",
        "lookup": "=VLOOKUP(D10,$A$1:$B$20,2,FALSE)",
        "conditional": '=IF(D10="GRS","Yes","No")',
        "date_logic": "=TODAY()",
        "validation_logic": "=ISNUMBER(F10)",
        "reference": "=F10",
        "summary": "=SUM(F10:F20)",
        "ratio": "=F10/G10",
        "percentage": "=F10/G10*100",
        "unknown": '=CONCATENATE("a","b")',
    }
    for expected, text in fixtures.items():
        cell = "F21" if expected == "summary" else "H10"
        assert intent_of(text, cell)["intent"] == expected, text


# --- evidence discipline --------------------------------------------------

def test_function_name_alone_is_not_enough():
    verdict = intent_of("=IF(D10,1,2)")          # no comparison operator
    assert verdict["intent"] != "conditional"
    assert verdict["intent"] == "unknown"
    assert any("without a comparison" in item for item in verdict["ambiguity"])


def test_nested_aggregation_is_recorded_as_ambiguity():
    verdict = intent_of('=IF(SUM(F10:F20)>0,"x","y")', cell="H23")
    assert verdict["intent"] == "conditional"
    assert verdict["confidence"] == "MEDIUM"     # downgraded: ambiguity
    assert any("aggregation" in item for item in verdict["ambiguity"])


def test_error_wrapper_defers_to_inner_expression():
    verdict = intent_of("=IFERROR(INDEX(B:B,MATCH(D10,A:A,0)),0)")
    assert verdict["intent"] == "lookup"
    assert any("error-handling wrapper" in item for item in verdict["evidence"])


def test_summary_needs_position_evidence():
    above = intent_of("=SUM(F10:F20)", cell="F21")
    elsewhere = intent_of("=SUM(F10:F20)", cell="H20")
    assert above["intent"] == "summary"
    assert elsewhere["intent"] == "aggregation"


def test_percentage_vs_ratio_and_format_blind_note():
    ratio = intent_of("=F10/G10")
    percentage = intent_of("=F10/G10*100")
    assert ratio["intent"] == "ratio" and percentage["intent"] == "percentage"
    assert any("number format unavailable" in item
               for item in ratio["evidence"])
    assert ratio["confidence"] == "MEDIUM"       # format-blind cap


def test_unknown_is_never_silently_coerced():
    verdict = intent_of('=CONCATENATE("a","b")')
    assert verdict["intent"] == "unknown"
    assert verdict["confidence"] == "LOW"
    assert verdict["ambiguity"]


def test_intent_records_are_deterministic_and_serialisable():
    first = intent_of('=IF(D10="GRS","Yes","No")')
    second = intent_of('=IF(D10="GRS","Yes","No")')
    assert json.dumps(first, sort_keys=True) == json.dumps(second,
                                                           sort_keys=True)
    assert first["intent_version"] == fam.INTENT_VERSION


def test_intent_is_attached_to_families_singles_and_outliers():
    report = fam.discover(fam.build_entries({
        "H10": "=F10*G10", "H11": "=F11*G11", "H12": "=F12/H12",
        "J10": "=SUM(F10:G10)"}, "CAP"))
    family = report["families"][0]
    assert family["intent"]["intent"] == "calculation"
    assert report["outliers"][0]["intent"]["intent"] == "ratio"
    assert report["singles"][0]["intent"]["intent"] == "aggregation"


def test_family_members_never_disagree_silently():
    # identical signatures imply identical structure; the guard exists for
    # defensive honesty -- verify the verdict exists on every family
    report = fam.discover(fam.build_entries(
        {f"H{10 + i}": f"=F{10 + i}*G{10 + i}" for i in range(4)}, "CAP"))
    verdict = report["families"][0]["intent"]
    assert verdict["intent"] == "calculation"
    assert "disagree" not in " ".join(verdict["evidence"])
