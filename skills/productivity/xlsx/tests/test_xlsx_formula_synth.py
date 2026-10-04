#!/usr/bin/env python3
"""FAZ 4 / step 6a tests -- controlled synthesis: planning + validation.

Covers spec preconditions (section 22) and ACs:
  AC-4.10 unsupported -> nothing;            AC-4.11 no source -> nothing;
  AC-4.12 expected formula equality;         AC-4.13 absolute preserved;
  AC-4.14 relative translation;              AC-4.23 plan determinism;
  AC-4.24 QA expectations present; plus D7-A: synthesis never enters
  `xlsx_execute` (module boundary guard).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

import xlsx_formula_family as fam  # noqa: E402
import xlsx_formula_synth as synth  # noqa: E402


def make_report(cells, sheet="CAP"):
    entries = fam.build_entries(cells, sheet)
    report = fam.discover(entries)
    texts = {entry["cell"]: entry["result"]["text"] for entry in entries}
    families_with_text = {fam_["family_id"]: texts[fam_["representative"]["cell"]]
                          for fam_ in report["families"]}
    return report, families_with_text


def plan(cells, targets, states, **kwargs):
    report, texts = make_report(cells)
    return synth.plan_synthesis(report, targets,
                                families_with_text=texts,
                                target_states=states, **kwargs)


FAMILY = {"H10": "=F10*G10", "H11": "=F11*G11", "H12": "=F12*G12"}
STATE_EMPTY = {("CAP", "H13"): {"has_formula": False, "has_value": False}}


# --- happy path -----------------------------------------------------------

def test_controlled_synthesis_matches_expected_formula():
    result = plan(FAMILY, [{"sheet": "CAP", "cell": "H13"}], STATE_EMPTY)
    assert result["generated_formula_count"] == 1
    action = result["actions"][0]
    assert action["formula_text"] == "=F13*G13"     # AC-4.12 expected equality
    assert action["drow"] == 3
    assert action["validation"]["signature_matches_family"] is True


def test_absolute_references_survive_translation():
    cells = {"H10": "=SUM($F$1:F10)", "H11": "=SUM($F$1:F11)",
             "H12": "=SUM($F$1:F12)"}
    result = plan(cells, [{"sheet": "CAP", "cell": "H13"}], STATE_EMPTY)
    assert result["actions"][0]["formula_text"] == "=SUM($F$1:F13)"  # AC-4.13


def test_relative_offsets_translate_by_row_delta():
    result = plan(FAMILY, [{"sheet": "CAP", "cell": "H20"}],
                  {("CAP", "H20"): {"has_formula": False, "has_value": False}})
    action = result["actions"][0]
    assert action["drow"] == 10
    assert action["formula_text"] == "=F20*G20"     # AC-4.14


def test_auto_selection_picks_the_family_above_the_target():
    cells = dict(FAMILY)
    cells.update({"H21": "=F21/2", "H22": "=F22/2", "H23": "=F23/2"})
    report, texts = make_report(cells)
    states = {("CAP", "H13"): {"has_formula": False, "has_value": False},
              ("CAP", "H24"): {"has_formula": False, "has_value": False}}
    result = synth.plan_synthesis(report, [{"sheet": "CAP", "cell": "H13"},
                                           {"sheet": "CAP", "cell": "H24"}],
                                  families_with_text=texts,
                                  target_states=states)
    assert result["actions"][0]["formula_text"] == "=F13*G13"
    assert result["actions"][1]["formula_text"] == "=F24/2"


# --- blocking paths -------------------------------------------------------

def test_no_source_family_generates_nothing():
    report, _ = make_report({})
    result = synth.plan_synthesis(report, [{"sheet": "CAP", "cell": "H13"}],
                                  target_states=STATE_EMPTY)
    assert result["generated_formula_count"] == 0        # AC-4.11
    assert result["blocked"][0]["reason_code"] == "NO_SOURCE_FAMILY"


def test_unsupported_family_is_blocked():
    # auto-selection never even considers a non-eligible family (it cannot
    # reach the synthesis gate); an explicit family_id exercises the gate
    cells = {f"H{10 + i}": '=INDIRECT("F"&ROW())' for i in range(3)}
    report, texts = make_report(cells)
    target = {"sheet": "CAP", "cell": "H13",
              "family_id": report["families"][0]["family_id"]}
    result = synth.plan_synthesis(report, [target], families_with_text=texts,
                                  target_states=STATE_EMPTY)
    assert result["generated_formula_count"] == 0        # AC-4.10
    assert result["blocked"][0]["reason_code"] == "FAMILY_UNSUPPORTED_CONSTRUCT"
    auto = plan(cells, [{"sheet": "CAP", "cell": "H13"}], STATE_EMPTY)
    assert auto["blocked"][0]["reason_code"] == "NO_SOURCE_FAMILY"


def test_unknown_window_family_is_blocked():
    family_report = {"families": [{
        "family_id": "FAM-1", "sheet": "CAP", "orientation": "vertical",
        "confidence": "HIGH", "synthesis_allowed": True,
        "synthesis_eligible": True, "signature": "RCX",
        "window": {"window_type": "UNKNOWN_WINDOW"},
        "block": {"start": "H10", "end": "H12"},
        "representative": {"cell": "H10"}, "member_count": 3,
        "members": [{"cell": "H10"}],
    }]}
    result = synth.plan_synthesis(
        family_report, [{"sheet": "CAP", "cell": "H13"}],
        families_with_text={"FAM-1": "=F10*G10"},
        target_states=STATE_EMPTY)
    assert result["generated_formula_count"] == 0
    assert result["blocked"][0]["reason_code"] == "WINDOW_PATTERN_NOT_DETERMINISTIC"


def test_translation_mismatch_is_blocked():
    family_report = {"families": [{
        "family_id": "FAM-2", "sheet": "CAP", "orientation": "vertical",
        "confidence": "HIGH", "synthesis_allowed": True,
        "synthesis_eligible": True, "signature": "A-DIFFERENT-SIGNATURE",
        "window": {"window_type": "single_cell"},
        "block": {"start": "H10", "end": "H12"},
        "representative": {"cell": "H10"}, "member_count": 3,
        "members": [{"cell": "H10"}],
    }]}
    result = synth.plan_synthesis(
        family_report, [{"sheet": "CAP", "cell": "H13"}],
        families_with_text={"FAM-2": "=F10*G10"},
        target_states=STATE_EMPTY)
    assert result["generated_formula_count"] == 0
    assert result["blocked"][0]["reason_code"] == "REFERENCE_TRANSLATION_MISMATCH"


def test_target_state_unknown_is_blocked():
    result = plan(FAMILY, [{"sheet": "CAP", "cell": "H13"}], {})
    assert result["blocked"][0]["reason_code"] == "TARGET_STATE_UNKNOWN"


def test_existing_formula_target_is_blocked():
    states = {("CAP", "H13"): {"has_formula": True, "has_value": False}}
    result = plan(FAMILY, [{"sheet": "CAP", "cell": "H13"}], states)
    assert result["blocked"][0]["reason_code"] == "TARGET_ALREADY_FORMULA"


def test_target_outside_allowed_set_is_blocked():
    result = plan(FAMILY, [{"sheet": "CAP", "cell": "H13"}], STATE_EMPTY,
                  allowed_targets=[("CAP", "H14")])
    assert result["blocked"][0]["reason_code"] == "TARGET_NOT_ALLOWED"


def test_column_mismatch_is_blocked():
    report, texts = make_report(FAMILY)
    result = synth.plan_synthesis(report, [{"sheet": "CAP", "cell": "J13",
                                            "family_id":
                                            report["families"][0]["family_id"]}],
                                  families_with_text=texts,
                                  target_states={("CAP", "J13"):
                                                 {"has_formula": False,
                                                  "has_value": False}})
    assert result["blocked"][0]["reason_code"] == "TARGET_COLUMN_MISMATCH"


def test_identical_pasted_text_is_not_a_family():
    # "=F10*G10" in H10/H11/H12 is three DIFFERENT normalized formulas
    # (RC[-2] vs R[-1]C[-2] vs R[-2]C[-2]); evidence-first discovery never
    # calls a pasted-constant block a copy-down family
    report, _ = make_report({f"H{10 + i}": "=F10*G10" for i in range(3)})
    assert report["families"] == []
    assert len(report["singles"]) == 3


def test_low_confidence_family_is_blocked():
    cells = {f"H{10 + i}": f"=F{10 + i}*G{10 + i}" for i in range(3)}
    report, texts = make_report(cells)
    family = report["families"][0]
    family["confidence"] = "LOW"                      # force the gate
    result = synth.plan_synthesis(report,
                                  [{"sheet": "CAP", "cell": "H13"}],
                                  families_with_text=texts,
                                  target_states=STATE_EMPTY)
    assert result["blocked"][0]["reason_code"] == "FAMILY_CONFIDENCE_BELOW_HIGH"


# --- plan shape, determinism, module boundary -----------------------------

def test_plan_is_deterministic():                     # AC-4.23
    first = plan(FAMILY, [{"sheet": "CAP", "cell": "H13"}], STATE_EMPTY)
    second = plan(FAMILY, [{"sheet": "CAP", "cell": "H13"}], STATE_EMPTY)
    assert first["plan_id"] == second["plan_id"]
    assert json.dumps(first, sort_keys=True) == json.dumps(second,
                                                           sort_keys=True)


def test_qa_expectations_present():                   # AC-4.24 (plan side)
    result = plan(FAMILY, [{"sheet": "CAP", "cell": "H13"}], STATE_EMPTY)
    qa = result["qa_expectations"]
    assert qa["formula_count_delta_expected"] == 1
    assert qa["values_recalculated"] is False
    assert len(qa["checks"]) >= 7


def test_blocked_codes_are_closed_taxonomy():
    assert isinstance(synth.BLOCK_CODES, tuple)
    assert synth.BLOCKED == "FORMULA_SYNTHESIS_BLOCKED"


def test_synthesis_never_enters_the_executor_module():  # D7-A guard
    source = Path(synth.__file__).read_text(encoding="utf-8")
    assert "import xlsx_execute" not in source
    assert "from xlsx_execute" not in source
