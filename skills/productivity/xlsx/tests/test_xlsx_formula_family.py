#!/usr/bin/env python3
"""FAZ 4 / step 2 tests -- formula family discovery, confidence, outliers.

Coverage map (spec sections 8-10, 20-21; acceptance criteria AC-4.06 and
AC-4.07):

  * a labelled synthetic oracle covering the spec's family edge cases
    (single, two-member, non-contiguous, interrupted, same signature on a
    different sheet, same signature in a different block, mixed
    absolute/relative, cross-sheet, nested functions, IF, SUM, COUNT,
    arithmetic, percentage, date logic) -- accuracy is measured against the
    labels (>= 95%);
  * outlier detection: a known deviating member is reported with
    `outlier_of`, never absorbed (100% on the labelled cases);
  * confidence is evidence-driven (HIGH/MEDIUM/LOW + evidence[]), never a
    probability;
  * determinism: shuffled input order yields a byte-identical report.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

import xlsx_formula_family as fam  # noqa: E402

MUL = "=RC[-2]*RC[-1]"


def report_for(cells, sheet="CAP"):
    return fam.discover(fam.build_entries(cells, sheet))


# --- labelled oracle (AC-4.06, AC-4.07) -----------------------------------

def _row(start, count, pattern):
    """{cell: formula} for rows start..start+count-1 using an f-string."""
    return {f"H{start + i}": pattern.format(r=start + i)
            for i in range(count)}


LABELLED = [
    ("single_formula",
     {"CAP": {"H10": "=F10*G10"}},
     {"families": 0, "singles": 1, "outliers": 0, "confidence": []}),
    ("two_member_adjacent",
     {"CAP": {"H10": "=F10*G10", "H11": "=F11*G11"}},
     {"families": 1, "singles": 0, "outliers": 0,
      "confidence": ["MEDIUM"], "members": [2]}),
    ("three_member_adjacent",
     {"CAP": _row(10, 3, "=F{r}*G{r}")},
     {"families": 1, "singles": 0, "outliers": 0,
      "confidence": ["HIGH"], "members": [3]}),
    ("non_contiguous_row_gap",
     {"CAP": {"H10": "=F10*G10", "H11": "=F11*G11", "H13": "=F13*G13"}},
     {"families": 1, "singles": 0, "outliers": 0,
      "confidence": ["MEDIUM"], "members": [3], "gaps": [1]}),
    ("interrupted_family",
     {"CAP": {"H10": "=F10*G10", "H11": "=F11*G11", "H14": "=F14*G14"}},
     {"families": 1, "singles": 0, "outliers": 0,
      "confidence": ["MEDIUM"], "members": [3], "gaps": [1]}),
    ("separate_block",
     {"CAP": {"H10": "=F10*G10", "H11": "=F11*G11",
              "H15": "=F15*G15", "H16": "=F16*G16"}},
     {"families": 2, "singles": 0, "outliers": 0,
      "confidence": ["MEDIUM", "MEDIUM"]}),
    ("same_signature_different_sheet",
     {"CAP": {"H10": "=F10*G10", "H11": "=F11*G11"},
      "OZET": {"H10": "=F10*G10", "H11": "=F11*G11"}},
     {"families": 2, "singles": 0, "outliers": 0}),
    ("mixed_abs_rel_split",
     {"CAP": {"H10": "=F10*G10", "H11": "=F$10*G10", "H12": "=F12*G12"}},
     {"families": 1, "singles": 0, "outliers": 1,
      "confidence": ["LOW"], "members": [2]}),
    ("cross_sheet_family",
     {"CAP": _row(10, 3, "=SUM('Data'!A$1:A{r})")},
     {"families": 1, "singles": 0, "outliers": 0, "confidence": ["HIGH"]}),
    ("nested_if_family",
     {"CAP": _row(10, 3, '=IF(AND(F{r}>0,G{r}<100),"ok","check")')},
     {"families": 1, "singles": 0, "outliers": 0, "confidence": ["HIGH"]}),
    ("sum_family",
     {"CAP": _row(10, 4, "=SUM(F{r}:G{r})")},
     {"families": 1, "singles": 0, "outliers": 0, "confidence": ["HIGH"]}),
    ("count_family",
     {"CAP": _row(10, 3, "=COUNT(F{r}:G{r})")},
     {"families": 1, "singles": 0, "outliers": 0, "confidence": ["HIGH"]}),
    ("arithmetic_family",
     {"CAP": _row(10, 5, "=F{r}-G{r}+100")},
     {"families": 1, "singles": 0, "outliers": 0, "confidence": ["HIGH"]}),
    ("percentage_family",
     {"CAP": _row(10, 3, "=F{r}/G{r}")},
     {"families": 1, "singles": 0, "outliers": 0, "confidence": ["HIGH"]}),
    ("date_family",
     {"CAP": _row(10, 3, "=E{r}+30")},
     {"families": 1, "singles": 0, "outliers": 0, "confidence": ["HIGH"]}),
    # an interior outlier leaves a hole between members -> confidence is
    # capped at MEDIUM (member continuity evidence is not "no gaps")
    ("outlier_interior_deviation",
     {"CAP": {"H10": "=F10*G10", "H11": "=F11*G11", "H12": "=F12*G12",
              "H13": "=F13/H13", "H14": "=F14*G14"}},
     {"families": 1, "singles": 0, "outliers": 1, "confidence": ["MEDIUM"],
      "members": [4]}),
    # an edge outlier does not interrupt the member run -> HIGH stays valid
    ("outlier_at_edge",
     {"CAP": {"H10": "=F10*G10", "H11": "=F11*G11", "H12": "=F12*G12",
              "H13": "=F13/H13"}},
     {"families": 1, "singles": 0, "outliers": 1, "confidence": ["HIGH"],
      "members": [3]}),
    ("all_different_no_family",
     {"CAP": {"H10": "=F10*G10", "H11": "=F11-H11", "H12": "=F12/G12"}},
     {"families": 0, "singles": 3, "outliers": 0}),
    ("horizontal_family",
     {"CAP": {"E5": "=E3*E4", "F5": "=F3*F4", "G5": "=G3*G4"}},
     {"families": 1, "singles": 0, "outliers": 0, "confidence": ["HIGH"]}),
    ("unsupported_family_blocks_synthesis",
     {"CAP": _row(10, 3, '=INDIRECT("F"&ROW())*G{r}')},
     {"families": 1, "singles": 0, "outliers": 0, "confidence": ["HIGH"],
      "synthesis_allowed": [False]}),
]


def _check(case_report, expected):
    if expected.get("families") is not None:
        assert case_report["stats"]["families"] == expected["families"], \
            case_report["stats"]
    if "singles" in expected:
        assert case_report["stats"]["singles"] == expected["singles"]
    if "outliers" in expected:
        assert case_report["stats"]["outliers"] == expected["outliers"]
    if "confidence" in expected:
        assert [f["confidence"] for f in case_report["families"]] == \
            expected["confidence"]
    if "members" in expected:
        assert [f["member_count"] for f in case_report["families"]] == \
            expected["members"]
    if "gaps" in expected:
        assert [f["gaps"] for f in case_report["families"]] == expected["gaps"]
    if "synthesis_allowed" in expected:
        assert [f["synthesis_allowed"] for f in case_report["families"]] == \
            expected["synthesis_allowed"]


def test_labelled_oracle_accuracy():
    passed = 0
    for name, sheets, expected in LABELLED:
        entries = []
        for sheet, cells in sheets.items():
            entries.extend(fam.build_entries(cells, sheet))
        case_report = fam.discover(entries)
        try:
            _check(case_report, expected)
            passed += 1
        except AssertionError as exc:
            print("case %s failed: %s" % (name, exc))
    accuracy = passed / len(LABELLED)
    assert accuracy >= 0.95, "family accuracy %.3f (%d/%d)" % (
        accuracy, passed, len(LABELLED))


def test_outlier_is_reported_with_outlier_of():
    report = report_for({"H10": "=F10*G10", "H11": "=F11*G11",
                         "H12": "=F12*G12", "H13": "=F13/H13"})
    assert len(report["outliers"]) == 1
    outlier = report["outliers"][0]
    assert outlier["cell"] == "H13"
    assert outlier["outlier_of"] == report["families"][0]["family_id"]
    assert outlier["signature"] != report["families"][0]["signature"]


def test_no_silent_absorption_when_no_signature_repeats():
    report = report_for({"H10": "=F10*G10", "H11": "=F11-H11",
                         "H12": "=F12/G12"})
    assert report["families"] == []
    assert report["stats"]["singles"] == 3


def test_confidence_is_evidence_driven_not_a_probability():
    report = report_for(_row(10, 3, "=F{r}*G{r}"))
    family = report["families"][0]
    assert family["confidence"] == "HIGH"
    assert family["evidence"]
    assert all(isinstance(item, str) for item in family["evidence"])
    assert "confidence_score" not in json.dumps(family)


def test_unsupported_family_is_reported_in_evidence():
    report = report_for(_row(10, 3, '=INDIRECT("F"&ROW())*G{r}'))
    family = report["families"][0]
    assert family["synthesis_eligible"] is False
    assert family["synthesis_allowed"] is False
    assert any("unsupported member" in item for item in family["evidence"])


def test_unsupported_outlier_is_reported():
    report = report_for({"H10": "=F10*G10", "H11": "=F11*G11",
                         "H12": '=INDIRECT("F12")*G12'})
    outlier = report["outliers"][0]
    assert outlier["cell"] == "H12"
    assert outlier["unsupported"] == ["UNSUPPORTED_LATE_BOUND"]
    assert outlier["synthesis_eligible"] is False


def test_determinism_shuffled_input():
    cells = _row(10, 5, "=F{r}*G{r}")
    cells["H20"] = "=F20/H20"
    ordered = fam.build_entries(cells, "CAP")
    shuffled = list(reversed(ordered))
    first = fam.discover(ordered)
    second = fam.discover(shuffled)
    assert json.dumps(first, sort_keys=True) == json.dumps(second,
                                                          sort_keys=True)


def test_family_ids_are_stable_and_ordered():
    report = report_for({"H10": "=F10*G10", "H11": "=F11*G11",
                         "J10": "=F10/G10", "J11": "=F11/G11"})
    ids = [family["family_id"] for family in report["families"]]
    assert ids == sorted(ids)
    assert len(set(ids)) == len(ids)


def test_single_formula_yields_no_family_but_a_single_record():
    report = report_for({"H10": "=F10*G10"})
    assert report["families"] == []
    assert report["singles"][0]["cell"] == "H10"
    assert report["stats"]["coverage_fraction"] == 0.0


def test_line_orientation_is_recorded():
    vertical = report_for(_row(10, 3, "=F{r}*G{r}"))
    assert vertical["families"][0]["orientation"] == "vertical"
    horizontal = report_for({"E5": "=E3*E4", "F5": "=F3*F4", "G5": "=G3*G4"})
    assert horizontal["families"][0]["orientation"] == "horizontal"


def test_cross_sheet_topology_is_in_evidence():
    report = report_for(_row(10, 3, "=SUM('Data'!A$1:A{r})"))
    family = report["families"][0]
    assert family["cross_sheet"] == ["Data"]
    assert any("cross-sheet topology" in item for item in family["evidence"])


def test_non_formula_cells_are_ignored():
    report = report_for({"H10": "=F10*G10", "H11": "hello", "H12": None})
    assert report["stats"]["formulas"] == 1
