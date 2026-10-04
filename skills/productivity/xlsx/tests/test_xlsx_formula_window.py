#!/usr/bin/env python3
"""FAZ 4 / step 3 tests -- window classification (D4).

Rules under test:
  * a verdict needs >= 3 instances on a regularly spaced line AND a unanimous
    start/end behaviour -- a single formula (or a two-member family) is never
    classified;
  * UNKNOWN_WINDOW is a real result: insufficient instances, irregular
    spacing, contracting windows and mixed behaviours all stay UNKNOWN and
    are never forced into fixed/rolling (spec section 12);
  * fixed / expanding / rolling / single_row / single_cell / whole_column /
    whole_row / cross-sheet variants are distinguished from start/end
    behaviour, not from a function name.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

import xlsx_formula_family as fam  # noqa: E402


def window_of(cells, sheet="CAP"):
    report = fam.discover(fam.build_entries(cells, sheet))
    if report["families"]:
        return report["families"][0]["window"]
    if report["singles"]:
        return report["singles"][0]["window"]
    return report["outliers"][0]


def vertical(pattern, count=3, start=10):
    return {f"H{start + i}": pattern.format(r=start + i) for i in range(count)}


def vertical_fn(build, count=3, start=10):
    """Row-dependent formulas whose text needs arithmetic (F{r}:F{r+1})."""
    return {f"H{start + i}": build(start + i) for i in range(count)}


# --- classified windows ---------------------------------------------------

def test_fixed_window():
    verdict = window_of(vertical("=SUM($F$1:$F$10)"))
    assert verdict["window_type"] == "fixed"
    assert verdict["instances"] == 3
    assert verdict["spacing"]["step"] == 1


def test_expanding_window():
    assert window_of(vertical("=SUM($F$1:F{r})"))["window_type"] == "expanding"


def test_rolling_window():
    verdict = window_of(vertical_fn(lambda r: f"=AVERAGE(F{r}:F{r + 1})"))
    assert verdict["window_type"] == "rolling"


def test_single_row_window():
    assert window_of(vertical("=SUM(F{r}:H{r})"))["window_type"] == "single_row"


def test_single_cell_window():
    assert window_of(vertical("=F{r}*G{r}"))["window_type"] == "single_cell"


def test_whole_column_window():
    assert window_of(vertical("=SUM($F:$F)"))["window_type"] == "whole_column"


def test_whole_row_window():
    assert window_of(vertical("=SUM($3:$3)"))["window_type"] == "whole_row"


def test_cross_sheet_fixed_window():
    verdict = window_of(vertical("=SUM('Data'!$A$1:$A$10)"))
    assert verdict["window_type"] == "cross_sheet_fixed"
    assert verdict["cross_sheet"] is True


def test_cross_sheet_relative_window():
    verdict = window_of(vertical_fn(lambda r: f"=SUM('Data'!A{r}:A{r + 9})"))
    assert verdict["window_type"] == "cross_sheet_relative"


# --- UNKNOWN_WINDOW stays UNKNOWN (D4, spec section 12) -------------------

def test_single_formula_is_never_classified():
    verdict = window_of({"H10": "=SUM($F$1:$F$10)"})
    assert verdict["window_type"] == "UNKNOWN_WINDOW"
    assert any("single formula" in item or "insufficient" in item
               for item in verdict["ambiguity"])


def test_two_instances_are_never_classified():
    verdict = window_of(vertical("=F{r}*G{r}", count=2))
    assert verdict["window_type"] == "UNKNOWN_WINDOW"
    assert verdict["instances"] == 2
    assert any("insufficient instances" in item
               for item in verdict["ambiguity"])


def test_irregular_spacing_is_unknown():
    cells = {f"H{row}": f"=F{row}*G{row}" for row in (10, 11, 13)}
    verdict = window_of(cells)
    assert verdict["window_type"] == "UNKNOWN_WINDOW"
    assert verdict["spacing"]["regular"] is False
    assert any("irregular instance spacing" in item
               for item in verdict["ambiguity"])


def test_contracting_window_is_unknown():
    verdict = window_of(vertical("=SUM(F{r}:$F$10)"))
    assert verdict["window_type"] == "UNKNOWN_WINDOW"
    assert any("contracting" in item for item in verdict["ambiguity"])


def test_mixed_behaviours_are_unknown():
    verdict = window_of(vertical("=SUM(F{r}:G{r})+SUM($F$1:$F$5)"))
    assert verdict["window_type"] == "UNKNOWN_WINDOW"
    assert any("mixed window behaviours" in item
               for item in verdict["ambiguity"])


def test_unknown_is_never_forced_into_fixed_or_rolling():
    unknown_cases = [
        {"H10": "=SUM($F$1:$F$10)"},                       # single
        vertical("=F{r}*G{r}", count=2),                   # two instances
        {f"H{row}": f"=F{row}*G{row}" for row in (10, 11, 13)},
        vertical("=SUM(F{r}:$F$10)"),                      # contracting
        vertical("=SUM(F{r}:G{r})+SUM($F$1:$F$5)"),        # mixed
    ]
    for cells in unknown_cases:
        verdict = window_of(cells)
        assert verdict["window_type"] == "UNKNOWN_WINDOW", verdict
        assert verdict["window_type"] not in ("fixed", "rolling", "expanding")


def test_every_verdict_carries_evidence_and_ambiguity():
    for cells in (vertical("=F{r}*G{r}"),
                  vertical("=SUM(F{r}:$F$10)"),
                  {"H10": "=F10*G10"}):
        verdict = window_of(cells)
        assert isinstance(verdict["evidence"], list)
        assert isinstance(verdict["ambiguity"], list)
        assert verdict["window_type"] in fam.WINDOW_TYPES


def test_window_types_are_closed_and_unique():
    assert len(fam.WINDOW_TYPES) == len(set(fam.WINDOW_TYPES))
    assert "UNKNOWN_WINDOW" in fam.WINDOW_TYPES


def test_window_is_serialisable_and_deterministic():
    cells = vertical("=SUM($F$1:F{r})")
    first = window_of(cells)
    second = window_of(dict(reversed(list(cells.items()))))
    assert json.dumps(first, sort_keys=True) == json.dumps(second,
                                                          sort_keys=True)


def test_spacing_step_is_recorded_for_every_other_row_tables():
    cells = {f"H{row}": f"=F{row}*G{row}" for row in (10, 12, 14, 16)}
    verdict = window_of(cells)
    assert verdict["spacing"]["regular"] is True
    assert verdict["spacing"]["step"] == 2
    assert verdict["window_type"] == "single_cell"
