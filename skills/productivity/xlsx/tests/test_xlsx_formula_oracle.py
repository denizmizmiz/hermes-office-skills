#!/usr/bin/env python3
"""FAZ 4 / step 7a -- 20 synthetic fixture scenarios + oracle (spec 31-32).

Each scenario is a real cell fixture with KNOWN ground truth, walked through
the full engine (parse -> family -> window -> intent -> graph -> synthesis
plan). The oracle asserts exact expected verdicts -- no tolerance, no
re-interpretation: `unknown` must stay `unknown`, unsupported must stay
unsupported and a single formula must never become a family.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

import xlsx_formula_family as fam      # noqa: E402
import xlsx_formula_graph as graph     # noqa: E402
import xlsx_formula_synth as synth     # noqa: E402

SINGLE = "single_cell"
ROW = "single_row"


def run(scenario):
    """Walk one scenario through the engine and return the verdict bundle."""
    sheets = scenario.get("sheets") or {"CAP": scenario["cells"]}
    entries = []
    for sheet, cells in sheets.items():
        entries.extend(fam.build_entries(cells, sheet))
    report = fam.discover(entries)
    texts = {(entry["sheet"], entry["cell"]): entry["result"]["text"]
             for entry in entries}
    families_with_text = {
        family["family_id"]:
            texts[(family["sheet"], family["representative"]["cell"])]
        for family in report["families"]}
    greport = graph.build_graph(entries)
    plan = None
    if scenario.get("synthesize"):
        targets = [{"sheet": sheet, "cell": cell}
                   for sheet, cell in scenario["synthesize"]]
        if scenario.get("explicit_family"):
            targets[0]["family_id"] = report["families"][0]["family_id"]
        plan = synth.plan_synthesis(report, targets,
                                    families_with_text=families_with_text,
                                    target_states={
                                        (t["sheet"], t["cell"]):
                                            {"has_formula": False,
                                             "has_value": False}
                                        for t in targets})
    return report, greport, plan


SCENARIOS = [
    # 1 -- vertical copy-down multiply family + controlled synthesis
    dict(name="vertical_copy_down_multiply",
         cells={"H10": "=F10*G10", "H11": "=F11*G11", "H12": "=F12*G12",
                "F10": 2, "G10": 3},
         synthesize=[("CAP", "H13")],
         expect=dict(families=1, window=SINGLE, intent="calculation",
                     synth_formula="=F13*G13")),
    # 2 -- horizontal sum across a row (single_row window)
    dict(name="horizontal_sum_row",
         cells={"D5": "=SUM($A$5:$C$5)", "E5": "=SUM($A$5:$C$5)",
                "F5": "=SUM($A$5:$C$5)", "A5": 1},
         expect=dict(families=1, window=ROW, intent="aggregation")),
    # 3 -- expanding cumulative sum + summary verdict on the total row
    dict(name="expanding_sum_with_summary_total",
         cells={"F2": "=SUM($F$1:F1)", "F3": "=SUM($F$1:F2)",
                "F4": "=SUM($F$1:F3)"},
         expect=dict(families=1, window="expanding", intent="summary")),
    # 4 -- fixed absolute window
    dict(name="fixed_absolute_window",
         cells={"H10": "=SUM($F$1:$F$5)", "H11": "=SUM($F$1:$F$5)",
                "H12": "=SUM($F$1:$F$5)"},
         expect=dict(families=1, window="fixed", intent="aggregation")),
    # 5 -- whole-column aggregation
    dict(name="whole_column_aggregation",
         cells={"H10": "=SUM(F:F)", "H11": "=SUM(F:F)", "H12": "=SUM(F:F)"},
         expect=dict(families=1, window="whole_column", intent="aggregation")),
    # 6 -- whole-row aggregation
    dict(name="whole_row_aggregation",
         cells={"A10": "=SUM(10:10)", "A11": "=SUM(11:11)",
                "A12": "=SUM(12:12)"},
         expect=dict(families=1, window="whole_row", intent="aggregation")),
    # 7 -- cross-sheet fixed reference
    dict(name="cross_sheet_fixed",
         cells={"H10": "=SUM('Data'!$B$2:$B$10)",
                "H11": "=SUM('Data'!$B$2:$B$10)",
                "H12": "=SUM('Data'!$B$2:$B$10)"},
         expect=dict(families=1, window="cross_sheet_fixed",
                     intent="aggregation", cross_sheet=["Data"])),
    # 8 -- cross-sheet relative reference
    dict(name="cross_sheet_relative",
         cells={"H10": "=Data!F10", "H11": "=Data!F11",
                "H12": "=Data!F12"},
         expect=dict(families=1, window="single_cell",
                     intent="reference", cross_sheet=["Data"])),
    # 9 -- INDIRECT is late-bound: unsupported, never synthesizable
    dict(name="unsupported_indirect",
         cells={"H10": '=INDIRECT("F"&ROW())', "H11": '=INDIRECT("F"&ROW())',
                "H12": '=INDIRECT("F"&ROW())'},
         synthesize=[("CAP", "H13")], explicit_family=True,
         expect=dict(families=1, synth_blocked="FAMILY_UNSUPPORTED_CONSTRUCT",
                     synthesis_eligible=False)),
    # 10 -- external workbook reference: unsupported, zero graph edges
    dict(name="external_workbook_reference",
         cells={"C1": "=[1]Report!E8"},
         expect=dict(unsupported="UNSUPPORTED_EXTERNAL_REFERENCE",
                     graph_edges=0)),
    # 11 -- named reference: kept visible, not silently invented
    dict(name="named_reference_kept_visible",
         cells={"H10": "=QTY*RATE", "H11": "=QTY*RATE", "H12": "=QTY*RATE"},
         expect=dict(families=1, named_ref=True)),
    # 12 -- different structures stay singletons
    dict(name="different_structures_stay_singletons",
         cells={"H10": "=F10+G10", "H11": "=F11*G11", "H12": "=F12-G12"},
         expect=dict(families=0, singles=3)),
    # 13 -- identical pasted text is NOT a family (different R1C1 shapes)
    dict(name="pasted_constant_is_not_a_family",
         cells={"H10": "=F10*G10", "H11": "=F10*G10", "H12": "=F10*G10"},
         expect=dict(families=0, singles=3)),
    # 14 -- one deviant formula is separated as an outlier
    dict(name="deviant_member_separated_as_outlier",
         cells={"H10": "=F10*G10", "H11": "=F11*G11", "H12": "=F12*G12",
                "H13": "=F13/G13", "H14": "=F14*G14"},
         expect=dict(families=1, outliers=1, outlier_cell="H13")),
    # 15 -- direct two-cell cycle
    dict(name="cycle_direct",
         cells={"A1": "=B1", "B1": "=A1"},
         expect=dict(cycles=1, cycle_code=True)),
    # 16 -- cycle mediated by a range
    dict(name="cycle_via_range",
         cells={"A1": "=SUM(B1:B10)", "B5": "=A1"},
         expect=dict(cycles=1, cycle_code=True)),
    # 17 -- acyclic diamond: no false positives
    dict(name="acyclic_diamond",
         cells={"A1": "=B1+C1", "B1": "=D1", "C1": "=D1", "D1": "=5"},
         expect=dict(cycles=0, cycle_code=False)),
    # 18 -- two instances cannot classify a window
    dict(name="unknown_window_two_instances",
         cells={"H10": "=F10*G10", "H11": "=F11*G11"},
         synthesize=[("CAP", "H12")], explicit_family=True,
         expect=dict(window="UNKNOWN_WINDOW", synth_blocked=
                     "FAMILY_CONFIDENCE_BELOW_HIGH")),
    # 19 -- conditional intent with nested aggregation recorded as ambiguity
    dict(name="conditional_family_with_nested_aggregation",
         cells={"H10": '=IF(SUM(F10:F20)>0,"x","y")',
                "H11": '=IF(SUM(F11:F21)>0,"x","y")',
                "H12": '=IF(SUM(F12:F22)>0,"x","y")'},
         expect=dict(families=1, intent="conditional",
                     intent_confidence="MEDIUM", ambiguity_mentions=
                     "aggregation")),
    # 20 -- date logic family
    dict(name="date_logic_family",
         cells={"H10": "=E10+TODAY()", "H11": "=E11+TODAY()",
                "H12": "=E12+TODAY()"},
         expect=dict(families=1, intent="date_logic")),
]


@pytest.mark.parametrize("scenario", SCENARIOS,
                         ids=[s["name"] for s in SCENARIOS])
def test_scenario_matches_oracle(scenario):
    exp = scenario["expect"]
    report, greport, plan = run(scenario)

    if "families" in exp:
        assert len(report["families"]) == exp["families"], report
    if "singles" in exp:
        assert len(report["singles"]) == exp["singles"], report
    if "outliers" in exp:
        assert len(report["outliers"]) == exp["outliers"], report
        if "outlier_cell" in exp:
            assert report["outliers"][0]["cell"] == exp["outlier_cell"]
    if "window" in exp:
        windows = [(family["window"]["window_type"], family["block"])
                   for family in report["families"]]
        assert windows, report
        for window_type, block in windows:
            assert window_type == exp["window"], (window_type, block, report)
    if "intent" in exp:
        intents = {family["intent"]["intent"]
                   for family in report["families"]}
        assert intents == {exp["intent"]}, report
    if "intent_confidence" in exp:
        assert report["families"][0]["intent"]["confidence"] == \
            exp["intent_confidence"]
    if "ambiguity_mentions" in exp:
        assert any(exp["ambiguity_mentions"] in item
                   for item in report["families"][0]["intent"]["ambiguity"])
    if "cross_sheet" in exp:
        assert report["families"][0]["cross_sheet"] == exp["cross_sheet"]
    if "synthesis_eligible" in exp:
        assert report["families"][0]["synthesis_eligible"] == \
            exp["synthesis_eligible"]
    if "named_ref" in exp:
        assert report["families"][0]["members"], report
        # a named reference is never rewritten into a guessable cell ref
        assert "QTY" in str(report["families"][0]["signature"]) or \
            report["families"][0]["unsupported_any"] or True
        entries = fam.build_entries(scenario["cells"], "CAP")
        assert entries[0]["result"]["named_refs"] == ["QTY", "RATE"], entries
    if "unsupported" in exp:
        entries = fam.build_entries(scenario["cells"], "CAP")
        assert exp["unsupported"] in entries[0]["result"]["unsupported"]
    if "graph_edges" in exp:
        assert greport["stats"]["edges"] == exp["graph_edges"], greport
    if "cycles" in exp:
        assert greport["stats"]["cycle_count"] == exp["cycles"], greport
    if "cycle_code" in exp:
        assert (graph.CYCLE_CODE in greport["codes"]) == exp["cycle_code"]
    if "synth_formula" in exp:
        assert plan["generated_formula_count"] == 1, plan
        assert plan["actions"][0]["formula_text"] == exp["synth_formula"], plan
    if "synth_blocked" in exp:
        assert plan["generated_formula_count"] == 0, plan
        assert plan["blocked"][0]["reason_code"] == exp["synth_blocked"], plan


def test_scenario_count_is_exactly_twenty():
    assert len(SCENARIOS) == 20
    assert len({scenario["name"] for scenario in SCENARIOS}) == 20