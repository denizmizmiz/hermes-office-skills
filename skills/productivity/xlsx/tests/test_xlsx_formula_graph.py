#!/usr/bin/env python3
"""FAZ 4 / step 5 tests -- bounded dependency graph + cycle detection.

Coverage map (spec sections 27-28; AC-4.16 edges recovered >= 99%,
AC-4.17 cycle detection 100%):

  * a labelled graph fixture recovers every expected edge;
  * direct, indirect, self and range-mediated cycles are all detected and
    reported as FORMULA_CYCLE_DETECTED, and cycles are never repaired;
  * acyclic diamonds produce zero false positives;
  * workbook-internal scope (D6): external-workbook formulas contribute no
    edges; cross-sheet edges are counted;
  * caps (D11): exceeding max_edges/max_nodes yields truncated=true plus a
    warning and total/returned counts -- never a silent truncation;
  * determinism: shuffled input yields a byte-identical graph JSON.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

import xlsx_formula_graph as graph  # noqa: E402


def edges_of(report):
    return sorted((edge["from"], edge["to"], edge["kind"])
                  for edge in report["edges"])


# --- edge recovery (AC-4.16) ----------------------------------------------

def test_labelled_graph_recovers_every_expected_edge():
    report = graph.build_from_cells({"CAP": {
        "A1": "=B1+C1", "B1": "=D1*2", "C1": "=SUM(D1:D3)", "D1": "=5"}})
    expected = {("CAP!A1", "CAP!B1", "point"), ("CAP!A1", "CAP!C1", "point"),
                ("CAP!B1", "CAP!D1", "point"), ("CAP!C1", "CAP!D1", "range"),
                ("CAP!C1", "CAP!D1:D3", "range_node")}
    recovered = set(edges_of(report))
    accuracy = len(expected & recovered) / len(expected)
    assert accuracy >= 0.99, recovered
    assert expected <= recovered
    assert recovered - expected == set(), "no spurious edges expected"


def test_range_expands_only_to_known_nodes():
    report = graph.build_from_cells({"CAP": {"A1": "=SUM(B1:B1000000)"}})
    assert report["stats"]["edges"] == 1
    assert edges_of(report)[0][2] == "range_node"


def test_whole_column_reference_uses_known_nodes():
    report = graph.build_from_cells({"CAP": {"A1": "=SUM($F:$F)",
                                             "F5": "=1+1"}})
    assert ("CAP!A1", "CAP!F5", "range") in edges_of(report)


def test_cross_sheet_edge_is_recorded():
    report = graph.build_from_cells({
        "OZET": {"B2": "='Data'!F2"}, "Data": {"F2": "=1+1"}})
    assert ("OZET!B2", "Data!F2", "point") in edges_of(report)
    assert report["stats"]["cross_sheet_edges"] == 1


def test_external_workbook_formula_contributes_no_edges():
    report = graph.build_from_cells({
        "X": {"C1": "=[1]Report!E8"}, "CAP": {"A1": "=B1"}})
    assert report["stats"]["skipped_external_formulas"] == 1
    assert all(edge["from"] != "X!C1" for edge in report["edges"])


# --- cycle detection (AC-4.17) --------------------------------------------

def test_direct_cycle_detected():
    report = graph.build_from_cells({"CAP": {"A1": "=B1", "B1": "=A1"}})
    assert report["stats"]["cycle_count"] >= 1
    assert graph.CYCLE_CODE in report["codes"]
    assert report["cycles"][0]["nodes"] == ["CAP!A1", "CAP!B1"]


def test_indirect_cycle_detected():
    report = graph.build_from_cells({"CAP": {"A1": "=B1", "B1": "=C1",
                                             "C1": "=A1"}})
    assert report["stats"]["cycle_count"] >= 1
    assert set(report["cycles"][0]["nodes"]) == {"CAP!A1", "CAP!B1", "CAP!C1"}


def test_self_reference_detected():
    report = graph.build_from_cells({"CAP": {"A1": "=A1+1"}})
    assert report["stats"]["cycle_count"] == 1
    assert graph.CYCLE_CODE in report["codes"]


def test_range_mediated_cycle_detected():
    report = graph.build_from_cells({"CAP": {"A1": "=SUM(B1:B10)",
                                             "B5": "=A1"}})
    assert report["stats"]["cycle_count"] >= 1
    assert set(report["cycles"][0]["nodes"]) == {"CAP!A1", "CAP!B5"}


def test_acyclic_diamond_has_no_false_cycle():
    report = graph.build_from_cells({"CAP": {
        "A1": "=B1+C1", "B1": "=D1", "C1": "=D1", "D1": "=5"}})
    assert report["stats"]["cycle_count"] == 0
    assert report["codes"] == []
    assert report["cycles"] == []


def test_cycles_are_reported_never_repaired():
    report = graph.build_from_cells({"CAP": {"A1": "=B1", "B1": "=A1"}})
    assert "nodes" in report and "edges" in report      # graph intact
    assert all(edge["kind"] in ("point", "range", "range_node")
               for edge in report["edges"])


# --- caps and determinism (D11) -------------------------------------------

def test_edge_cap_truncates_with_warning_and_counts():
    import xlsx_formula_family as family
    entries = family.build_entries({f"H{i}": f"=F{i}" for i in range(1, 21)},
                                   "CAP")
    report = graph.build_graph(entries, max_edges=5)
    stats = report["stats"]
    assert stats["truncated"] is True
    assert stats["edges"] == 5
    assert stats["edges_returned"] == 5
    assert report["warnings"]
    assert any("edge cap" in warning for warning in report["warnings"])


def test_node_cap_truncates_node_list_with_warning():
    import xlsx_formula_family as family
    entries = family.build_entries({f"H{i}": f"=F{i}" for i in range(1, 11)},
                                   "CAP")
    report = graph.build_graph(entries, max_nodes=3)
    assert report["stats"]["truncated"] is True
    assert report["stats"]["nodes_returned"] == 3
    assert any("node cap" in warning for warning in report["warnings"])


def test_graph_is_deterministic_for_shuffled_input():
    import xlsx_formula_family as family
    cells = {"A1": "=B1+C1", "B1": "=D1", "C1": "=D1", "D1": "=5"}
    forward = graph.build_graph(family.build_entries(cells, "CAP"))
    backward = graph.build_graph(list(reversed(
        family.build_entries(cells, "CAP"))))
    assert json.dumps(forward, sort_keys=True) == json.dumps(backward,
                                                             sort_keys=True)


def test_locked_edge_cap_value():
    assert graph.MAX_EDGES == 1_000_000
