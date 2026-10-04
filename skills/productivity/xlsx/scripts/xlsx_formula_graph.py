#!/usr/bin/env python3
# MIT License. Part of the Hermes xlsx skill.
"""FAZ 4 / step 5 -- bounded formula dependency graph + cycle detection.

Scope and limits (locked decisions):
  D6  workbook-internal only: external-workbook formulas contribute no edges
      (they stay `UNSUPPORTED_EXTERNAL_REFERENCE` upstream);
  D11 caps: a graph is bounded by `MAX_NODES` / `MAX_EDGES` (1,000,000 edges,
      locked); a capped list reports count_total / returned / truncated and a
      warning -- never a silent truncation;
  cycles: detected, reported as `FORMULA_CYCLE_DETECTED` + the cycle path,
      and never "fixed" (spec section 28). No Excel calculation engine is
      rebuilt (spec section 27).

Deterministic: nodes and edges are sorted before output, so the same workbook
always produces the same graph JSON.
"""
from __future__ import annotations

import bisect
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_formula_parse as fp  # noqa: E402

GRAPH_VERSION = "4.5"
MAX_NODES = 500_000
MAX_EDGES = 1_000_000              # D11 locked
MAX_CYCLES_REPORTED = 100
CYCLE_CODE = "FORMULA_CYCLE_DETECTED"


def node_id(sheet: str, cell: str) -> str:
    return "%s!%s" % (sheet, cell.upper())


def _cell_of(node: str):
    sheet, _, cell = node.rpartition("!")
    return sheet, cell


def _sheet_index(entries):
    """Per-sheet indexes for membership tests, sorted for bisect."""
    formula_nodes = set()
    by_sheet = {}
    for entry in entries:
        if not entry["result"].get("is_formula"):
            continue
        node = node_id(entry["sheet"], entry["cell"])
        formula_nodes.add(node)
        col_row = fp.parse_cell(entry["cell"])
        if col_row is None:
            continue
        by_sheet.setdefault(entry["sheet"], []).append((col_row[1], col_row[0],
                                                        node))
    for sheet, rows in by_sheet.items():
        rows.sort()
    # secondary index: (sheet, col) -> sorted row numbers, so a whole-column
    # range is a bisect instead of a full-table scan per reference
    by_sheet_cols = {}
    for sheet, rows in by_sheet.items():
        cols = {}
        for row, col, node in rows:
            cols.setdefault(col, []).append(row)
        by_sheet_cols[sheet] = cols
    return formula_nodes, by_sheet, by_sheet_cols


MAX_RANGE_MEMBERS = 1000                # bounded per-reference expansion


def _bound_label(pos) -> str:
    """A1-ish label for one range bound; cols may be None (whole-row refs)."""
    if not pos:
        return "?"
    col = (fp._index_to_col(pos["col"])
           if pos.get("col") is not None else "")
    row = pos.get("row")
    return "%s%s" % (col, row if row is not None else "")


def _members_in_range(by_sheet, by_sheet_cols, sheet, start, end):
    """Known formula nodes inside an A1 range (bounds may be partial).

    Bounded twice (D11): the lookup itself is indexed/bisect-based, and the
    member list is capped at `MAX_RANGE_MEMBERS`. The range's own
    `range_node` edge still records the full dependency, so the cap can hide
    edges but never a dependency.
    """
    rows = by_sheet.get(sheet)
    if not rows:
        return []
    cols_index = by_sheet_cols.get(sheet, {})
    start_col = start.get("col")
    end_col = end.get("col")
    lo_col, hi_col = ((start_col, end_col) if start_col is not None
                      and end_col is not None and end_col > start_col
                      else (end_col, start_col))
    out = []
    if start.get("row") is not None and end.get("row") is not None:
        low, high = sorted((start["row"], end["row"]))
        slice_ = rows[bisect.bisect_left(rows, (low,)):bisect.bisect_right(
            rows, (high, float("inf")))]
        for row in slice_:                       # early exit at the cap --
            if lo_col is not None and not lo_col <= row[1] <= hi_col:
                continue
            out.append(row[2])                   # never build the full list
            if len(out) >= MAX_RANGE_MEMBERS:
                break
    elif start_col is not None and end_col is not None:            # whole col
        for col in range(lo_col, hi_col + 1):
            if not cols_index.get(col):
                continue
            out.extend(_nodes_of_row_range(by_sheet, sheet, col,
                                           start.get("row"),
                                           end.get("row"),
                                           cap=MAX_RANGE_MEMBERS - len(out)))
            if len(out) >= MAX_RANGE_MEMBERS:
                break
    return out


def _nodes_of_row_range(by_sheet, sheet, col, row_low, row_high,
                        cap=MAX_RANGE_MEMBERS):
    rows = by_sheet[sheet]
    low = (row_low if row_low is not None else 0, col)
    high = ((row_high if row_high is not None else float("inf")), float("inf"))
    slice_ = rows[bisect.bisect_left(rows, low):bisect.bisect_right(rows,
                                                                  high)]
    out = []
    for row in slice_:
        if row[1] != col:
            continue
        out.append(row[2])
        if len(out) >= cap:
            break
    return out


def build_graph(entries, *, max_nodes: int = MAX_NODES,
                max_edges: int = MAX_EDGES) -> dict:
    """Dependency graph: formula nodes -> every node they read.

    Ranges expand to the *known* nodes inside them (formula cells and
    point-referenced cells), never to whole-columns worth of new cells, so
    the graph stays bounded on real files (large real-world sheets).
    """
    formula_entries = [entry for entry in entries
                       if entry["result"].get("is_formula")]
    formula_nodes, by_sheet, by_sheet_cols = _sheet_index(formula_entries)
    nodes = set(formula_nodes)
    edges = set()
    warnings = []
    skipped_external = 0
    truncated_edges = False

    def add_edge(source, target, kind):
        nonlocal truncated_edges
        if len(edges) >= max_edges:
            truncated_edges = True
            return False
        edges.add((source, target, kind))
        nodes.add(source)
        nodes.add(target)
        return True

    for entry in formula_entries:
        result = entry["result"]
        if "UNSUPPORTED_EXTERNAL_REFERENCE" in (result.get("unsupported") or []):
            skipped_external += 1
            continue                      # another workbook: no edges (D6)
        source = node_id(entry["sheet"], entry["cell"])
        nodes.add(source)
        home = entry["sheet"]
        for ref in result.get("refs") or []:
            sheet = ref.get("sheet") or home
            offsets = ref.get("offsets") or {}
            start = offsets.get("start")
            end = offsets.get("end")
            if start is None:
                continue
            if end is None:
                target = node_id(sheet, "%s%d" % (fp._index_to_col(start["col"]),
                                                  start["row"]))
                add_edge(source, target, "point")
                continue
            if not truncated_edges:
                members = _members_in_range(by_sheet, by_sheet_cols, sheet,
                                            start, end)
                for member in members:
                    add_edge(source, member, "range")
            # one aggregate node per range keeps the dependency visible
            # without expanding a whole column into a million cells
            add_edge(source, "%s!%s:%s" % (sheet, _bound_label(start),
                                          _bound_label(end)), "range_node")

    if truncated_edges:
        warnings.append("edge cap reached (%d); graph truncated" % max_edges)
    if len(nodes) > max_nodes:
        node_list = sorted(nodes)[:max_nodes]
        warnings.append("node cap reached (%d); node list truncated"
                        % max_nodes)
        truncated_nodes = True
    else:
        node_list = sorted(nodes)
        truncated_nodes = False

    edge_list = sorted(edges)
    report = {
        "graph_version": GRAPH_VERSION,
        "nodes": node_list,
        "edges": [{"from": a, "to": b, "kind": kind} for a, b, kind in edge_list],
        "stats": {
            "nodes": len(nodes),
            "nodes_returned": len(node_list),
            "formula_nodes": len(formula_nodes),
            "edges": len(edges),
            "edges_returned": len(edge_list),
            "cross_sheet_edges": len([e for e in edge_list
                                      if _cell_of(e[0])[0] != _cell_of(e[1])[0]
                                      and "!" in e[1]]),
            "skipped_external_formulas": skipped_external,
            "truncated": truncated_edges or truncated_nodes,
        },
        "warnings": warnings,
        "cycles": [],
        "codes": [],
    }
    cycles, cycle_count = detect_cycles(report)
    report["cycles"] = cycles
    report["stats"]["cycle_count"] = cycle_count
    report["stats"]["cycles_returned"] = len(cycles)
    if cycle_count:
        report["codes"].append(CYCLE_CODE)
    return report


def detect_cycles(report, *, limit: int = MAX_CYCLES_REPORTED):
    """Iterative DFS cycle detection; returns (cycles, total_count).

    Deterministic (sorted nodes/edges), bounded (`limit` reported cycles) and
    read-only: cycles are reported, never repaired (spec section 28).
    """
    adjacency = {}
    for edge in report["edges"]:
        adjacency.setdefault(edge["from"], []).append(edge["to"])
    for source in adjacency:
        adjacency[source] = sorted(set(adjacency[source]))
    color = {}
    parent = {}
    cycles, seen = [], set()
    total = 0
    for root in sorted(report["nodes"]):
        if color.get(root, 0) != 0:
            continue
        stack = [(root, iter(adjacency.get(root, ())))]
        color[root] = 1
        parent[root] = None
        while stack:
            node, children = stack[-1]
            advanced = False
            for child in children:
                if color.get(child, 0) == 0:
                    color[child] = 1
                    parent[child] = node
                    stack.append((child, iter(adjacency.get(child, ()))))
                    advanced = True
                    break
                if color.get(child, 0) == 1:
                    total += 1
                    if len(cycles) < limit:
                        path = [child, node]
                        walker = parent.get(node)
                        while walker is not None and walker != child:
                            path.append(walker)
                            walker = parent.get(walker)
                        path.reverse()
                        key = _canonical_cycle(path)
                        if tuple(key) not in seen:
                            seen.add(tuple(key))
                            cycles.append({"code": CYCLE_CODE,
                                           "nodes": key,
                                           "length": len(key)})
            if not advanced:
                color[node] = 2
                stack.pop()
    cycles.sort(key=lambda item: item["nodes"])
    return cycles, total


def _canonical_cycle(path):
    """Rotate a cycle so it starts at its smallest node (dedupe-safe)."""
    if not path:
        return path
    start = path.index(min(path))
    return path[start:] + path[:start]


def build_from_cells(cells_by_sheet: dict) -> dict:
    """Convenience: {sheet: {cell: formula}} -> graph report."""
    import xlsx_formula_family as family     # local import: avoids a cycle
    entries = []
    for sheet, cells in cells_by_sheet.items():
        entries.extend(family.build_entries(cells, sheet))
    return build_graph(entries)


def main(argv=None) -> int:
    import argparse
    import json
    parser = argparse.ArgumentParser(description="formula dependency graph")
    parser.add_argument("--cells", required=True,
                        help="JSON dump: {sheet: {cell: formula}}")
    parser.add_argument("--out")
    args = parser.parse_args(argv)
    cells_by_sheet = json.loads(Path(args.cells).read_text(encoding="utf-8"))
    report = build_from_cells(cells_by_sheet)
    if args.out:
        Path(args.out).write_text(json.dumps(report, indent=1,
                                             ensure_ascii=False),
                                  encoding="utf-8")
    console = dict(report)
    console["nodes"] = report["nodes"][:40]
    console["edges"] = report["edges"][:40]
    print(json.dumps(console, indent=1, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
