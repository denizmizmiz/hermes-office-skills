#!/usr/bin/env python3
# MIT License. Part of the Hermes xlsx skill.
"""FAZ 4 / step 7b -- formula intelligence CLI (spec sections 31-33).

Three verbs, all deterministic and JSON-emitting:

  analyze  read-only workbook walk: parse -> family -> window -> intent ->
           graph (never writes; `values_recalculated` is never claimed);
  plan     read-only synthesis planning (nine precondition gates live in the
           planner; the CLI only collects target state from the workbook);
  execute  applies an APPROVED plan through the Faz 1/3C safe-write chain
           (`xlsx_execute.execute_synthesis_plan`); in-place writes require
           `--approve <plan_id>`.

No synthesis logic lives here or in the executor (D7-A): this CLI wires the
pipeline together and prints the verification protocol (spec section 33).
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import xlsx_common as common          # noqa: E402
import xlsx_execute as execute        # noqa: E402
import xlsx_formula_family as fam     # noqa: E402
import xlsx_formula_graph as graph    # noqa: E402
import xlsx_formula_synth as synth    # noqa: E402

CLI_VERSION = "4.7"


def _raw(loaded):
    return execute._raw(loaded)


def collect_formula_cells(path, sheets=None):
    """read_only pass -> ({sheet: {cell: formula_text}}, target-state scan)."""
    loaded = common.load_workbook_safe(path, data_only=False, read_only=True)
    wb = _raw(loaded)
    out = {}
    for name in (sheets or wb.sheetnames):
        ws = wb[name]
        out[name] = {c.coordinate: c.value for row in ws.iter_rows()
                     for c in row
                     if isinstance(c.value, str) and c.value.startswith("=")}
    wb.close()
    return out


def scan_target_state(path, targets):
    """Writable-scan refused; a read_only pass records emptiness per target."""
    loaded = common.load_workbook_safe(path, data_only=False, read_only=True)
    wb = _raw(loaded)
    states = {}
    for target in targets:
        sheet, cell = target["sheet"], target["cell"]
        try:
            value = wb[sheet][cell].value
        except Exception:                        # noqa: BLE001
            value = None
        states[(sheet, cell)] = {
            "has_formula": isinstance(value, str) and value.startswith("="),
            "has_value": value is not None}
    wb.close()
    return states


def build_engine_report(path, sheets=None):
    cells_by_sheet = collect_formula_cells(path, sheets)
    entries = []
    for sheet, cells in cells_by_sheet.items():
        entries.extend(fam.build_entries(cells, sheet))
    report = fam.discover(entries)
    texts = {(entry["sheet"], entry["cell"]): entry["result"]["text"]
             for entry in entries}
    families_with_text = {
        family["family_id"]:
            texts[(family["sheet"], family["representative"]["cell"])]
        for family in report["families"]}
    greport = graph.build_graph(entries)
    windows, intents = {}, {}
    for family in report["families"]:
        windows[family["window"]["window_type"]] = (
            windows.get(family["window"]["window_type"], 0) + 1)
        intents[family["intent"]["intent"]] = (
            intents.get(family["intent"]["intent"], 0) + 1)
    for single in report["singles"]:
        intent = single["intent"]["intent"]
        intents[intent] = intents.get(intent, 0) + 1
    return {
        "cli_version": CLI_VERSION,
        "workbook": str(path),
        "sha256": hashlib.sha256(Path(path).read_bytes()).hexdigest(),
        "formulas": len(entries),
        "families": len(report["families"]),
        "singles": len(report["singles"]),
        "outliers": len(report["outliers"]),
        "window_breakdown": windows,
        "intent_breakdown": intents,
        "graph": greport["stats"],
        "graph_codes": greport["codes"],
        "unsupported": sorted({code for entry in entries
                               for code in (entry["result"].get("unsupported")
                                            or [])}),
        "values_recalculated": False,
    }, report, families_with_text, entries


def cmd_analyze(args) -> int:
    bundle, report, _, _ = build_engine_report(
        args.workbook, [args.sheet] if args.sheet else None)
    if args.out:
        Path(args.out).write_text(
            json.dumps(bundle, indent=1, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(bundle, indent=1, ensure_ascii=False))
    return 0


def cmd_plan(args) -> int:
    _, report, texts, _ = build_engine_report(
        args.workbook, [args.sheet] if args.sheet else None)
    targets = []
    for item in args.targets.split(","):
        item = item.strip()
        if not item:
            continue
        sheet, cell = item.rsplit("!", 1)
        entry = {"sheet": sheet, "cell": cell}
        if args.family:
            entry["family_id"] = args.family
        targets.append(entry)
    states = scan_target_state(args.workbook, targets)
    plan = synth.plan_synthesis(report, targets, families_with_text=texts,
                                target_states=states)
    rendered = json.dumps(plan, indent=1, ensure_ascii=False)
    if args.out:
        Path(args.out).write_text(rendered, encoding="utf-8")
    summary = {"plan_id": plan["plan_id"],
               "generated_formula_count": plan["generated_formula_count"],
               "blocked": [(b["target"], b["reason_code"])
                           for b in plan["blocked"]],
               "values_recalculated": False}
    print(json.dumps(summary, indent=1, ensure_ascii=False))
    return 0 if plan["generated_formula_count"] else 3


def cmd_execute(args) -> int:
    plan = execute.load_synthesis_plan(args.plan)
    in_place = bool(args.in_place)
    report = execute.execute_synthesis_plan(
        plan, target_path=args.workbook, out_path=args.out,
        in_place=in_place, approve_token=args.approve,
        manifest_dir=Path(args.manifest) if args.manifest else None)
    print(json.dumps(report, indent=1, ensure_ascii=False))
    return 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="xlsx_formula",
                                     description="FAZ 4 formula intelligence")
    parser.add_argument("--version", action="version", version=CLI_VERSION)
    sub = parser.add_subparsers(dest="cmd", required=True)

    analyze = sub.add_parser("analyze", help="read-only formula intelligence")
    analyze.add_argument("--workbook", required=True)
    analyze.add_argument("--sheet")
    analyze.add_argument("--out")
    analyze.set_defaults(func=cmd_analyze)

    plan = sub.add_parser("plan", help="read-only synthesis planning")
    plan.add_argument("--workbook", required=True)
    plan.add_argument("--targets", required=True,
                      help="comma list: SHEET!CELL[,SHEET!CELL...]")
    plan.add_argument("--family", help="explicit source family id")
    plan.add_argument("--sheet")
    plan.add_argument("--out")
    plan.set_defaults(func=cmd_plan)

    run = sub.add_parser("execute", help="apply an approved plan")
    run.add_argument("--plan", required=True)
    run.add_argument("--workbook", required=True)
    run.add_argument("--out")
    run.add_argument("--in-place", action="store_true")
    run.add_argument("--approve", help="plan_id token (required in-place)")
    run.add_argument("--manifest")
    run.set_defaults(func=cmd_execute)

    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except common.XlsxError as exc:
        print(json.dumps({"error": exc.code, "message": str(exc),
                          "recovery": exc.recovery,
                          "context": exc.context or {}},
                         indent=1, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())