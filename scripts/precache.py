#!/usr/bin/env python3
"""
Precompute chart outputs for the GDP forecast shock slider.

This script evaluates the chart payload once at the default (0 bps), plus all non-zero
shock settings from GDP_SHOCK_BPS_MIN..GDP_SHOCK_BPS_MAX (forecast cells set to
baseline + baseline * (bps * 1e-4)), and writes the results to a JSON file for
fast, pre-cached loading in the webapp.

Run:
  uv run python scripts/precache.py
  uv run python scripts/precache.py --workbook dsf-uga.xlsm
  uv run python scripts/precache.py --libreoffice-check
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from excel_grapher import DependencyGraph, DynamicRefError, FormulaEvaluator, create_dependency_graph

from lic_dsf import graph as lic_graph
from lic_dsf.payload import (
    GDP_SHOCK_BPS_MAX,
    GDP_SHOCK_BPS_MIN,
    build_figure1_payload,
    gdp_forecast_baselines,
    gdp_forecast_cell_keys,
    gdp_forecast_value_from_bps,
)


@dataclass(frozen=True, slots=True)
class CacheEntry:
    bps: int
    payload: dict[str, Any]


def _collect_export_targets() -> list[str]:
    targets: list[str] = []
    for entry in lic_graph.EXPORT_RANGES:
        sheet_name, range_a1 = lic_graph.parse_range_spec(entry["range_spec"])
        targets.extend(lic_graph.cells_in_range(sheet_name, range_a1))
    return targets


def _load_graph(*, no_cache: bool) -> DependencyGraph:
    wb = lic_graph.WORKBOOK_PATH
    cache_path = lic_graph._default_graph_cache_path(wb)
    targets = _collect_export_targets()

    if not no_cache:
        g = lic_graph.try_load_graph_cache(
            cache_path,
            wb,
            targets,
            max_depth=lic_graph.GRAPH_MAX_DEPTH,
        )
        if g is not None:
            return g

    if not wb.is_file():
        raise FileNotFoundError(
            f"Workbook not found: {wb}. No valid graph cache found at {cache_path}."
        )

    try:
        g = create_dependency_graph(
            wb,
            targets,
            load_values=lic_graph.GRAPH_LOAD_VALUES,
            max_depth=lic_graph.GRAPH_MAX_DEPTH,
            use_cached_dynamic_refs=lic_graph.GRAPH_USE_CACHED_DYNAMIC_REFS,
        )
    except DynamicRefError as exc:
        raise RuntimeError(
            "Could not build the dependency graph (DynamicRefError). "
            "Run `uv run python scripts/extract_graph.py --no-cache` "
            "(or `python -m lic_dsf.graph`), "
            "add dynamic-ref constraints, then retry. "
            f"Detail: {exc}"
        ) from exc

    if not no_cache:
        try:
            lic_graph.save_graph_cache(
                cache_path,
                g,
                wb,
                targets,
                max_depth=lic_graph.GRAPH_MAX_DEPTH,
            )
        except OSError:
            pass

    return g


def _compute_entries(*, graph: DependencyGraph) -> tuple[CacheEntry, list[CacheEntry]]:
    keys = gdp_forecast_cell_keys(graph, workbook_path=lic_graph.WORKBOOK_PATH)
    if not keys:
        raise RuntimeError(
            "GDP shock inputs are not in the graph. "
            "Expected forecast cells from the workbook layout to be included in export targets."
        )
    baselines = gdp_forecast_baselines(
        graph, workbook_path=lic_graph.WORKBOOK_PATH, keys=keys
    )
    if len(baselines) != len(keys):
        raise RuntimeError("GDP baseline/key mismatch.")

    settings = lic_graph.get_calc_settings(lic_graph.WORKBOOK_PATH)
    ev = FormulaEvaluator(
        graph,
        iterate_enabled=settings.iterate_enabled,
        iterate_count=settings.iterate_count,
        iterate_delta=settings.iterate_delta,
    )

    def eval_at_bps(bps: int) -> CacheEntry:
        for k, base in zip(keys, baselines, strict=True):
            ev.set_value(k, gdp_forecast_value_from_bps(float(base), bps))
        payload = build_figure1_payload(
            graph,
            workbook_path=lic_graph.WORKBOOK_PATH,
            evaluator=ev,
        )
        return CacheEntry(bps=bps, payload=payload)

    default_entry = eval_at_bps(0)

    shocks: list[CacheEntry] = []
    for bps in range(GDP_SHOCK_BPS_MIN, GDP_SHOCK_BPS_MAX + 1):
        if bps == 0:
            continue
        shocks.append(eval_at_bps(int(bps)))

    return default_entry, shocks


def _entry_to_json(e: CacheEntry) -> dict[str, Any]:
    return {"bps": e.bps, "payload": e.payload}


def _cache_entry_for_bps(
    default_entry: CacheEntry, shocks: list[CacheEntry], bps: int
) -> CacheEntry:
    if bps == 0:
        return default_entry
    for e in shocks:
        if e.bps == bps:
            return e
    raise SystemExit(
        f"No precache entry for bps={bps} (expected one of 0 or GDP_SHOCK_BPS_MIN..MAX)."
    )


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Precompute chart outputs for all GDP shock slider states."
    )
    ap.add_argument(
        "--workbook",
        type=Path,
        default=lic_graph.WORKBOOK_PATH,
        help=f"Source .xlsm workbook (default: {lic_graph.WORKBOOK_PATH}).",
    )
    ap.add_argument(
        "--out",
        type=Path,
        default=Path(".cache/gdp-shocks.json"),
        help="Output JSON path (default: .cache/gdp-shocks.json).",
    )
    ap.add_argument(
        "--no-graph-cache",
        action="store_true",
        help="Ignore graph pickle cache and rebuild the dependency graph from workbook.",
    )
    ap.add_argument(
        "--libreoffice-check",
        action="store_true",
        help="After caching, run LibreOffice recalc and compare Chart Data to FormulaEvaluator (Python vs LO).",
    )
    ap.add_argument(
        "--lo-baseline-bps",
        type=int,
        default=0,
        help="LibreOffice check: reference copy bps (default 0).",
    )
    ap.add_argument(
        "--lo-shock-bps",
        type=int,
        default=10,
        help="LibreOffice check: comparison copy bps (default 10).",
    )
    ap.add_argument(
        "--lo-timeout",
        type=int,
        default=600,
        help="LibreOffice check: convert timeout per run in seconds (default 600).",
    )
    ap.add_argument(
        "--lo-soffice",
        type=str,
        default=None,
        help="LibreOffice check: path or binary name for soffice (default: PATH).",
    )
    ap.add_argument(
        "--lo-keep-temps",
        action="store_true",
        help="LibreOffice check: keep temp dir; path is stored in JSON if check runs.",
    )
    ap.add_argument(
        "--lo-top-n",
        type=int,
        default=15,
        help="LibreOffice check: number of largest |Δ| cells to embed in JSON (default 15).",
    )
    ap.add_argument(
        "--lo-python-tolerance",
        type=float,
        default=None,
        metavar="EPS",
        help="If set, exit 1 when max|Python-LO| at --lo-shock-bps exceeds EPS (after successful LO run).",
    )
    args = ap.parse_args()
    lic_graph.WORKBOOK_PATH = args.workbook.resolve()

    t0 = time.perf_counter()
    graph = _load_graph(no_cache=bool(args.no_graph_cache))

    t1 = time.perf_counter()
    default_entry, shocks = _compute_entries(graph=graph)
    t2 = time.perf_counter()

    cache_path = lic_graph._default_graph_cache_path(lic_graph.WORKBOOK_PATH)
    payload: dict[str, Any] = {
        "schema": 1,
        "generated_at_unix_s": time.time(),
        "workbook_path": str(lic_graph.WORKBOOK_PATH),
        "graph_cache_path": str(cache_path),
        "bps_min": GDP_SHOCK_BPS_MIN,
        "bps_max": GDP_SHOCK_BPS_MAX,
        "default": _entry_to_json(default_entry),
        "shocks": [_entry_to_json(e) for e in shocks],
        "timing_s": {
            "graph_load": t1 - t0,
            "evaluate_all": t2 - t1,
            "total": t2 - t0,
        },
    }

    if args.libreoffice_check:
        from lic_dsf.libreoffice import (
            print_check_report,
            run_libreoffice_gdp_shock_check,
        )

        base_e = _cache_entry_for_bps(
            default_entry, shocks, int(args.lo_baseline_bps)
        )
        shock_e = _cache_entry_for_bps(
            default_entry, shocks, int(args.lo_shock_bps)
        )

        t_lo0 = time.perf_counter()
        lo_result = run_libreoffice_gdp_shock_check(
            lic_graph.WORKBOOK_PATH,
            baseline_bps=int(args.lo_baseline_bps),
            shock_bps=int(args.lo_shock_bps),
            timeout_s=int(args.lo_timeout),
            soffice=args.lo_soffice,
            keep_temps=bool(args.lo_keep_temps),
            top_n=int(args.lo_top_n),
            python_baseline_payload=base_e.payload,
            python_shock_payload=shock_e.payload,
        )
        t_lo1 = time.perf_counter()
        payload["timing_s"]["libreoffice_check"] = t_lo1 - t_lo0
        payload["timing_s"]["total"] = t_lo1 - t0
        payload["libreoffice_check"] = lo_result

        print()
        print_check_report(lo_result)
        if args.lo_python_tolerance is not None and lo_result.get("ok"):
            pvs = lo_result.get("python_vs_libreoffice")
            if isinstance(pvs, dict):
                sk = f"at_{int(args.lo_shock_bps)}_bps"
                block = pvs.get(sk)
                if isinstance(block, dict):
                    mx = block.get("max_abs_error")
                    if mx is not None and mx > float(args.lo_python_tolerance):
                        print(
                            f"Python vs LibreOffice max|error| at {sk} = {mx:.6g} "
                            f"> tolerance {args.lo_python_tolerance!r}; failing.",
                            file=sys.stderr,
                        )
                        args.out.parent.mkdir(parents=True, exist_ok=True)
                        args.out.write_text(
                            json.dumps(payload, ensure_ascii=False, indent=2),
                            encoding="utf-8",
                        )
                        print(f"Wrote {args.out} ({len(shocks)} shocks + default).")
                        sys.exit(1)
        if not lo_result.get("ok"):
            print("Precache JSON was still written; exit 1 due to failed LibreOffice check.", file=sys.stderr)
            args.out.parent.mkdir(parents=True, exist_ok=True)
            args.out.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            print(f"Wrote {args.out} ({len(shocks)} shocks + default).")
            sys.exit(1)

    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {args.out} ({len(shocks)} shocks + default).")


if __name__ == "__main__":
    main()
