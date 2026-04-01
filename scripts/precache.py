#!/usr/bin/env python3
"""
Precompute chart outputs for the GDP forecast shock slider.

This script evaluates the chart payload once at the default (0% shock), plus every
other level from GDP_SHOCK_PCT_MIN..GDP_SHOCK_PCT_MAX in GDP_SHOCK_PCT_STEP increments
(forecast cells set to baseline * (1 + pct/100)), and writes the results to JSON for
the webapp slider.

Run:
  uv run python scripts/precache.py
  uv run python scripts/precache.py --workbook lic-dsf-template-2025-08-12.xlsm
  uv run python scripts/precache.py --sanity-check-backend auto
  uv run python scripts/precache.py --sanity-check

From the repo root, ``lic_dsf`` is on ``sys.path`` automatically; you can still set
``PYTHONPATH=.`` if you prefer.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

_REPO_ROOT = Path(__file__).resolve().parents[1]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from excel_grapher import DependencyGraph, DynamicRefError, FormulaEvaluator, create_dependency_graph

from lic_dsf import graph as lic_graph
from lic_dsf.payload import (
    GDP_SHOCK_PCT_MAX,
    GDP_SHOCK_PCT_MIN,
    GDP_SHOCK_PCT_STEP,
    build_figure1_payload,
    gdp_forecast_baselines,
    gdp_forecast_cell_keys,
    gdp_forecast_value_from_percent,
    gdp_shock_percent_levels,
)
from lic_dsf.workbook_backend import resolve_backend


def _norm_pct(pct: float) -> float:
    return round(float(pct), 6)


@dataclass(frozen=True, slots=True)
class CacheEntry:
    pct: float
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


def _compute_formula_evaluator_entries(
    *, graph: DependencyGraph
) -> tuple[CacheEntry, list[CacheEntry]]:
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

    def eval_at_pct(pct: float) -> CacheEntry:
        for k, base in zip(keys, baselines, strict=True):
            ev.set_value(k, gdp_forecast_value_from_percent(float(base), pct))
        payload = build_figure1_payload(
            graph,
            workbook_path=lic_graph.WORKBOOK_PATH,
            evaluator=ev,
        )
        return CacheEntry(pct=pct, payload=payload)

    default_entry = eval_at_pct(0.0)

    shocks: list[CacheEntry] = []
    for pct in gdp_shock_percent_levels():
        if abs(pct) < 1e-9:
            continue
        shocks.append(eval_at_pct(pct))

    return default_entry, shocks


def _entry_to_json(e: CacheEntry) -> dict[str, Any]:
    return {"pct": e.pct, "payload": e.payload}


def _cache_entry_for_pct(
    default_entry: CacheEntry, shocks: list[CacheEntry], pct: float
) -> CacheEntry:
    target = _norm_pct(pct)
    if abs(target) < 1e-9:
        return default_entry
    for e in shocks:
        if _norm_pct(e.pct) == target:
            return e
    raise SystemExit(
        f"No precache entry for pct={pct} (expected one of the GDP shock percent levels)."
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
        "--sanity-check",
        "--libreoffice-check",
        dest="sanity_check",
        action="store_true",
        help="After caching, run the workbook-engine sanity check. `--libreoffice-check` remains as a compatibility alias.",
    )
    ap.add_argument(
        "--sanity-check-backend",
        choices=("auto", "libreoffice", "xlwings"),
        default="auto",
        help="Sanity-check engine: auto picks xlwings on Windows and LibreOffice on Linux.",
    )
    ap.add_argument(
        "--lo-baseline-pct",
        type=float,
        default=0.0,
        help="Sanity check: reference copy GDP shock in %% (default 0).",
    )
    ap.add_argument(
        "--lo-shock-pct",
        type=float,
        default=1.0,
        help="Sanity check: comparison copy GDP shock in %% (default 1).",
    )
    ap.add_argument(
        "--lo-timeout",
        type=int,
        default=600,
        help="Sanity check: LibreOffice convert timeout per run in seconds (default 600).",
    )
    ap.add_argument(
        "--lo-soffice",
        type=str,
        default=None,
        help="Sanity check: path or binary name for soffice (default: PATH).",
    )
    ap.add_argument(
        "--lo-keep-temps",
        action="store_true",
        help="Sanity check: keep temp dir; path is stored in JSON if check runs.",
    )
    ap.add_argument(
        "--lo-top-n",
        type=int,
        default=15,
        help="Sanity check: number of largest |Δ| cells to embed in JSON (default 15).",
    )
    ap.add_argument(
        "--lo-python-tolerance",
        type=float,
        default=None,
        metavar="EPS",
        help="If set, exit 1 when max|Python-backend| at --lo-shock-pct exceeds EPS.",
    )
    args = ap.parse_args()
    lic_graph.WORKBOOK_PATH = args.workbook.resolve()

    t0 = time.perf_counter()
    graph = _load_graph(no_cache=bool(args.no_graph_cache))
    t1 = time.perf_counter()
    default_entry, shocks = _compute_formula_evaluator_entries(graph=graph)
    t2 = time.perf_counter()

    cache_path = lic_graph._default_graph_cache_path(lic_graph.WORKBOOK_PATH)
    payload: dict[str, Any] = {
        "schema": 2,
        "generated_at_unix_s": time.time(),
        "workbook_path": str(lic_graph.WORKBOOK_PATH),
        "graph_cache_path": str(cache_path),
        "pct_min": GDP_SHOCK_PCT_MIN,
        "pct_max": GDP_SHOCK_PCT_MAX,
        "pct_step": GDP_SHOCK_PCT_STEP,
        "default": _entry_to_json(default_entry),
        "shocks": [_entry_to_json(e) for e in shocks],
        "timing_s": {
            "graph_load": t1 - t0,
            "evaluate_all": t2 - t1,
            "total": t2 - t0,
        },
    }

    if args.sanity_check:
        from lic_dsf.workbook_backend import (
            print_check_report,
            run_workbook_gdp_shock_check,
        )

        base_e = _cache_entry_for_pct(
            default_entry, shocks, float(args.lo_baseline_pct)
        )
        shock_e = _cache_entry_for_pct(
            default_entry, shocks, float(args.lo_shock_pct)
        )

        t_lo0 = time.perf_counter()
        check_backend = resolve_backend(args.sanity_check_backend)
        lo_result = run_workbook_gdp_shock_check(
            lic_graph.WORKBOOK_PATH,
            backend=check_backend,
            baseline_pct=float(args.lo_baseline_pct),
            shock_pct=float(args.lo_shock_pct),
            timeout_s=int(args.lo_timeout),
            soffice=args.lo_soffice,
            keep_temps=bool(args.lo_keep_temps),
            top_n=int(args.lo_top_n),
            python_baseline_payload=base_e.payload,
            python_shock_payload=shock_e.payload,
        )
        t_lo1 = time.perf_counter()
        payload["timing_s"]["sanity_check"] = t_lo1 - t_lo0
        payload["timing_s"]["total"] = t_lo1 - t0
        payload["sanity_check"] = lo_result

        print()
        print_check_report(lo_result)
        if args.lo_python_tolerance is not None and lo_result.get("ok"):
            pvs = lo_result.get("python_vs_backend")
            if isinstance(pvs, dict):
                sk = f"at_{_norm_pct(float(args.lo_shock_pct)):g}_pct"
                block = pvs.get(sk)
                if isinstance(block, dict):
                    mx = block.get("max_abs_error")
                    if mx is not None and mx > float(args.lo_python_tolerance):
                        print(
                            f"Python vs workbook backend max|error| at {sk} = {mx:.6g} "
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
            print(
                "Precache JSON was still written; exit 1 due to failed workbook sanity check.",
                file=sys.stderr,
            )
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
