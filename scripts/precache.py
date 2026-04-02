#!/usr/bin/env python3
"""
Precompute chart outputs for the GDP forecast shock slider.

This script evaluates the chart payload once at the default (0% shock), plus every
other level from GDP_SHOCK_PCT_MIN..GDP_SHOCK_PCT_MAX in GDP_SHOCK_PCT_STEP increments
(forecast growth rates shocked by pct/100 and then rebuilt into forecast levels),
and writes the results to JSON for the webapp slider.

Run:
  uv run python scripts/precache.py
  uv run python scripts/precache.py --backend xlwings
  uv run python scripts/precache.py --backend libreoffice --soffice soffice
  uv run python scripts/precache.py --workbook lic-dsf-template-2025-08-12.xlsm

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
from typing import Any, Callable

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
    gdp_forecast_values_from_percent,
    gdp_shock_percent_levels,
)


@dataclass(frozen=True, slots=True)
class CacheEntry:
    pct: float
    payload: dict[str, Any]


def _default_out_path(backend: str) -> Path:
    return Path(".cache") / f"gdp-shocks-{backend}.json"


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
        shocked_series = gdp_forecast_values_from_percent(baselines, pct)
        for k, value in zip(keys, shocked_series, strict=True):
            ev.set_value(k, value)
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


def _compute_entries_from_payload_builder(
    build_payload_for_pct: Callable[[float], dict[str, Any]],
) -> tuple[CacheEntry, list[CacheEntry]]:
    default_entry = CacheEntry(pct=0.0, payload=build_payload_for_pct(0.0))
    shocks: list[CacheEntry] = []
    for pct in gdp_shock_percent_levels():
        if abs(pct) < 1e-9:
            continue
        shocks.append(CacheEntry(pct=pct, payload=build_payload_for_pct(pct)))
    return default_entry, shocks


def _compute_workbook_backend_entries(
    *,
    backend: str,
    workbook: Path,
    timeout_s: int,
    soffice: str | None,
) -> tuple[CacheEntry, list[CacheEntry]]:
    if backend == "libreoffice":
        from lic_dsf.libreoffice_backend import recalculate_figure1_payload_with_libreoffice

        return _compute_entries_from_payload_builder(
            lambda pct: recalculate_figure1_payload_with_libreoffice(
                workbook,
                pct=pct,
                timeout_s=timeout_s,
                soffice=soffice,
            )
        )

    if backend == "xlwings":
        from lic_dsf.xlwings_backend import recalculate_figure1_payload_with_xlwings

        return _compute_entries_from_payload_builder(
            lambda pct: recalculate_figure1_payload_with_xlwings(workbook, pct=pct)
        )

    raise ValueError(f"Unknown backend: {backend}")


def _entry_to_json(e: CacheEntry) -> dict[str, Any]:
    return {"pct": e.pct, "payload": e.payload}


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Precompute chart outputs for all GDP shock slider states."
    )
    ap.add_argument(
        "--backend",
        choices=("excel-grapher", "xlwings", "libreoffice"),
        default="excel-grapher",
        help=(
            "Calculation backend used to build the cache. "
            "Default: excel-grapher."
        ),
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
        default=None,
        help="Output JSON path (default: .cache/gdp-shocks-<backend>.json).",
    )
    ap.add_argument(
        "--no-graph-cache",
        action="store_true",
        help=(
            "Ignore graph pickle cache and rebuild the dependency graph from workbook. "
            "Only used with --backend excel-grapher."
        ),
    )
    ap.add_argument(
        "--timeout",
        type=int,
        default=600,
        help="LibreOffice convert timeout per run in seconds (default 600).",
    )
    ap.add_argument(
        "--soffice",
        type=str,
        default=None,
        help="Path or binary name for soffice/libreoffice (default: search PATH).",
    )
    args = ap.parse_args()
    lic_graph.WORKBOOK_PATH = args.workbook.resolve()
    if args.backend != "excel-grapher" and args.no_graph_cache:
        raise SystemExit("--no-graph-cache is only valid with --backend excel-grapher.")
    out_path = args.out.resolve() if args.out else _default_out_path(args.backend).resolve()

    t0 = time.perf_counter()
    timing_s: dict[str, float] = {}
    if args.backend == "excel-grapher":
        graph = _load_graph(no_cache=bool(args.no_graph_cache))
        t1 = time.perf_counter()
        default_entry, shocks = _compute_formula_evaluator_entries(graph=graph)
        t2 = time.perf_counter()
        timing_s["graph_load"] = t1 - t0
        timing_s["evaluate_all"] = t2 - t1
    else:
        default_entry, shocks = _compute_workbook_backend_entries(
            backend=args.backend,
            workbook=lic_graph.WORKBOOK_PATH,
            timeout_s=int(args.timeout),
            soffice=args.soffice,
        )
        t2 = time.perf_counter()
        timing_s["evaluate_all"] = t2 - t0

    cache_path = lic_graph._default_graph_cache_path(lic_graph.WORKBOOK_PATH)
    payload: dict[str, Any] = {
        "schema": 3,
        "backend": args.backend,
        "generated_at_unix_s": time.time(),
        "workbook_path": str(lic_graph.WORKBOOK_PATH),
        "graph_cache_path": (
            str(cache_path) if args.backend == "excel-grapher" else None
        ),
        "pct_min": GDP_SHOCK_PCT_MIN,
        "pct_max": GDP_SHOCK_PCT_MAX,
        "pct_step": GDP_SHOCK_PCT_STEP,
        "default": _entry_to_json(default_entry),
        "shocks": [_entry_to_json(e) for e in shocks],
        "timing_s": {**timing_s, "total": t2 - t0},
    }
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {out_path} ({len(shocks)} shocks + default).")


if __name__ == "__main__":
    main()
