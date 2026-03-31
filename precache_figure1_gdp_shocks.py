#!/usr/bin/env python3
"""
Precompute Figure 1 outputs for the GDP forecast "shock" slider.

This script evaluates Figure 1 once at the default (0 bps), plus all non-zero
shock settings from GDP_SHOCK_BPS_MIN..GDP_SHOCK_BPS_MAX, and writes the results
to a JSON file for fast, pre-cached loading in the webapp.

Run:
  uv run python precache_figure1_gdp_shocks.py
  uv run python precache_figure1_gdp_shocks.py --out .cache/figure1-gdp-shocks.json --with-svg
"""

from __future__ import annotations

import argparse
import json
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import extract_graph as lic_graph
from excel_grapher import DependencyGraph, DynamicRefError, FormulaEvaluator, create_dependency_graph
from figure1_data import (
    GDP_SHOCK_BPS_MAX,
    GDP_SHOCK_BPS_MIN,
    build_figure1_payload,
    gdp_forecast_baselines,
    gdp_forecast_cell_keys,
    gdp_forecast_value_from_bps,
)
from ssr_charts import render_figure_svg_panels


@dataclass(frozen=True, slots=True)
class CacheEntry:
    bps: int
    payload: dict[str, Any]
    svg_panels: list[dict[str, str]] | None


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
            "Run `uv run python extract_graph.py --no-cache`, add dynamic-ref constraints, then retry. "
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


def _compute_entries(*, graph: DependencyGraph, with_svg: bool) -> tuple[CacheEntry, list[CacheEntry]]:
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
        payload = build_figure1_payload(graph, workbook_path=lic_graph.WORKBOOK_PATH, evaluator=ev)
        svg_panels = render_figure_svg_panels(payload) if with_svg else None
        return CacheEntry(bps=bps, payload=payload, svg_panels=svg_panels)

    default_entry = eval_at_bps(0)

    shocks: list[CacheEntry] = []
    for bps in range(GDP_SHOCK_BPS_MIN, GDP_SHOCK_BPS_MAX + 1):
        if bps == 0:
            continue
        shocks.append(eval_at_bps(int(bps)))

    return default_entry, shocks


def _entry_to_json(e: CacheEntry) -> dict[str, Any]:
    out: dict[str, Any] = {"bps": e.bps, "payload": e.payload}
    if e.svg_panels is not None:
        out["svg_panels"] = e.svg_panels
    return out


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Precompute Figure 1 outputs for all GDP shock slider states."
    )
    ap.add_argument(
        "--out",
        type=Path,
        default=Path(".cache/figure1-gdp-shocks.json"),
        help="Output JSON path (default: .cache/figure1-gdp-shocks.json).",
    )
    ap.add_argument(
        "--no-graph-cache",
        action="store_true",
        help="Ignore graph pickle cache and rebuild the dependency graph from workbook.",
    )
    ap.add_argument(
        "--with-svg",
        action="store_true",
        help="Also pre-render SVG for each panel and store in JSON (larger file, faster UI).",
    )
    args = ap.parse_args()

    t0 = time.perf_counter()
    graph = _load_graph(no_cache=bool(args.no_graph_cache))

    t1 = time.perf_counter()
    default_entry, shocks = _compute_entries(graph=graph, with_svg=bool(args.with_svg))

    cache_path = lic_graph._default_graph_cache_path(lic_graph.WORKBOOK_PATH)
    payload = {
        "schema": 1,
        "generated_at_unix_s": time.time(),
        "workbook_path": str(lic_graph.WORKBOOK_PATH),
        "graph_cache_path": str(cache_path),
        "bps_min": GDP_SHOCK_BPS_MIN,
        "bps_max": GDP_SHOCK_BPS_MAX,
        "with_svg": bool(args.with_svg),
        "default": _entry_to_json(default_entry),
        "shocks": [_entry_to_json(e) for e in shocks],
        "timing_s": {
            "graph_load": t1 - t0,
            "evaluate_all": time.perf_counter() - t1,
            "total": time.perf_counter() - t0,
        },
    }

    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {args.out} ({len(shocks)} shocks + default).")


if __name__ == "__main__":
    main()

