from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import fastpyxl
import numpy as np

from excel_grapher import (
    DependencyGraph,
    FormulaEvaluator,
    XlError,
    format_cell_key,
    get_calc_settings,
)

CHART_SHEET = "Chart Data"
YEAR_ROW = 35

# Baseline GDP forecast inputs, forecasts start at column X.
# Slider applies relative bps: new = baseline + baseline * (bps * 1e-4) = baseline * (1 + bps * 1e-4).
GDP_FORECAST_SHEET = "Input 3 - Macro-Debt data(DMX)"
GDP_FORECAST_ROWS = (12,)
GDP_FORECAST_START_COL = "X"
GDP_SHOCK_BPS_MIN = -10
GDP_SHOCK_BPS_MAX = 10

# Columns D:N (11 points), matching the workbook charts.
_VALUE_COL_INDICES = range(4, 15)


def col_letters() -> list[str]:
    return [fastpyxl.utils.cell.get_column_letter(i) for i in _VALUE_COL_INDICES]


@dataclass(frozen=True, slots=True)
class SeriesSpec:
    """value_row: Chart Data row for D:N series values (in the dependency graph)."""

    value_row: int
    legend: str
    color: str
    dash: list[int]


@dataclass(frozen=True, slots=True)
class PanelSpec:
    title: str
    series: tuple[SeriesSpec, ...]


# Mirrors ``xl/charts/chart17/16/18/15.xml`` on Output 2-1 Stress_Charts_Ex.
# Legends match Excel chart series titles (B/C column labels on Chart Data).
FIGURE1_PANELS: tuple[PanelSpec, ...] = (
    PanelSpec(
        title="PV of debt-to-GDP ratio",
        series=(
            SeriesSpec(61, "Baseline", "#4b82ad", []),
            SeriesSpec(62, "Historical scenario", "#ff0000", [10, 5]),
            SeriesSpec(63, "MX shock Standard&Tailored", "#000000", []),
            SeriesSpec(
                64,
                "MX value, 1 yr only shock Standard&Tailored - for chart",
                "#e46c0a",
                [],
            ),
            SeriesSpec(66, "Threshold", "#339966", [6, 4]),
            SeriesSpec(51, "Risk band", "#00ff00", []),
        ),
    ),
    PanelSpec(
        title="PV of debt-to-revenue ratio",
        series=(
            SeriesSpec(103, "Baseline", "#4b82ad", []),
            SeriesSpec(104, "Historical scenario", "#ff0000", [10, 5]),
            SeriesSpec(105, "MX shock Standard&Tailored", "#000000", []),
            SeriesSpec(
                106,
                "MX value, 1 yr only shock Standard&Tailored - for chart",
                "#f79646",
                [],
            ),
            SeriesSpec(108, "Threshold", "#339966", [6, 4]),
            SeriesSpec(93, "Risk band", "#00ff00", []),
        ),
    ),
    PanelSpec(
        title="Debt service-to-revenue ratio",
        series=(
            SeriesSpec(145, "Baseline", "#4b82ad", []),
            SeriesSpec(146, "Historical scenario", "#ff0000", [10, 5]),
            SeriesSpec(147, "MX shock Standard&Tailored", "#000000", []),
            SeriesSpec(
                148,
                "MX value, 1 yr only shock Standard&Tailored - for chart",
                "#f79646",
                [],
            ),
            SeriesSpec(150, "Threshold", "#339966", [6, 4]),
            SeriesSpec(135, "Risk band", "#00ff00", []),
        ),
    ),
    PanelSpec(
        title="Debt service-to-GDP ratio",
        series=(
            SeriesSpec(187, "Baseline", "#4b82ad", []),
            SeriesSpec(188, "Historical scenario", "#ff0000", [10, 5]),
            SeriesSpec(189, "MX shock Standard&Tailored", "#000000", []),
            SeriesSpec(
                190,
                "MX value, 1 yr only shock Standard&Tailored - for chart",
                "#f79646",
                [],
            ),
            SeriesSpec(192, "Threshold", "#339966", [6, 4]),
            SeriesSpec(177, "Risk band", "#00ff00", []),
        ),
    ),
)


def _gdp_forecast_start_col_idx() -> int:
    return fastpyxl.utils.cell.column_index_from_string(GDP_FORECAST_START_COL)


def _read_gdp_forecast_cell_values_from_workbook(
    workbook_path,
) -> tuple[list[str], list[float | None]] | None:
    """
    Read baseline GDP forecast values from the workbook itself.

    Returns:
        (keys, values) where keys are sheet-qualified cell keys in the same order
        as values. Values may include None for blank/non-numeric cells.
    """
    if not getattr(workbook_path, "is_file", None) or not workbook_path.is_file():
        return None

    wb = fastpyxl.load_workbook(workbook_path, data_only=True, keep_vba=False)
    try:
        if GDP_FORECAST_SHEET not in wb.sheetnames:
            return None
        ws = wb[GDP_FORECAST_SHEET]

        start_idx = _gdp_forecast_start_col_idx()
        # Scan right until we hit a run of blank columns for both rows.
        blank_run = 0
        max_scan_cols = 512

        keys: list[str] = []
        vals: list[float | None] = []
        for col_idx in range(start_idx, start_idx + max_scan_cols):
            col = fastpyxl.utils.cell.get_column_letter(col_idx)
            row_vals: list[float | None] = []
            for r in GDP_FORECAST_ROWS:
                v = ws[f"{col}{r}"].value
                row_vals.append(numeric_scalar(v))

            if all(v is None for v in row_vals):
                blank_run += 1
            else:
                blank_run = 0

            for r, v in zip(GDP_FORECAST_ROWS, row_vals, strict=True):
                keys.append(format_cell_key(GDP_FORECAST_SHEET, col, r))
                vals.append(v)

            if blank_run >= 5:
                break

        return keys, vals
    finally:
        wb.close()


def gdp_forecast_cell_keys(graph: DependencyGraph, *, workbook_path) -> list[str] | None:
    """
    GDP forecast input keys that exist in the dependency graph.

    We discover the intended X..end range from the workbook layout, then filter
    to keys that are present in the graph so FormulaEvaluator can accept them.
    """
    wb_read = _read_gdp_forecast_cell_values_from_workbook(workbook_path)
    if wb_read is None:
        return None
    keys, _vals = wb_read
    present = [k for k in keys if graph.get_node(k) is not None]
    return present or None


def gdp_forecast_baselines(
    graph: DependencyGraph, *, workbook_path, keys: list[str]
) -> list[float]:
    """
    Baseline values aligned to `keys`.

    Preference order:
    - Use graph node cached values (fast; matches graph build snapshot)
    - Fall back to reading from workbook (data_only)
    - Default to 0.0 if missing/non-numeric
    """
    out: list[float | None] = []
    for k in keys:
        n = graph.get_node(k)
        out.append(numeric_scalar(n.value if n else None))

    if any(v is None for v in out):
        wb_read = _read_gdp_forecast_cell_values_from_workbook(workbook_path)
        if wb_read is not None:
            wb_keys, wb_vals = wb_read
            wb_map = {k: v for k, v in zip(wb_keys, wb_vals, strict=True)}
            out = [
                v if v is not None else wb_map.get(k)
                for k, v in zip(keys, out, strict=True)
            ]

    return [float(v) if v is not None else 0.0 for v in out]


def gdp_forecast_value_from_bps(baseline: float, bps: int) -> float:
    return baseline + baseline * (bps * 1e-4)


def cell_key(col: str, row: int) -> str:
    return format_cell_key(CHART_SHEET, col, row)


def category_keys() -> list[str]:
    return [cell_key(c, YEAR_ROW) for c in col_letters()]


def read_category_labels_workbook(workbook_path) -> list[str]:
    wb = fastpyxl.load_workbook(workbook_path, data_only=True, keep_vba=False)
    try:
        ws = wb[CHART_SHEET]
        labels: list[str] = []
        for col in col_letters():
            v = ws[f"{col}{YEAR_ROW}"].value
            labels.append(text_scalar(v))
        return labels
    finally:
        wb.close()


def text_scalar(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, XlError):
        return v.value
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def numeric_scalar(v: Any) -> float | None:
    if v is None:
        return None
    if isinstance(v, XlError):
        return None
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v)
        except ValueError:
            return None
    return None


def build_figure1_payload(
    graph: DependencyGraph,
    *,
    workbook_path,
    evaluator: FormulaEvaluator | None = None,
) -> dict[str, Any]:
    settings = get_calc_settings(workbook_path)
    ev = evaluator or FormulaEvaluator(
        graph,
        iterate_enabled=settings.iterate_enabled,
        iterate_count=settings.iterate_count,
        iterate_delta=settings.iterate_delta,
    )

    cat_keys = category_keys()
    if not all(graph.get_node(k) for k in cat_keys):
        categories = read_category_labels_workbook(workbook_path)
    else:
        cat_vals = ev.evaluate(cat_keys)
        categories = [text_scalar(cat_vals[k]) for k in cat_keys]

    value_cols = col_letters()

    all_keys: list[str] = []
    for panel in FIGURE1_PANELS:
        for s in panel.series:
            for col in value_cols:
                all_keys.append(cell_key(col, s.value_row))

    missing = [k for k in all_keys if graph.get_node(k) is None]
    if missing:
        sample = ", ".join(missing[:5])
        raise KeyError(
            f"{len(missing)} chart value cells are missing from the graph "
            f"(rebuild: uv run python scripts/extract_graph.py --no-cache). Examples: {sample}"
        )

    evaluated = ev.evaluate(all_keys)

    panels_out: list[dict[str, Any]] = []
    for panel in FIGURE1_PANELS:
        series_out: list[dict[str, Any]] = []
        for s in panel.series:
            name = s.legend.strip() or "(unlabeled)"
            ys = [
                numeric_scalar(evaluated[cell_key(col, s.value_row)]) for col in value_cols
            ]
            series_out.append(
                {
                    "name": name,
                    "data": ys,
                    "borderColor": s.color,
                    "borderDash": s.dash,
                }
            )
        panels_out.append({"title": panel.title, "series": series_out})

    return {"categories": categories, "panels": panels_out}
