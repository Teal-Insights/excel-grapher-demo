from __future__ import annotations

from pathlib import Path
from typing import Any

from .libreoffice import read_figure1_chart_values
from .payload import (
    CHART_SHEET,
    FIGURE1_PANELS,
    col_letters,
    read_category_labels_workbook,
)


def figure1_payload_from_chart_map(
    categories: list[str],
    chart_map: dict[str, float | None],
) -> dict[str, Any]:
    panels_out: list[dict[str, Any]] = []
    for panel in FIGURE1_PANELS:
        series_out: list[dict[str, Any]] = []
        for spec in panel.series:
            data = [chart_map[f"{CHART_SHEET}!{col}{spec.value_row}"] for col in col_letters()]
            series_out.append(
                {
                    "name": spec.legend.strip() or "(unlabeled)",
                    "data": data,
                    "borderColor": spec.color,
                    "borderDash": spec.dash,
                }
            )
        panels_out.append({"title": panel.title, "series": series_out})
    return {"categories": categories, "panels": panels_out}


def read_figure1_payload_from_workbook(path: Path) -> dict[str, Any]:
    categories = read_category_labels_workbook(path)
    chart_map = read_figure1_chart_values(path)
    return figure1_payload_from_chart_map(categories, chart_map)
