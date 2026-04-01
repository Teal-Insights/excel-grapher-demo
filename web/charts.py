from __future__ import annotations

import html
import json
import math
from pathlib import Path
from typing import Any, Iterable


def load_shock_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _norm_pct(v: Any) -> float:
    return round(float(v), 6)


def payloads_by_shock(cache_doc: dict[str, Any]) -> dict[float, dict[str, Any]]:
    """
    Map shock level -> Figure 1 payload. New caches use ``pct`` (percent);
    schema 1 used integer ``bps`` (basis points).
    """
    out: dict[float, dict[str, Any]] = {}
    default = cache_doc.get("default")
    if isinstance(default, dict) and "payload" in default:
        pl = default["payload"]
        if isinstance(pl, dict):
            if "pct" in default:
                out[_norm_pct(default["pct"])] = pl
            elif "bps" in default:
                out[float(int(default["bps"]))] = pl
    for entry in cache_doc.get("shocks") or []:
        if not isinstance(entry, dict) or "payload" not in entry:
            continue
        pl = entry["payload"]
        if not isinstance(pl, dict):
            continue
        if "pct" in entry:
            out[_norm_pct(entry["pct"])] = pl
        elif "bps" in entry:
            out[float(int(entry["bps"]))] = pl
    return out


def _escape(text: str) -> str:
    return html.escape(text, quote=True)


def _bool_attr(value: Any) -> str:
    return "true" if bool(value) else "false"


def _panel_shock_label(panel: dict[str, Any]) -> str:
    return str(panel.get("mostExtremeShockLabel", "") or "").strip()


def _panel_breach_count(panel: dict[str, Any], key: str) -> str:
    value = panel.get(key)
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


_FOCAL_SERIES_NAMES = frozenset(
    {
        "Baseline",
        "Historical scenario",
        "MX shock Standard&Tailored",
        "Threshold",
    }
)
_HIDDEN_SERIES_NAMES = frozenset({"Risk band"})


def _series_name(series: dict[str, Any]) -> str:
    return str(series.get("name", "") or "")


def _visible_series(series: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [s for s in series if _series_name(s) not in _HIDDEN_SERIES_NAMES]


def _series_is_focal(series: dict[str, Any]) -> bool:
    explicit = series.get("isFocal")
    if explicit is not None:
        return bool(explicit)
    return _series_name(series) in _FOCAL_SERIES_NAMES


def _series_in_paint_order(series: list[dict[str, Any]]) -> list[dict[str, Any]]:
    indexed = list(enumerate(_visible_series(series)))
    indexed.sort(key=lambda item: (1 if _series_is_focal(item[1]) else 0, item[0]))
    return [series_item for _, series_item in indexed]


def _finite_series_values(series: list[dict[str, Any]]) -> list[float]:
    out: list[float] = []
    for s in series:
        for v in s.get("data") or []:
            if v is None:
                continue
            if isinstance(v, float) and math.isnan(v):
                continue
            try:
                out.append(float(v))
            except (TypeError, ValueError):
                continue
    return out


def _y_domain(series_list: list[dict[str, Any]]) -> tuple[float, float]:
    vals = _finite_series_values(series_list)
    if not vals:
        return 0.0, 1.0
    lo, hi = min(vals), max(vals)
    if math.isclose(lo, hi, rel_tol=0.0, abs_tol=1e-12):
        pad = 1.0 if lo == 0 else abs(lo) * 0.05
        return lo - pad, hi + pad
    span = hi - lo
    margin = span * 0.06
    return lo - margin, hi + margin


def _segments(ys: list[Any]) -> list[list[tuple[int, float]]]:
    runs: list[list[tuple[int, float]]] = []
    run: list[tuple[int, float]] = []
    for i, raw in enumerate(ys):
        if raw is None:
            if run:
                runs.append(run)
                run = []
            continue
        if isinstance(raw, float) and math.isnan(raw):
            if run:
                runs.append(run)
                run = []
            continue
        try:
            yv = float(raw)
        except (TypeError, ValueError):
            if run:
                runs.append(run)
                run = []
            continue
        run.append((i, yv))
    if run:
        runs.append(run)
    return runs


def _dash_attr(dash: list[int] | None) -> str:
    if not dash:
        return ""
    return f' stroke-dasharray="{" ".join(str(int(x)) for x in dash)}"'


def _x_px(i: int, n: int, margin_left: float, plot_w: float) -> float:
    if n <= 1:
        return margin_left + plot_w / 2.0
    return margin_left + (i / (n - 1)) * plot_w


def _y_px(v: float, y0: float, y1: float, margin_top: float, plot_h: float) -> float:
    t = (v - y0) / (y1 - y0) if y1 != y0 else 0.5
    return margin_top + (1.0 - t) * plot_h


def _render_panel_static(
    *,
    categories: list[str],
    series_for_legend: list[dict[str, Any]],
    y0: float,
    y1: float,
    width: int = 520,
    height: int = 260,
    margin_left: float = 42.0,
    margin_right: float = 12.0,
    margin_top: float = 28.0,
    margin_bottom: float = 52.0,
) -> str:
    n = len(categories)
    if n == 0:
        return '<text x="10" y="24" font-size="12" fill="#666">No categories</text>'

    plot_w = float(width) - margin_left - margin_right
    plot_h = float(height) - margin_top - margin_bottom

    parts: list[str] = [
        f'<rect x="0" y="0" width="{width}" height="{height}" fill="white"/>',
    ]

    grid_lines = 5
    for g in range(grid_lines + 1):
        gv = y0 + (y1 - y0) * (g / grid_lines)
        gy = _y_px(gv, y0, y1, margin_top, plot_h)
        parts.append(
            f'<line x1="{margin_left:.2f}" y1="{gy:.2f}" x2="{margin_left + plot_w:.2f}" y2="{gy:.2f}" '
            f'stroke="#e8e9eb" stroke-width="1"/>'
        )
        parts.append(
            f'<text x="{margin_left - 4:.2f}" y="{gy + 3:.2f}" font-size="8" fill="#666" text-anchor="end">'
            f"{_escape(f'{gv:.1f}')}</text>"
        )

    parts.append(
        f'<rect x="{margin_left:.2f}" y="{margin_top:.2f}" width="{plot_w:.2f}" height="{plot_h:.2f}" '
        f'fill="none" stroke="#cccccc" stroke-width="1"/>'
    )

    for i, cat in enumerate(categories):
        cx = _x_px(i, n, margin_left, plot_w)
        parts.append(
            f'<text x="{cx:.2f}" y="{margin_top + plot_h + 14:.2f}" font-size="8" fill="#666" '
            f'text-anchor="middle">{_escape(str(cat))}</text>'
        )

    leg_y = margin_top + plot_h + 34
    leg_x0 = margin_left
    legend_items: Iterable[tuple[str, str, str]] = (
        (
            _series_name(s),
            str(s.get("borderColor") or "#000000"),
            _bool_attr(_series_is_focal(s)),
        )
        for s in _visible_series(series_for_legend)
    )
    items = [(name, color, focal) for name, color, focal in legend_items if name]
    col_w = plot_w / 2.0
    for idx, (name, color, focal) in enumerate(items):
        col = idx % 2
        row = idx // 2
        lx = leg_x0 + col * col_w
        ly = leg_y + row * 12
        parts.append(
            f'<line x1="{lx:.2f}" y1="{ly:.2f}" x2="{lx + 14:.2f}" y2="{ly:.2f}" '
            f'class="legend-swatch" data-series-name="{_escape(name)}" data-focal="{focal}" '
            f'stroke="{_escape(color)}" stroke-width="2"/>'
        )
        parts.append(
            f'<text x="{lx + 18:.2f}" y="{ly + 3:.2f}" font-size="8" fill="#333">{_escape(name)}</text>'
        )

    return "".join(parts)


def _render_shock_polylines(
    *,
    series: list[dict[str, Any]],
    categories: list[str],
    y0: float,
    y1: float,
    width: int = 520,
    height: int = 260,
    margin_left: float = 42.0,
    margin_right: float = 12.0,
    margin_top: float = 28.0,
    margin_bottom: float = 52.0,
) -> str:
    n = len(categories)
    if n == 0:
        return ""

    plot_w = float(width) - margin_left - margin_right
    plot_h = float(height) - margin_top - margin_bottom
    parts: list[str] = []
    for s in _series_in_paint_order(series):
        name = _series_name(s)
        color = str(s.get("borderColor") or "#000000")
        focal = _bool_attr(_series_is_focal(s))
        dash = s.get("borderDash") or []
        ys = list(s.get("data") or [])
        d_attr = _dash_attr(dash if isinstance(dash, list) else [])
        for seg in _segments(ys):
            pts = " ".join(
                f"{_x_px(i, n, margin_left, plot_w):.2f},{_y_px(v, y0, y1, margin_top, plot_h):.2f}"
                for i, v in seg
            )
            parts.append(
                f'<polyline class="shock-line" data-series-name="{_escape(name)}" '
                f'data-focal="{focal}" fill="none" stroke="{_escape(color)}" '
                f'stroke-width="1.6" pointer-events="none" points="{pts}"{d_attr}/>'
            )
    return "".join(parts)


def _render_panel_group(
    *,
    title: str,
    categories: list[str],
    series: list[dict[str, Any]],
    width: int = 520,
    height: int = 260,
    margin_left: float = 42.0,
    margin_right: float = 12.0,
    margin_top: float = 28.0,
    margin_bottom: float = 52.0,
) -> str:
    """Single-shock full panel (grid + legend + colored series)."""
    y0, y1 = _y_domain(series)
    static = _render_panel_static(
        categories=categories,
        series_for_legend=series,
        y0=y0,
        y1=y1,
        width=width,
        height=height,
        margin_left=margin_left,
        margin_right=margin_right,
        margin_top=margin_top,
        margin_bottom=margin_bottom,
    )
    poly = _render_panel_group_polylines_colored(
        series=series,
        categories=categories,
        y0=y0,
        y1=y1,
        width=width,
        height=height,
        margin_left=margin_left,
        margin_right=margin_right,
        margin_top=margin_top,
        margin_bottom=margin_bottom,
    )
    return static + poly


def _render_panel_group_polylines_colored(
    *,
    series: list[dict[str, Any]],
    categories: list[str],
    y0: float,
    y1: float,
    width: int = 520,
    height: int = 260,
    margin_left: float = 42.0,
    margin_right: float = 12.0,
    margin_top: float = 28.0,
    margin_bottom: float = 52.0,
) -> str:
    n = len(categories)
    if n == 0:
        return ""
    plot_w = float(width) - margin_left - margin_right
    plot_h = float(height) - margin_top - margin_bottom
    parts: list[str] = []
    for s in _series_in_paint_order(series):
        color = str(s.get("borderColor") or "#000000")
        dash = s.get("borderDash") or []
        ys = list(s.get("data") or [])
        d_attr = _dash_attr(dash if isinstance(dash, list) else [])
        for seg in _segments(ys):
            pts = " ".join(
                f"{_x_px(i, n, margin_left, plot_w):.2f},{_y_px(v, y0, y1, margin_top, plot_h):.2f}"
                for i, v in seg
            )
            parts.append(
                f'<polyline fill="none" stroke="{_escape(color)}" stroke-width="1.6" '
                f'points="{pts}"{d_attr}/>'
            )
    return "".join(parts)


def _all_series_for_panel_domain(
    by_shock: dict[float, dict[str, Any]],
    shock_list: list[float],
    panel_idx: int,
) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for shock in shock_list:
        pl = by_shock[shock]
        panels = list(pl.get("panels") or [])
        if panel_idx < len(panels):
            out.extend(_visible_series(list(panels[panel_idx].get("series") or [])))
    return out


def build_chart_html(cache_doc: dict[str, Any]) -> str:
    """
    One <svg> per panel: a static <g class="panel-static"> (grid, axes, legend) plus one
    <g class="shock-layer" data-pct="…"> per shock containing only polylines. All layers
    are visible; CSS grays out non-selected shocks and keeps selected focal series at
    their legend colors.
    """
    by_shock = payloads_by_shock(cache_doc)
    if not by_shock:
        return '<p class="err">Cache has no chart entries.</p>'

    shock_list = sorted(by_shock.keys())
    ref0 = by_shock.get(0.0) or by_shock[shock_list[0]]
    categories0 = list(ref0.get("categories") or [])
    panels0 = list(ref0.get("panels") or [])
    default_shock = next((s for s in shock_list if abs(s) < 1e-9), shock_list[0])

    chunks: list[str] = []
    for panel_idx, ref_panel in enumerate(panels0):
        title = str(ref_panel.get("title") or f"Panel {panel_idx}")
        title_esc = html.escape(title)
        ref_series = list(ref_panel.get("series") or [])
        combined = _all_series_for_panel_domain(by_shock, shock_list, panel_idx)
        y0, y1 = _y_domain(combined)

        chunks.append('<div class="card">')
        chunks.append(f"<h2>{title_esc}</h2>")
        chunks.append('<div class="panel-meta">')
        for shock in shock_list:
            payload = by_shock[shock]
            panels = list(payload.get("panels") or [])
            if panel_idx >= len(panels):
                continue
            panel = panels[panel_idx]
            pct_attr = f"{shock:g}"
            sel = abs(shock - default_shock) < 1e-5
            sel_class = " is-selected" if sel else ""
            shock_label = _panel_shock_label(panel)
            baseline_breaches = _panel_breach_count(panel, "baselineBreaches")
            shock_breaches = _panel_breach_count(panel, "shockBreaches")
            chunks.append(f'<div class="panel-meta-entry{sel_class}" data-pct="{pct_attr}">')
            if shock_label:
                chunks.append(
                    f'<p class="panel-shock-label" aria-label="Most extreme shock">{html.escape(shock_label)}</p>'
                )
            if baseline_breaches or shock_breaches:
                chunks.append('<p class="panel-breach-counts">')
                chunks.append(
                    f'Baseline breaches: <span class="count">{html.escape(baseline_breaches or "-")}</span>'
                )
                chunks.append(
                    f' | Shock breaches: <span class="count">{html.escape(shock_breaches or "-")}</span>'
                )
                chunks.append("</p>")
            chunks.append("</div>")
        chunks.append("</div>")
        chunks.append('<div class="chart-wrap">')
        chunks.append(
            f'<svg xmlns="http://www.w3.org/2000/svg" width="520" height="260" '
            f'viewBox="0 0 520 260" class="figure-panel" role="group" '
            f'aria-label="{title_esc}">'
        )
        chunks.append('<g class="panel-static">')
        chunks.append(
            _render_panel_static(
                categories=categories0,
                series_for_legend=ref_series,
                y0=y0,
                y1=y1,
            )
        )
        chunks.append("</g>")
        for shock in shock_list:
            payload = by_shock[shock]
            panels = list(payload.get("panels") or [])
            if panel_idx >= len(panels):
                continue
            panel = panels[panel_idx]
            series = list(panel.get("series") or [])
            categories = list(payload.get("categories") or [])
            use_categories = categories0 if len(categories) == len(categories0) else categories
            pct_attr = f"{shock:g}"
            sel = abs(shock - default_shock) < 1e-5
            sel_class = " shock-selected" if sel else ""
            chunks.append(f'<g class="shock-layer{sel_class}" data-pct="{pct_attr}">')
            chunks.append(
                _render_shock_polylines(
                    series=series,
                    categories=use_categories,
                    y0=y0,
                    y1=y1,
                )
            )
            chunks.append("</g>")
        chunks.append("</svg></div></div>")

    return "".join(chunks)


def slim_chart_json_for_browser(cache_doc: dict[str, Any]) -> dict[str, Any]:
    by_shock = payloads_by_shock(cache_doc)
    shock_sorted = sorted(by_shock.keys())
    is_pct = cache_doc.get("pct_min") is not None
    entries: list[dict[str, Any]] = []
    for s in shock_sorted:
        row: dict[str, Any] = {"payload": by_shock[s]}
        if is_pct:
            row["pct"] = s
        else:
            row["bps"] = int(s)
        entries.append(row)
    slim: dict[str, Any] = {
        "schema": cache_doc.get("schema"),
        "entries": entries,
    }
    if "pct_min" in cache_doc:
        slim["pct_min"] = cache_doc.get("pct_min")
        slim["pct_max"] = cache_doc.get("pct_max")
        slim["pct_step"] = cache_doc.get("pct_step")
    if "bps_min" in cache_doc:
        slim["bps_min"] = cache_doc.get("bps_min")
        slim["bps_max"] = cache_doc.get("bps_max")
    return slim
