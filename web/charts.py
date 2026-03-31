from __future__ import annotations

import html
import json
import math
from pathlib import Path
from typing import Any, Iterable


def load_shock_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def payloads_by_bps(cache_doc: dict[str, Any]) -> dict[int, dict[str, Any]]:
    out: dict[int, dict[str, Any]] = {}
    default = cache_doc.get("default")
    if isinstance(default, dict) and "bps" in default and "payload" in default:
        out[int(default["bps"])] = default["payload"]  # type: ignore[arg-type]
    for entry in cache_doc.get("shocks") or []:
        if isinstance(entry, dict) and "bps" in entry and "payload" in entry:
            out[int(entry["bps"])] = entry["payload"]  # type: ignore[arg-type]
    return out


def _escape(text: str) -> str:
    return html.escape(text, quote=True)


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
    n = len(categories)
    if n == 0:
        return '<text x="10" y="24" font-size="12" fill="#666">No categories</text>'

    plot_w = float(width) - margin_left - margin_right
    plot_h = float(height) - margin_top - margin_bottom
    y0, y1 = _y_domain(series)

    def x_px(i: int) -> float:
        if n <= 1:
            return margin_left + plot_w / 2.0
        return margin_left + (i / (n - 1)) * plot_w

    def y_px(v: float) -> float:
        t = (v - y0) / (y1 - y0) if y1 != y0 else 0.5
        return margin_top + (1.0 - t) * plot_h

    parts: list[str] = [
        f'<rect x="0" y="0" width="{width}" height="{height}" fill="white"/>',
    ]

    grid_lines = 5
    for g in range(grid_lines + 1):
        gv = y0 + (y1 - y0) * (g / grid_lines)
        gy = y_px(gv)
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
        cx = x_px(i)
        parts.append(
            f'<text x="{cx:.2f}" y="{margin_top + plot_h + 14:.2f}" font-size="8" fill="#666" '
            f'text-anchor="middle">{_escape(str(cat))}</text>'
        )

    for s in series:
        color = str(s.get("borderColor") or "#000000")
        dash = s.get("borderDash") or []
        ys = list(s.get("data") or [])
        d_attr = _dash_attr(dash if isinstance(dash, list) else [])
        for seg in _segments(ys):
            pts = " ".join(f"{x_px(i):.2f},{y_px(v):.2f}" for i, v in seg)
            parts.append(
                f'<polyline fill="none" stroke="{_escape(color)}" stroke-width="1.6" '
                f'points="{pts}"{d_attr}/>'
            )

    leg_y = margin_top + plot_h + 34
    leg_x0 = margin_left
    names_colors: Iterable[tuple[str, str]] = (
        (str(s.get("name", "") or ""), str(s.get("borderColor") or "#000000")) for s in series
    )
    items = [(name, color) for name, color in names_colors if name]
    col_w = plot_w / 2.0
    for idx, (name, color) in enumerate(items):
        col = idx % 2
        row = idx // 2
        lx = leg_x0 + col * col_w
        ly = leg_y + row * 12
        parts.append(
            f'<line x1="{lx:.2f}" y1="{ly:.2f}" x2="{lx + 14:.2f}" y2="{ly:.2f}" '
            f'stroke="{_escape(color)}" stroke-width="2"/>'
        )
        parts.append(
            f'<text x="{lx + 18:.2f}" y="{ly + 3:.2f}" font-size="8" fill="#333">{_escape(name)}</text>'
        )

    return "".join(parts)


def build_chart_html(cache_doc: dict[str, Any]) -> str:
    """
    One <svg> per panel; each contains a <g class="shock-layer" data-bps="…"> per shock.
    Default bps (0) is unhidden; others use the HTML hidden attribute.
    """
    by_bps = payloads_by_bps(cache_doc)
    if not by_bps:
        return '<p class="err">Cache has no chart entries.</p>'

    bps_list = sorted(by_bps.keys())
    ref0 = by_bps.get(0) or by_bps[bps_list[0]]
    categories0 = list(ref0.get("categories") or [])
    panels0 = list(ref0.get("panels") or [])

    chunks: list[str] = []
    for panel_idx, ref_panel in enumerate(panels0):
        title = str(ref_panel.get("title") or f"Panel {panel_idx}")
        title_esc = html.escape(title)
        chunks.append('<div class="card">')
        chunks.append(f"<h2>{title_esc}</h2>")
        chunks.append('<div class="chart-wrap">')
        chunks.append(
            f'<svg xmlns="http://www.w3.org/2000/svg" width="520" height="260" '
            f'viewBox="0 0 520 260" class="figure-panel" role="group" '
            f'aria-label="{title_esc}">'
        )
        for bps in bps_list:
            payload = by_bps[bps]
            panels = list(payload.get("panels") or [])
            if panel_idx >= len(panels):
                continue
            panel = panels[panel_idx]
            series = list(panel.get("series") or [])
            categories = list(payload.get("categories") or [])
            use_categories = categories0 if len(categories) == len(categories0) else categories
            hidden_attr = "" if bps == 0 else " hidden"
            chunks.append(f'<g class="shock-layer" data-bps="{bps}"{hidden_attr}>')
            chunks.append(
                _render_panel_group(
                    title=title,
                    categories=use_categories,
                    series=series,
                )
            )
            chunks.append("</g>")
        chunks.append("</svg></div></div>")

    return "".join(chunks)


def slim_chart_json_for_browser(cache_doc: dict[str, Any]) -> dict[str, Any]:
    by_bps = payloads_by_bps(cache_doc)
    bps_sorted = sorted(by_bps.keys())
    return {
        "schema": cache_doc.get("schema"),
        "bps_min": cache_doc.get("bps_min"),
        "bps_max": cache_doc.get("bps_max"),
        "entries": [{"bps": bps, "payload": by_bps[bps]} for bps in bps_sorted],
    }
