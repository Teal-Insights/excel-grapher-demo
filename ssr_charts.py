from __future__ import annotations

import io
from dataclasses import dataclass
from typing import Any, Iterable

import matplotlib

# Use a non-interactive backend suitable for servers.
matplotlib.use("Agg")
import matplotlib.pyplot as plt


@dataclass(frozen=True, slots=True)
class SvgChartTheme:
    width_px: int = 520
    height_px: int = 240
    dpi: int = 120
    font_size: int = 9
    grid_color: str = "#e8e9eb"
    axis_color: str = "#666666"
    background: str = "white"


def _to_float_list(xs: Iterable[Any]) -> list[float | None]:
    out: list[float | None] = []
    for v in xs:
        if v is None:
            out.append(None)
        elif isinstance(v, (int, float)):
            out.append(float(v))
        else:
            try:
                out.append(float(v))
            except Exception:
                out.append(None)
    return out


def _dash_from_excel(border_dash: list[int] | None) -> tuple[float, ...] | None:
    if not border_dash:
        return None
    # Chart.js uses pixel lengths. Matplotlib expects "on, off" in points; treat
    # them proportionally.
    return tuple(float(x) for x in border_dash)


def render_panel_svg(
    *,
    title: str,
    categories: list[str],
    series: list[dict[str, Any]],
    theme: SvgChartTheme | None = None,
) -> str:
    """
    Render one panel (multiple series) as inline SVG.

    `series` format matches the payload produced by figure1_data.build_figure1_payload*:
      { name, data, borderColor, borderDash }
    """
    th = theme or SvgChartTheme()

    fig_w = th.width_px / th.dpi
    fig_h = th.height_px / th.dpi
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=th.dpi)
    fig.patch.set_facecolor(th.background)
    ax.set_facecolor(th.background)

    x = list(range(len(categories)))
    for s in series:
        name = str(s.get("name", "") or "")
        color = str(s.get("borderColor", "#000000") or "#000000")
        ys = _to_float_list(s.get("data") or [])
        dash = _dash_from_excel(s.get("borderDash"))

        # Matplotlib breaks lines on NaNs; convert None to NaN.
        y_plot = [float("nan") if v is None else float(v) for v in ys]
        (line,) = ax.plot(x, y_plot, label=name, color=color, linewidth=1.6)
        if dash:
            line.set_dashes(dash)

    ax.set_title(title, fontsize=th.font_size + 1, loc="left", pad=8)
    ax.grid(True, axis="y", color=th.grid_color, linewidth=0.8)
    ax.tick_params(axis="both", labelsize=th.font_size, colors=th.axis_color)

    # Keep x-axis readable (11 points typically). Prefer horizontal labels.
    ax.set_xticks(x)
    ax.set_xticklabels(categories, rotation=0, ha="center")

    # Legend at bottom, compact.
    ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.22),
        ncol=2,
        frameon=False,
        fontsize=th.font_size,
        handlelength=2.0,
        columnspacing=1.0,
    )

    fig.tight_layout()

    buf = io.StringIO()
    fig.savefig(buf, format="svg")
    plt.close(fig)

    svg = buf.getvalue()
    # Matplotlib includes an XML header; inline SVG works without it and it can
    # complicate HTML concatenation.
    if svg.startswith("<?xml"):
        svg = svg.split("?>", 1)[1]
    return svg.strip()


def render_figure_svg_panels(
    figure: dict[str, Any],
    *,
    theme: SvgChartTheme | None = None,
) -> list[dict[str, str]]:
    """
    Render all panels in the Figure 1 payload to SVG strings.

    Returns list entries: {title, svg}
    """
    categories = list(figure.get("categories") or [])
    panels = list(figure.get("panels") or [])
    out: list[dict[str, str]] = []
    for p in panels:
        title = str(p.get("title", "") or "")
        svg = render_panel_svg(
            title=title,
            categories=categories,
            series=list(p.get("series") or []),
            theme=theme,
        )
        out.append({"title": title, "svg": svg})
    return out

