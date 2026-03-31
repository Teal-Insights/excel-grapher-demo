"""
Stress chart web app: serves precomputed GDP shock data from JSON.

Library code lives in ``lic_dsf`` and ``web``. No graph load or ``FormulaEvaluator`` at request time;
inline SVG is built at startup from the cache; the slider toggles ``hidden`` on pre-rendered ``<g>`` layers.

Run:
  GDP_SHOCK_CACHE=.cache/gdp-shocks.json uv run uvicorn main:app --reload
  uv run python main.py
"""

from __future__ import annotations

import json
import os
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

from fastapi import FastAPI, Response
from fastapi.responses import HTMLResponse, JSONResponse
from jinja2 import Environment, FileSystemLoader, select_autoescape

from lic_dsf.payload import GDP_SHOCK_BPS_MAX, GDP_SHOCK_BPS_MIN
from web.charts import build_chart_html, load_shock_json, slim_chart_json_for_browser

_ROOT = Path(__file__).resolve().parent
_DEFAULT_CACHE = _ROOT / ".cache" / "gdp-shocks.json"
_TEMPLATES_DIR = _ROOT / "templates"
_jinja_env = Environment(
    loader=FileSystemLoader(_TEMPLATES_DIR),
    autoescape=select_autoescape(enabled_extensions=("html", "xml")),
)


def _json_for_script_tag(payload: object) -> str:
    s = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return s.replace("</", "<\\/")


def _cache_path() -> Path:
    raw = os.environ.get("GDP_SHOCK_CACHE", "").strip()
    if not raw:
        raw = os.environ.get("FIGURE1_SHOCK_CACHE", "").strip()
    return Path(raw) if raw else _DEFAULT_CACHE


@asynccontextmanager
async def _lifespan(app: FastAPI):
    path = _cache_path()
    err: str | None = None
    charts_html = ""
    chart_data: dict[str, Any] = {}
    meta: dict[str, Any] = {}
    if not path.is_file():
        err = f"Missing cache file: {path}. Run: uv run python scripts/precache.py"
    else:
        try:
            doc = load_shock_json(path)
            charts_html = build_chart_html(doc)
            chart_data = slim_chart_json_for_browser(doc)
            meta = {
                "cache_path": str(path),
                "schema": doc.get("schema"),
                "generated_at_unix_s": doc.get("generated_at_unix_s"),
                "workbook_path": doc.get("workbook_path"),
            }
        except (OSError, json.JSONDecodeError, KeyError, TypeError) as e:
            err = f"Could not read figure cache {path}: {e!r}"

    app.state.figure_error = err
    app.state.figure_charts_html = charts_html
    app.state.figure_data_slim = chart_data
    app.state.figure_meta = meta
    yield


app = FastAPI(title="LIC-DSF Figure 1", lifespan=_lifespan)


@app.get("/favicon.ico")
def favicon() -> Response:
    return Response(status_code=204)


@app.get("/api/figure1-data")
def api_figure1_data() -> JSONResponse:
    err: str | None = getattr(app.state, "figure_error", None)
    if err:
        return JSONResponse({"error": err}, status_code=503)
    return JSONResponse(getattr(app.state, "figure_data_slim", {}))


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    err: str | None = getattr(app.state, "figure_error", None)
    if err:
        return _jinja_env.get_template("layout.html").render(
            page_title="Figure 1 — Error",
            ok=False,
            error_text=err,
            charts_html="",
            figure_data_json="null",
            gdp_label="GDP forecast (baseline * (1 + bps*1e-4), +/-10 bps)",
            gdp_bps_min=GDP_SHOCK_BPS_MIN,
            gdp_bps_max=GDP_SHOCK_BPS_MAX,
            cache_meta_json=_json_for_script_tag(getattr(app.state, "figure_meta", {})),
        )

    slim = getattr(app.state, "figure_data_slim", {})
    return _jinja_env.get_template("layout.html").render(
        page_title="Figure 1 — External stress",
        ok=True,
        error_text="",
        charts_html=getattr(app.state, "figure_charts_html", ""),
        figure_data_json=_json_for_script_tag(slim),
        gdp_label="GDP forecast (baseline * (1 + bps*1e-4), +/-10 bps)",
        gdp_bps_min=int(slim.get("bps_min") or GDP_SHOCK_BPS_MIN),
        gdp_bps_max=int(slim.get("bps_max") or GDP_SHOCK_BPS_MAX),
        cache_meta_json=_json_for_script_tag(getattr(app.state, "figure_meta", {})),
    )


def main() -> None:
    import uvicorn

    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=False)


if __name__ == "__main__":
    main()
