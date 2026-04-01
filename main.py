"""
Stress chart web app: serves precomputed GDP shock data from JSON.

Library code lives in ``lic_dsf`` and ``web``. No graph load or ``FormulaEvaluator`` at request time;
inline SVG is built at startup from the cache; the slider highlights the selected GDP shock (red, on top) over faint gray overlays of all other shocks.

Run:
  GDP_SHOCK_CACHE=.cache/gdp-shocks-excel-grapher.json uv run uvicorn main:app --reload
  uv run python main.py
  uv run python main.py --cache .cache/gdp-shocks-xlwings.json
"""

from __future__ import annotations

import json
import os
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Response
from fastapi.responses import HTMLResponse, JSONResponse
from jinja2 import Environment, FileSystemLoader, select_autoescape

from lic_dsf.cache_compare import STANDARD_CACHE_BACKENDS, available_backend_cache_paths
from lic_dsf.payload import (
    GDP_SHOCK_PCT_MAX,
    GDP_SHOCK_PCT_MIN,
    GDP_SHOCK_PCT_STEP,
)
from web.charts import build_chart_html, load_shock_json, slim_chart_json_for_browser

_ROOT = Path(__file__).resolve().parent
_DEFAULT_CACHE = _ROOT / ".cache" / "gdp-shocks-excel-grapher.json"
_TEMPLATES_DIR = _ROOT / "templates"
_BACKEND_LABELS = {
    "excel-grapher": "excel-grapher",
    "xlwings": "xlwings",
    "libreoffice": "LibreOffice",
}
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


def _backend_label(backend: str) -> str:
    return _BACKEND_LABELS.get(backend, backend)


def _shock_controls_from_slim(slim: dict[str, Any]) -> dict[str, Any]:
    if slim.get("pct_min") is not None:
        shock_min = float(slim["pct_min"])
        shock_max = float(slim["pct_max"])
        shock_step = float(slim["pct_step"])
        shock_unit = "pct"
        gdp_label = (
            "GDP forecast growth-rate shock (implied growth + % points, "
            f"{shock_min:g}% to {shock_max:g}%)"
        )
    else:
        shock_min = float(slim.get("bps_min", -10))
        shock_max = float(slim.get("bps_max", 10))
        shock_step = 1.0
        shock_unit = "bps"
        gdp_label = (
            f"GDP forecast (legacy cache: baseline × (1 + bps×10⁻⁴), "
            f"{int(shock_min)} to {int(shock_max)} bps)"
        )
    return {
        "gdp_label": gdp_label,
        "gdp_shock_min": shock_min,
        "gdp_shock_max": shock_max,
        "gdp_shock_step": shock_step,
        "gdp_shock_unit": shock_unit,
    }


def _load_backend_view_state(backend: str, path: Path) -> dict[str, Any]:
    doc = load_shock_json(path)
    slim = slim_chart_json_for_browser(doc)
    return {
        "backend": backend,
        "backend_label": _backend_label(backend),
        "charts_html": build_chart_html(doc),
        "figure_data": slim,
        "meta": {
            "cache_path": str(path),
            "backend": doc.get("backend"),
            "schema": doc.get("schema"),
            "generated_at_unix_s": doc.get("generated_at_unix_s"),
            "workbook_path": doc.get("workbook_path"),
        },
        **_shock_controls_from_slim(slim),
    }


def _ordered_backend_ids(backend_states: dict[str, dict[str, Any]]) -> list[str]:
    standard = [backend for backend in STANDARD_CACHE_BACKENDS if backend in backend_states]
    custom = sorted(backend for backend in backend_states if backend not in STANDARD_CACHE_BACKENDS)
    return [*standard, *custom]


def _backend_options(backend_states: dict[str, dict[str, Any]]) -> list[dict[str, str]]:
    return [
        {"id": backend, "label": backend_states[backend]["backend_label"]}
        for backend in _ordered_backend_ids(backend_states)
    ]


def _resolve_backend_choice(
    requested_backend: str | None,
    backend_states: dict[str, dict[str, Any]],
    default_backend: str | None,
) -> str | None:
    if requested_backend in backend_states:
        return requested_backend
    if default_backend in backend_states:
        return default_backend
    ordered = _ordered_backend_ids(backend_states)
    return ordered[0] if ordered else None


@asynccontextmanager
async def _lifespan(app: FastAPI):
    configured_path = _cache_path().expanduser()
    discovered_paths = available_backend_cache_paths(_ROOT)
    backend_states: dict[str, dict[str, Any]] = {}
    backend_errors: dict[str, str] = {}
    selected_backend: str | None = None

    if configured_path.is_file():
        try:
            configured_doc = load_shock_json(configured_path)
            configured_backend = str(configured_doc.get("backend") or "").strip()
            if configured_backend:
                discovered_paths[configured_backend] = configured_path.resolve()
                selected_backend = configured_backend
        except (OSError, json.JSONDecodeError, KeyError, TypeError) as e:
            backend_errors["configured"] = f"Could not read figure cache {configured_path}: {e!r}"
    elif configured_path != _DEFAULT_CACHE:
        backend_errors["configured"] = (
            f"Missing cache file: {configured_path}. "
            "Run: uv run python scripts/precache.py --backend excel-grapher"
        )

    for backend, path in discovered_paths.items():
        try:
            backend_states[backend] = _load_backend_view_state(backend, path)
        except (OSError, json.JSONDecodeError, KeyError, TypeError) as e:
            backend_errors[backend] = f"Could not read figure cache {path}: {e!r}"

    selected_backend = _resolve_backend_choice(selected_backend, backend_states, "excel-grapher")
    if backend_states:
        err = None
    elif configured_path.is_file():
        err = backend_errors.get("configured") or next(iter(backend_errors.values()), None)
    else:
        err = (
            f"Missing cache file: {configured_path}. "
            "Run: uv run python scripts/precache.py --backend excel-grapher"
        )

    app.state.figure_error = err
    app.state.figure_backends = backend_states
    app.state.figure_backend_errors = backend_errors
    app.state.figure_selected_backend = selected_backend
    yield


app = FastAPI(title="LIC-DSF Figure 1", lifespan=_lifespan)


@app.get("/favicon.ico")
def favicon() -> Response:
    return Response(status_code=204)


@app.get("/api/figure1-data")
def api_figure1_data(backend: str | None = None) -> JSONResponse:
    err: str | None = getattr(app.state, "figure_error", None)
    if err:
        return JSONResponse({"error": err}, status_code=503)
    backend_states = getattr(app.state, "figure_backends", {})
    selected_backend = _resolve_backend_choice(
        backend,
        backend_states,
        getattr(app.state, "figure_selected_backend", None),
    )
    if selected_backend is None:
        return JSONResponse({"error": "No backend caches available."}, status_code=503)
    return JSONResponse(backend_states[selected_backend]["figure_data"])


@app.get("/api/figure1-state/{backend}")
def api_figure1_state(backend: str) -> JSONResponse:
    err: str | None = getattr(app.state, "figure_error", None)
    if err:
        return JSONResponse({"error": err}, status_code=503)
    backend_states = getattr(app.state, "figure_backends", {})
    if backend not in backend_states:
        raise HTTPException(status_code=404, detail=f"Unknown backend: {backend}")
    return JSONResponse(backend_states[backend])


@app.get("/", response_class=HTMLResponse)
def index(backend: str | None = None) -> str:
    err: str | None = getattr(app.state, "figure_error", None)
    backend_states = getattr(app.state, "figure_backends", {})
    backend_options = _backend_options(backend_states)
    selected_backend = _resolve_backend_choice(
        backend,
        backend_states,
        getattr(app.state, "figure_selected_backend", None),
    )
    if err:
        return _jinja_env.get_template("layout.html").render(
            page_title="Figure 1 — Error",
            ok=False,
            error_text=err,
            charts_html="",
            figure_data_json="null",
            gdp_label=(
                "GDP forecast growth-rate shock (implied growth + % points, "
                f"{GDP_SHOCK_PCT_MIN:g}% to {GDP_SHOCK_PCT_MAX:g}%)"
            ),
            gdp_shock_min=GDP_SHOCK_PCT_MIN,
            gdp_shock_max=GDP_SHOCK_PCT_MAX,
            gdp_shock_step=GDP_SHOCK_PCT_STEP,
            gdp_shock_unit="pct",
            cache_meta_json=_json_for_script_tag({}),
            backend_options=backend_options,
            selected_backend=selected_backend,
        )

    if selected_backend is None:
        raise HTTPException(status_code=503, detail="No backend caches available.")
    state = backend_states[selected_backend]
    return _jinja_env.get_template("layout.html").render(
        page_title="Figure 1 — External stress",
        ok=True,
        error_text="",
        charts_html=state["charts_html"],
        figure_data_json=_json_for_script_tag(state["figure_data"]),
        gdp_label=state["gdp_label"],
        gdp_shock_min=state["gdp_shock_min"],
        gdp_shock_max=state["gdp_shock_max"],
        gdp_shock_step=state["gdp_shock_step"],
        gdp_shock_unit=state["gdp_shock_unit"],
        cache_meta_json=_json_for_script_tag(state["meta"]),
        backend_options=backend_options,
        selected_backend=selected_backend,
    )


def main() -> None:
    import argparse
    import uvicorn

    p = argparse.ArgumentParser(description="Serve LIC-DSF Figure 1 stress chart.")
    p.add_argument(
        "--cache",
        metavar="PATH",
        help="GDP shock JSON cache file (overrides GDP_SHOCK_CACHE / default path)",
    )
    args = p.parse_args()
    if args.cache:
        os.environ["GDP_SHOCK_CACHE"] = str(Path(args.cache).expanduser().resolve())

    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=False)


if __name__ == "__main__":
    main()
