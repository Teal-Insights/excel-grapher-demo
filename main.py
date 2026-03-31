"""
Figure 1 (Output 2-1 external stress charts): serve data from the dependency
graph (pickle under ``.cache/`` if valid, otherwise built on first request like
``graph.py``), with a fast first paint from each node's ``value``, then refresh
from ``FormulaEvaluator`` when evaluation finishes.

Run: ``uv run python main.py`` or ``uv run uvicorn main:app --reload``.
"""

from __future__ import annotations

import asyncio
import logging
import threading
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.responses import HTMLResponse, StreamingResponse
from jinja2 import Environment, FileSystemLoader, select_autoescape

import extract_graph as lic_graph
from excel_grapher import (
    DependencyGraph,
    DynamicRefError,
    FormulaEvaluator,
    create_dependency_graph,
)

from figure1_data import (
    GDP_SHOCK_BPS_MAX,
    GDP_SHOCK_BPS_MIN,
    build_figure1_payload,
    build_figure1_payload_from_graph_node_cache,
    GDP_FORECAST_SHEET,
    GDP_FORECAST_ROWS,
    GDP_FORECAST_START_COL,
    gdp_forecast_baselines,
    gdp_forecast_cell_keys,
    gdp_forecast_value_from_bps,
    numeric_scalar,
)


def _patch_excel_grapher_functions() -> None:
    """
    Extend excel-grapher's function table with Excel boolean constants.

    Some workbooks encode boolean literals as `FALSE()` / `TRUE()`, which may be
    parsed as function calls by excel-grapher. Treat them as zero-arg functions.
    """
    try:
        import excel_grapher.evaluator.evaluator as evmod
    except Exception:
        return

    evmod.FUNCTIONS.setdefault("FALSE", lambda *args: False)
    evmod.FUNCTIONS.setdefault("_XLFN.FALSE", lambda *args: False)
    evmod.FUNCTIONS.setdefault("TRUE", lambda *args: True)
    evmod.FUNCTIONS.setdefault("_XLFN.TRUE", lambda *args: True)

_log = logging.getLogger("uvicorn.error")

_TEMPLATES_DIR = Path(__file__).resolve().parent / "templates"
_jinja_env = Environment(
    loader=FileSystemLoader(_TEMPLATES_DIR),
    autoescape=select_autoescape(enabled_extensions=("html", "xml")),
)


def _json_for_script_tag(payload: object) -> str:
    """
    Serialize JSON for embedding inside <script type="application/json">.
    Avoids accidentally terminating the script tag.
    """
    import json

    s = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return s.replace("</", "<\\/")


def _render_template(name: str, **ctx: object) -> str:
    return _jinja_env.get_template(name).render(**ctx)


def _collect_export_targets() -> list[str]:
    targets: list[str] = []
    for entry in lic_graph.EXPORT_RANGES:
        sheet_name, range_a1 = lic_graph.parse_range_spec(entry["range_spec"])
        targets.extend(lic_graph.cells_in_range(sheet_name, range_a1))
    return targets


def _load_graph_sync() -> DependencyGraph:
    wb = lic_graph.WORKBOOK_PATH
    cache_path = lic_graph._default_graph_cache_path(wb)
    targets = _collect_export_targets()
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
            "Run `uv run python graph.py`, add dynamic-ref constraints, then retry. "
            f"Detail: {exc}"
        ) from exc
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


def _require_graph(app: FastAPI) -> DependencyGraph:
    g: DependencyGraph | None = getattr(app.state, "dependency_graph", None)
    if g is None:
        raise HTTPException(
            status_code=503,
            detail="Dependency graph not loaded (server startup failed)",
        )
    return g


def _require_figure_evaluator(app: FastAPI) -> FormulaEvaluator:
    ev: FormulaEvaluator | None = getattr(app.state, "figure1_evaluator", None)
    if ev is None:
        raise HTTPException(
            status_code=503,
            detail="FormulaEvaluator not ready (server startup failed)",
        )
    return ev


def _init_gdp_shock_state(graph: DependencyGraph) -> dict[str, Any] | None:
    keys = gdp_forecast_cell_keys(graph, workbook_path=lic_graph.WORKBOOK_PATH)
    if not keys:
        return None
    baselines = gdp_forecast_baselines(
        graph, workbook_path=lic_graph.WORKBOOK_PATH, keys=keys
    )
    if len(baselines) != len(keys):
        return None
    return {"keys": keys, "baselines": baselines}


def _gdp_meta_or_disabled(app: FastAPI) -> dict[str, Any]:
    gs: dict[str, Any] | None = getattr(app.state, "gdp_shock", None)
    if not gs:
        return {"enabled": False}
    g = _require_graph(app)
    with app.state.figure_eval_lock:
        # Infer the current slider position from the first key in the bundle.
        key0 = gs["keys"][0]
        base0 = float(gs["baselines"][0])
        node0 = g.get_node(key0)
        cur0 = numeric_scalar(node0.value if node0 else None)
        if cur0 is None:
            cur0 = base0
        raw_bps = (float(cur0) - base0) / 1e-4
        bps = int(round(raw_bps))
        bps = max(GDP_SHOCK_BPS_MIN, min(GDP_SHOCK_BPS_MAX, bps))

    delta = bps * 1e-4
    readout = f"{bps} bps (add {delta:.4f} to baseline GDP forecast cells)"
    cell_range = f"{GDP_FORECAST_SHEET}!{GDP_FORECAST_START_COL}{GDP_FORECAST_ROWS[0]}…"
    return {
        "enabled": True,
        "bps": bps,
        "bps_min": GDP_SHOCK_BPS_MIN,
        "bps_max": GDP_SHOCK_BPS_MAX,
        "cell_value": float(cur0),
        "baseline_value": base0,
        "label": "GDP forecast adjustment (applied to baseline, \u00b110 bps)",
        "cell": cell_range,
        "readout": readout,
    }


def _figure_payload(app: FastAPI) -> dict[str, Any]:
    fig: dict[str, Any] = getattr(app.state, "figure1", _figure1_state_init())
    return {
        "source": fig.get("source"),
        "categories": fig.get("categories", []),
        "panels": fig.get("panels", []),
        "eval_error": fig.get("eval_error"),
    }


def _render_charts_fragment(app: FastAPI) -> str:
    payload = _figure_payload(app)
    return _render_template(
        "partials/charts.html",
        figure=payload,
        figure_payload_json=_json_for_script_tag(payload),
    )


def _figure1_state_init() -> dict[str, Any]:
    return {
        "source": "loading",
        "categories": [],
        "panels": [],
        "eval_error": None,
    }


@asynccontextmanager
async def _lifespan(app: FastAPI):
    _patch_excel_grapher_functions()
    app.state.figure1 = _figure1_state_init()
    app.state.gdp_shock = None
    app.state.dependency_graph = None
    app.state.figure1_evaluator = None
    app.state.figure_eval_lock = threading.Lock()
    app.state.figure1_user_touched = False
    app.state.figure1_bg_eval_finished = False
    app.state.figure1_sse_clients: set[asyncio.Queue[tuple[str, str]]] = set()

    def _sse_publish(event: str, data: str) -> None:
        clients: set[asyncio.Queue[tuple[str, str]]] = getattr(
            app.state, "figure1_sse_clients", set()
        )
        for q in list(clients):
            try:
                q.put_nowait((event, data))
            except asyncio.QueueFull:
                pass

    async def run_evaluated() -> None:
        try:
            graph: DependencyGraph | None = getattr(
                app.state, "dependency_graph", None
            )
            ev: FormulaEvaluator | None = getattr(
                app.state, "figure1_evaluator", None
            )
            if graph is None or ev is None:
                return

            def work() -> dict[str, Any] | None:
                with app.state.figure_eval_lock:
                    if getattr(app.state, "figure1_user_touched", False):
                        return None
                    return build_figure1_payload(
                        graph, workbook_path=lic_graph.WORKBOOK_PATH, evaluator=ev
                    )

            payload = await asyncio.to_thread(work)
            if payload is None:
                return
            if getattr(app.state, "figure1_user_touched", False):
                return
            app.state.figure1 = {
                "source": "evaluated",
                "categories": payload["categories"],
                "panels": payload["panels"],
                "eval_error": None,
            }
            _sse_publish("figure1_status", "Showing FormulaEvaluator results.")
            _sse_publish("figure1_charts", _render_charts_fragment(app))
        except Exception as exc:
            if not getattr(app.state, "figure1_user_touched", False):
                app.state.figure1["eval_error"] = repr(exc)
                _sse_publish("figure1_error", f"Evaluation: {repr(exc)}")
        finally:
            app.state.figure1_bg_eval_finished = True

    try:
        def _prime_figure1() -> tuple[dict[str, Any], DependencyGraph]:
            g = _load_graph_sync()
            prev = build_figure1_payload_from_graph_node_cache(g)
            return prev, g

        preview, graph0 = await asyncio.to_thread(_prime_figure1)
        app.state.dependency_graph = graph0
        app.state.gdp_shock = _init_gdp_shock_state(graph0)
        app.state.figure1 = {
            "source": "graph_node_cache",
            "categories": preview["categories"],
            "panels": preview["panels"],
            "eval_error": None,
        }
        iterate_enabled = False
        iterate_count = 100
        iterate_delta = 0.001
        if lic_graph.WORKBOOK_PATH.is_file():
            settings = lic_graph.get_calc_settings(lic_graph.WORKBOOK_PATH)
            iterate_enabled = settings.iterate_enabled
            iterate_count = settings.iterate_count
            iterate_delta = settings.iterate_delta
        app.state.figure1_evaluator = FormulaEvaluator(
            graph0,
            iterate_enabled=iterate_enabled,
            iterate_count=iterate_count,
            iterate_delta=iterate_delta,
        )
        _sse_publish(
            "figure1_status",
            "Showing values stored on graph nodes (Excel cache at graph build). Recomputing\u2026",
        )
    except FileNotFoundError as exc:
        app.state.figure1 = {
            "source": "error",
            "categories": [],
            "panels": [],
            "eval_error": str(exc),
        }
        _sse_publish("figure1_error", str(exc))
    except RuntimeError as exc:
        app.state.figure1 = {
            "source": "error",
            "categories": [],
            "panels": [],
            "eval_error": str(exc),
        }
        _sse_publish("figure1_error", str(exc))
    except KeyError as exc:
        app.state.figure1 = {
            "source": "error",
            "categories": [],
            "panels": [],
            "eval_error": str(exc),
        }
        _sse_publish("figure1_error", str(exc))

    asyncio.create_task(run_evaluated())
    yield


app = FastAPI(title="LIC-DSF Figure 1", lifespan=_lifespan)


@app.get("/favicon.ico")
def favicon() -> Response:
    return Response(status_code=204)


@app.get("/sse/figure1")
async def sse_figure1() -> StreamingResponse:
    q: asyncio.Queue[tuple[str, str]] = asyncio.Queue(maxsize=8)
    app.state.figure1_sse_clients.add(q)

    fig = _figure_payload(app)
    if fig.get("eval_error"):
        q.put_nowait(("figure1_error", f"Evaluation: {fig['eval_error']}"))
    if fig.get("source") == "evaluated":
        q.put_nowait(("figure1_status", "Showing FormulaEvaluator results."))
        q.put_nowait(("figure1_charts", _render_charts_fragment(app)))
    elif fig.get("source") == "graph_node_cache":
        q.put_nowait(
            (
                "figure1_status",
                "Showing values stored on graph nodes (Excel cache at graph build). Recomputing\u2026",
            )
        )

    async def gen():
        try:
            while True:
                try:
                    event, data = await asyncio.wait_for(q.get(), timeout=15.0)
                    yield f"event: {event}\n"
                    for line in (data or "").splitlines():
                        yield f"data: {line}\n"
                    yield "\n"
                except asyncio.TimeoutError:
                    yield ": keep-alive\n\n"
        finally:
            app.state.figure1_sse_clients.discard(q)

    return StreamingResponse(gen(), media_type="text/event-stream")


@app.post("/figure1/gdp-shock", response_class=HTMLResponse)
async def htmx_gdp_shock(request: Request) -> str:
    form = await request.form()
    try:
        bps = int(form.get("bps", 0))
    except (TypeError, ValueError):
        bps = 0
    bps = max(GDP_SHOCK_BPS_MIN, min(GDP_SHOCK_BPS_MAX, bps))

    gs: dict[str, Any] | None = getattr(app.state, "gdp_shock", None)
    if not gs:
        return _render_template(
            "partials/error.html",
            error_text="B1_GDP_ext!AA63 is not in the graph. Rebuild: uv run python graph.py --no-cache",
        )

    g = _require_graph(app)
    ev = _require_figure_evaluator(app)
    with app.state.figure_eval_lock:
        keys: list[str] = gs["keys"]
        baselines: list[float] = gs["baselines"]
        for k, base in zip(keys, baselines, strict=True):
            ev.set_value(k, gdp_forecast_value_from_bps(float(base), bps))
        payload = build_figure1_payload(
            g, workbook_path=lic_graph.WORKBOOK_PATH, evaluator=ev
        )
    app.state.figure1_user_touched = True
    app.state.figure1 = {
        "source": "evaluated",
        "categories": payload["categories"],
        "panels": payload["panels"],
        "eval_error": None,
    }
    status = _render_template(
        "partials/status.html",
        status_text="FormulaEvaluator results (baseline GDP forecast adjusted by slider bps).",
    )
    return _render_charts_fragment(app) + status


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    fig = _figure_payload(app)
    gdp = _gdp_meta_or_disabled(app)
    if fig.get("source") == "evaluated":
        status_text = "Showing FormulaEvaluator results."
    elif fig.get("source") == "graph_node_cache":
        status_text = (
            "Showing values stored on graph nodes (Excel cache at graph build). Recomputing\u2026"
        )
    elif fig.get("source") == "error":
        status_text = "Figure data not ready."
    else:
        status_text = ""
    return _render_template(
        "layout.html",
        page_title="Figure 1 — External stress (graph-backed)",
        figure=fig,
        figure_payload_json=_json_for_script_tag(fig),
        gdp=gdp,
        status_text=status_text,
    )


def main() -> None:
    import uvicorn

    uvicorn.run(
        "main:app",
        host="127.0.0.1",
        port=8000,
        reload=False,
    )


if __name__ == "__main__":
    main()
