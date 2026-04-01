from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any, Literal

from .libreoffice import (
    _compare_maps_python_minus_lo,
    _compare_shock_increment_python_minus_lo,
    _norm_pct,
    diff_chart_maps,
    find_soffice,
    figure1_payload_to_chart_map,
    load_gdp_input_targets,
    libreoffice_to_xlsx,
    read_figure1_chart_values,
    write_shocked_xlsm,
)
from .payload import (
    CHART_SHEET,
    FIGURE1_PANELS,
    col_letters,
    read_category_labels_workbook,
)

BackendName = Literal["libreoffice", "xlwings"]


def resolve_backend(name: str) -> BackendName:
    if name == "auto":
        if sys.platform == "win32":
            return "xlwings"
        if sys.platform.startswith("linux"):
            return "libreoffice"
        raise RuntimeError(
            "backend=auto is only supported on Windows and Linux in this script. "
            "Choose --sanity-check-backend libreoffice or --sanity-check-backend xlwings explicitly."
        )
    if name == "libreoffice":
        return "libreoffice"
    if name == "xlwings":
        if sys.platform != "win32":
            raise RuntimeError(
                "backend=xlwings is only supported on Windows in this project."
            )
        return "xlwings"
    raise ValueError(f"Unknown backend: {name}")


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


def _recalculate_with_xlwings(workbook: Path) -> None:
    try:
        import xlwings as xw
    except ImportError as exc:
        raise RuntimeError(
            "xlwings is required for sanity-check-backend=xlwings but is not installed. "
            "Install it in this environment and retry."
        ) from exc

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        book = app.books.open(str(workbook), update_links=False, read_only=False)
        try:
            app.calculation = "automatic"
            api = app.api
            if hasattr(api, "CalculateFullRebuild"):
                api.CalculateFullRebuild()
            elif hasattr(api, "CalculateFull"):
                api.CalculateFull()
            else:
                app.calculate()
            book.save()
        finally:
            book.close()
    finally:
        app.quit()


def recalculate_figure1_payload(
    workbook: Path,
    *,
    pct: float,
    backend: BackendName,
    timeout_s: int = 600,
    soffice: str | None = None,
    tmpdir: Path | None = None,
    keep_temps: bool = False,
) -> dict[str, Any]:
    src = workbook.resolve()
    if not src.is_file():
        raise FileNotFoundError(f"Workbook not found: {src}")

    tmp = tmpdir or Path(tempfile.mkdtemp(prefix=f"{backend}-gdp-shock-"))
    try:
        targets = load_gdp_input_targets(src)
        p = _norm_pct(pct)
        shocked_xlsm = tmp / f"{src.stem}_{backend}_{p}pct.xlsm"
        write_shocked_xlsm(src, shocked_xlsm, targets, pct)

        if backend == "libreoffice":
            bin_soffice = find_soffice(soffice)
            if not bin_soffice:
                raise RuntimeError(
                    "LibreOffice not found (try --lo-soffice or PATH: soffice, libreoffice)."
                )
            recalculated = libreoffice_to_xlsx(
                shocked_xlsm,
                tmp,
                soffice=bin_soffice,
                timeout_s=timeout_s,
            )
        else:
            _recalculate_with_xlwings(shocked_xlsm)
            recalculated = shocked_xlsm

        return read_figure1_payload_from_workbook(recalculated)
    finally:
        if tmpdir is None and not keep_temps:
            shutil.rmtree(tmp, ignore_errors=True)


def run_workbook_gdp_shock_check(
    workbook: Path,
    *,
    backend: BackendName,
    baseline_pct: float = 0.0,
    shock_pct: float = 1.0,
    timeout_s: int = 600,
    soffice: str | None = None,
    keep_temps: bool = False,
    top_n: int = 15,
    python_baseline_payload: dict[str, Any] | None = None,
    python_shock_payload: dict[str, Any] | None = None,
) -> dict[str, Any]:
    src = workbook.resolve()
    if not src.is_file():
        return {"ok": False, "error": f"Workbook not found: {src}"}

    tmp = Path(tempfile.mkdtemp(prefix=f"{backend}-gdp-check-"))
    try:
        targets = load_gdp_input_targets(src)
        b0, b1 = _norm_pct(baseline_pct), _norm_pct(shock_pct)
        base_payload = recalculate_figure1_payload(
            src,
            pct=baseline_pct,
            backend=backend,
            timeout_s=timeout_s,
            soffice=soffice,
            tmpdir=tmp,
            keep_temps=keep_temps,
        )
        shock_payload = recalculate_figure1_payload(
            src,
            pct=shock_pct,
            backend=backend,
            timeout_s=timeout_s,
            soffice=soffice,
            tmpdir=tmp,
            keep_temps=keep_temps,
        )

        vb = figure1_payload_to_chart_map(base_payload)
        vs = figure1_payload_to_chart_map(shock_payload)
        rows = diff_chart_maps(vb, vs)

        finite_deltas = [r[3] for r in rows if r[3] is not None]
        max_abs = max(abs(d) for d in finite_deltas) if finite_deltas else None
        mean_abs = (
            sum(abs(d) for d in finite_deltas) / len(finite_deltas)
            if finite_deltas
            else None
        )

        top: list[dict[str, Any]] = []
        for addr, b0, b1, d in rows:
            if d is None:
                continue
            top.append(
                {
                    "cell": addr,
                    "baseline_value": b0,
                    "shock_value": b1,
                    "delta": d,
                }
            )
            if len(top) >= top_n:
                break

        out: dict[str, Any] = {
            "ok": True,
            "backend": backend,
            "baseline_pct": b0,
            "shock_pct": b1,
            "gdp_input_cells": len(targets),
            "output_cells_compared": len(rows),
            "backend_internal": {
                "description": f"{backend} recalc: shock minus baseline (same engine only)",
                "max_abs_delta": max_abs,
                "mean_abs_delta": mean_abs,
                "top_deltas": top,
            },
            "max_abs_delta": max_abs,
            "mean_abs_delta": mean_abs,
            "top_deltas": top,
        }
        if backend == "libreoffice":
            out["soffice"] = find_soffice(soffice)

        if python_baseline_payload is not None and python_shock_payload is not None:
            try:
                py_b = figure1_payload_to_chart_map(python_baseline_payload)
                py_s = figure1_payload_to_chart_map(python_shock_payload)
                out["python_vs_backend"] = {
                    "description": (
                        "FormulaEvaluator payload vs workbook recalc payload "
                        f"(same cell keys); errors are python - {backend}"
                    ),
                    f"at_{b0:g}_pct": _compare_maps_python_minus_lo(
                        py_b, vb, top_n=top_n
                    ),
                    f"at_{b1:g}_pct": _compare_maps_python_minus_lo(
                        py_s, vs, top_n=top_n
                    ),
                    "shock_increment": _compare_shock_increment_python_minus_lo(
                        py_b, py_s, vb, vs, top_n=top_n
                    ),
                }
            except ValueError as e:
                out["python_vs_backend"] = {
                    "error": f"Could not map Python payload: {e!s}",
                }
        if keep_temps:
            out["temp_dir"] = str(tmp)
        return out
    except RuntimeError as e:
        err_out: dict[str, Any] = {"ok": False, "error": repr(e)}
        if keep_temps:
            err_out["temp_dir"] = str(tmp)
        return err_out
    finally:
        if not keep_temps:
            shutil.rmtree(tmp, ignore_errors=True)


def print_check_report(result: dict[str, Any]) -> None:
    if not result.get("ok"):
        print(f"Workbook sanity check failed: {result.get('error', 'unknown')}")
        if result.get("temp_dir"):
            print(f"Temp dir: {result['temp_dir']}")
        return

    backend = result.get("backend", "workbook")
    detail = result.get("soffice") if backend == "libreoffice" else "Excel via xlwings"
    print(f"Workbook sanity check ({backend}): {detail}")
    print(f"  GDP forecast input cells: {result['gdp_input_cells']}")
    print(
        f"  Internal: {result['baseline_pct']}% vs {result['shock_pct']}% "
        f"(recalc delta under {backend} only)"
    )
    print(f"  Output cells compared: {result['output_cells_compared']}")
    ma = result.get("max_abs_delta")
    mn = result.get("mean_abs_delta")
    if ma is not None:
        print(f"  Max |shock - baseline|: {ma:.6g}")
    if mn is not None:
        print(f"  Mean |shock - baseline|: {mn:.6g}")
    print("  Largest |shock - baseline| (sample):")
    for row in result.get("top_deltas") or []:
        print(
            f"    {row['cell']:28}  base={row['baseline_value']!s:>14}  "
            f"shock={row['shock_value']!s:>14}  Δ={row['delta']: .6g}"
        )

    pvb = result.get("python_vs_backend")
    if pvb:
        if pvb.get("error"):
            print(f"  Python vs {backend}: {pvb['error']}")
        else:
            print(f"  Python (FormulaEvaluator) vs {backend} — same Chart Data cells:")
            for key in sorted(k for k in pvb if k not in ("description", "error")):
                block = pvb[key]
                if not isinstance(block, dict):
                    continue
                if block.get("error"):
                    print(f"    {key}: {block['error']}")
                    continue
                mx = block.get("max_abs_error")
                mn_b = block.get("mean_abs_error")
                nc = block.get("cells_compared")
                if mx is not None and mn_b is not None:
                    print(
                        f"    {key}: cells={nc}  max|py-{backend}|={mx:.6g}  "
                        f"mean|py-{backend}|={mn_b:.6g}"
                    )
                else:
                    print(f"    {key}: cells={nc}")
                for row in (block.get("top_errors") or [])[:5]:
                    if "python_minus_libreoffice" in row:
                        print(
                            f"      {row['cell']}  py={row['python']:.6g}  "
                            f"{backend}={row['libreoffice']:.6g}  "
                            f"py-{backend}={row['python_minus_libreoffice']:.6g}"
                        )
                    elif "python_delta_minus_lo_delta" in row:
                        print(
                            f"      {row['cell']}  Δpy={row['python_delta']:.6g}  "
                            f"Δ{backend}={row['libreoffice_delta']:.6g}  "
                            f"Δpy-Δ{backend}={row['python_delta_minus_lo_delta']:.6g}"
                        )

    if result.get("temp_dir"):
        print(f"  Kept temp dir: {result['temp_dir']}")
