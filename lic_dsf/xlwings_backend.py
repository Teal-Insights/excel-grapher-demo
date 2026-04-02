from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from typing import Any

from .libreoffice import load_gdp_input_targets
from .payload import gdp_forecast_values_from_percent
from .workbook_payload import read_figure1_payload_from_workbook


def _norm_pct(v: float) -> float:
    return round(float(v), 6)


def _write_shocked_inputs_with_xlwings(
    book: Any,
    *,
    targets: list[tuple[str, str, float]],
    pct: float,
) -> None:
    shocked_series = gdp_forecast_values_from_percent(
        [base for _sheet_name, _a1, base in targets],
        pct,
    )
    for (sheet_name, a1, _base), shocked_value in zip(
        targets,
        shocked_series,
        strict=True,
    ):
        book.sheets[sheet_name].range(a1).value = shocked_value


def _recalculate_with_xlwings(
    workbook: Path,
    *,
    targets: list[tuple[str, str, float]],
    pct: float,
) -> None:
    try:
        import xlwings as xw
    except ImportError as exc:
        raise RuntimeError(
            "xlwings is required for backend=xlwings but is not installed. "
            "Install it in this environment and retry."
        ) from exc

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        book = app.books.open(str(workbook), update_links=False, read_only=False)
        try:
            _write_shocked_inputs_with_xlwings(book, targets=targets, pct=pct)
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


def recalculate_figure1_payload_with_xlwings(
    workbook: Path,
    *,
    pct: float,
    tmpdir: Path | None = None,
    keep_temps: bool = False,
) -> dict[str, Any]:
    src = workbook.resolve()
    if not src.is_file():
        raise FileNotFoundError(f"Workbook not found: {src}")

    tmp = tmpdir or Path(tempfile.mkdtemp(prefix="xlwings-gdp-shock-"))
    try:
        targets = load_gdp_input_targets(src)
        shocked_xlsm = tmp / f"{src.stem}_xlwings_{_norm_pct(pct):g}pct.xlsm"
        shutil.copy2(src, shocked_xlsm)
        _recalculate_with_xlwings(shocked_xlsm, targets=targets, pct=pct)
        return read_figure1_payload_from_workbook(shocked_xlsm)
    finally:
        if tmpdir is None and not keep_temps:
            shutil.rmtree(tmp, ignore_errors=True)
