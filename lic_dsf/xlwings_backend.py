from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from typing import Any

from .libreoffice import load_gdp_input_targets, write_shocked_xlsm
from .workbook_payload import read_figure1_payload_from_workbook


def _norm_pct(v: float) -> float:
    return round(float(v), 6)


def _recalculate_with_xlwings(workbook: Path) -> None:
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
        write_shocked_xlsm(src, shocked_xlsm, targets, pct)
        _recalculate_with_xlwings(shocked_xlsm)
        return read_figure1_payload_from_workbook(shocked_xlsm)
    finally:
        if tmpdir is None and not keep_temps:
            shutil.rmtree(tmp, ignore_errors=True)
