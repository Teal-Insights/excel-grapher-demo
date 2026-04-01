from __future__ import annotations

import shutil
import tempfile
from pathlib import Path
from typing import Any

from .libreoffice import (
    find_soffice,
    libreoffice_to_xlsx,
    load_gdp_input_targets,
    write_shocked_xlsm,
)
from .workbook_payload import read_figure1_payload_from_workbook


def _norm_pct(v: float) -> float:
    return round(float(v), 6)


def recalculate_figure1_payload_with_libreoffice(
    workbook: Path,
    *,
    pct: float,
    timeout_s: int = 600,
    soffice: str | None = None,
    tmpdir: Path | None = None,
    keep_temps: bool = False,
) -> dict[str, Any]:
    src = workbook.resolve()
    if not src.is_file():
        raise FileNotFoundError(f"Workbook not found: {src}")

    tmp = tmpdir or Path(tempfile.mkdtemp(prefix="libreoffice-gdp-shock-"))
    try:
        targets = load_gdp_input_targets(src)
        shocked_xlsm = tmp / f"{src.stem}_libreoffice_{_norm_pct(pct):g}pct.xlsm"
        write_shocked_xlsm(src, shocked_xlsm, targets, pct)

        bin_soffice = find_soffice(soffice)
        if not bin_soffice:
            raise RuntimeError(
                "LibreOffice not found (try --soffice or PATH: soffice, libreoffice)."
            )

        recalculated = libreoffice_to_xlsx(
            shocked_xlsm,
            tmp,
            soffice=bin_soffice,
            timeout_s=timeout_s,
        )
        return read_figure1_payload_from_workbook(recalculated)
    finally:
        if tmpdir is None and not keep_temps:
            shutil.rmtree(tmp, ignore_errors=True)
