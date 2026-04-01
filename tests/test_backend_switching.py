from __future__ import annotations

import json
import unittest
from contextlib import contextmanager
from typing import Iterator

import main


def _backend_state(name: str, *, cache_path: str) -> dict[str, object]:
    return {
        "backend": name,
        "backend_label": "LibreOffice" if name == "libreoffice" else name,
        "charts_html": f'<div class="card">{name}</div>',
        "figure_data": {
            "schema": 3,
            "pct_min": -5.0,
            "pct_max": 5.0,
            "pct_step": 0.5,
            "entries": [{"pct": 0.0, "payload": {"panels": []}}],
        },
        "meta": {
            "cache_path": cache_path,
            "backend": name,
            "schema": 3,
            "generated_at_unix_s": 0,
            "workbook_path": "workbook.xlsm",
        },
        "gdp_label": "GDP forecast growth-rate shock (implied growth + % points, -5% to 5%)",
        "gdp_shock_min": -5.0,
        "gdp_shock_max": 5.0,
        "gdp_shock_step": 0.5,
        "gdp_shock_unit": "pct",
    }


@contextmanager
def _app_state(
    *,
    error: str | None = None,
    selected_backend: str = "excel-grapher",
) -> Iterator[None]:
    sentinel = object()
    saved = {
        "figure_error": getattr(main.app.state, "figure_error", sentinel),
        "figure_backends": getattr(main.app.state, "figure_backends", sentinel),
        "figure_backend_errors": getattr(main.app.state, "figure_backend_errors", sentinel),
        "figure_selected_backend": getattr(main.app.state, "figure_selected_backend", sentinel),
    }
    try:
        main.app.state.figure_error = error
        main.app.state.figure_backends = {
            "excel-grapher": _backend_state(
                "excel-grapher", cache_path=".cache/gdp-shocks-excel-grapher.json"
            ),
            "xlwings": _backend_state("xlwings", cache_path=".cache/gdp-shocks-xlwings.json"),
        }
        main.app.state.figure_backend_errors = {}
        main.app.state.figure_selected_backend = selected_backend
        yield
    finally:
        for key, value in saved.items():
            if value is sentinel:
                try:
                    delattr(main.app.state, key)
                except AttributeError:
                    pass
            else:
                setattr(main.app.state, key, value)


class BackendSwitchingTests(unittest.TestCase):
    def test_index_renders_backend_select_and_selected_option(self) -> None:
        with _app_state():
            html = main.index("xlwings")

        self.assertIn('id="backend-select"', html)
        self.assertIn('option value="excel-grapher"', html)
        self.assertIn('option value="xlwings" selected', html)

    def test_api_figure1_state_returns_requested_backend(self) -> None:
        with _app_state():
            response = main.api_figure1_state("xlwings")

        payload = json.loads(response.body)
        self.assertEqual(payload["backend"], "xlwings")
        self.assertEqual(payload["meta"]["cache_path"], ".cache/gdp-shocks-xlwings.json")

    def test_api_figure1_data_uses_requested_backend(self) -> None:
        with _app_state():
            response = main.api_figure1_data("xlwings")

        payload = json.loads(response.body)
        self.assertEqual(payload["schema"], 3)
        self.assertEqual(payload["entries"][0]["pct"], 0.0)


if __name__ == "__main__":
    unittest.main()
