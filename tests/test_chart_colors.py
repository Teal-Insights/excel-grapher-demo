from __future__ import annotations

import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

from web.charts import build_chart_html


SVG_NS = {"svg": "http://www.w3.org/2000/svg"}
TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "templates" / "layout.html"


def _sample_cache_doc() -> dict[str, object]:
    categories = ["2024", "2025", "2026"]
    series = [
        {"name": "Baseline", "data": [10.0, 11.0, 12.0], "borderColor": "#4b82ad", "borderDash": []},
        {
            "name": "Historical scenario",
            "data": [9.0, 10.0, 11.0],
            "borderColor": "#ff0000",
            "borderDash": [10, 5],
        },
        {
            "name": "MX shock Standard&Tailored",
            "data": [8.0, 8.5, 9.0],
            "borderColor": "#000000",
            "borderDash": [],
        },
        {"name": "Threshold", "data": [12.0, 12.0, 12.0], "borderColor": "#339966", "borderDash": [6, 4]},
    ]
    payload = {"categories": categories, "panels": [{"title": "Panel 1", "series": series}]}
    return {
        "schema": 3,
        "pct_min": -0.5,
        "pct_max": 0.5,
        "pct_step": 0.5,
        "default": {"pct": 0.0, "payload": payload},
        "shocks": [{"pct": -0.5, "payload": payload}],
    }


def _legacy_cache_doc_with_hidden_series() -> dict[str, object]:
    categories = ["2024", "2025", "2026"]
    series = [
        {"name": "Baseline", "data": [10.0, 11.0, 12.0], "borderColor": "#4b82ad", "borderDash": []},
        {
            "name": "MX value, 1 yr only shock Standard&Tailored - for chart",
            "data": [7.0, 7.5, 8.0],
            "borderColor": "#e46c0a",
            "borderDash": [],
        },
        {"name": "Risk band", "data": [6.0, 6.0, 6.0], "borderColor": "#00ff00", "borderDash": []},
        {"name": "Threshold", "data": [12.0, 12.0, 12.0], "borderColor": "#339966", "borderDash": [6, 4]},
    ]
    payload = {"categories": categories, "panels": [{"title": "Panel 1", "series": series}]}
    return {
        "schema": 3,
        "pct_min": -0.5,
        "pct_max": 0.5,
        "pct_step": 0.5,
        "default": {"pct": 0.0, "payload": payload},
        "shocks": [{"pct": -0.5, "payload": payload}],
    }


class ChartColorTests(unittest.TestCase):
    def test_rendered_lines_infer_focal_metadata_for_legacy_cache_entries(self) -> None:
        html = build_chart_html(_sample_cache_doc())
        root = ET.fromstring(f"<root>{html}</root>")

        selected_layer = root.find(".//svg:g[@class='shock-layer shock-selected']", SVG_NS)
        self.assertIsNotNone(selected_layer)
        assert selected_layer is not None

        focal_by_name = {
            line.attrib["data-series-name"]: line.attrib["data-focal"]
            for line in selected_layer.findall(".//svg:polyline", SVG_NS)
        }

        self.assertEqual(
            focal_by_name,
            {
                "Baseline": "true",
                "Historical scenario": "true",
                "MX shock Standard&Tailored": "true",
                "Threshold": "true",
            },
        )

    def test_selected_focal_lines_match_legend_colors(self) -> None:
        html = build_chart_html(_sample_cache_doc())
        root = ET.fromstring(f"<root>{html}</root>")

        svg = root.find(".//svg:svg", SVG_NS)
        self.assertIsNotNone(svg)
        assert svg is not None

        swatches = {
            line.attrib["data-series-name"]: line.attrib["stroke"]
            for line in svg.findall(".//svg:line[@class='legend-swatch']", SVG_NS)
        }

        selected_layer = svg.find(".//svg:g[@class='shock-layer shock-selected']", SVG_NS)
        self.assertIsNotNone(selected_layer)
        assert selected_layer is not None

        focal_lines = {
            line.attrib["data-series-name"]: line.attrib
            for line in selected_layer.findall(".//svg:polyline", SVG_NS)
            if line.attrib["data-focal"] == "true"
        }

        self.assertEqual(focal_lines["Baseline"]["stroke"], swatches["Baseline"])
        self.assertEqual(focal_lines["Historical scenario"]["stroke"], swatches["Historical scenario"])
        self.assertEqual(
            focal_lines["MX shock Standard&Tailored"]["stroke"],
            swatches["MX shock Standard&Tailored"],
        )
        self.assertEqual(focal_lines["Threshold"]["stroke"], swatches["Threshold"])

    def test_layout_css_tints_background_layers_and_grays_selected_non_focal_lines(self) -> None:
        template = TEMPLATE_PATH.read_text(encoding="utf-8")

        self.assertIn('g.shock-layer:not(.shock-selected) polyline.shock-line', template)
        self.assertIn("stroke-opacity: 0.36;", template)
        self.assertIn('g.shock-layer.shock-selected polyline.shock-line[data-focal="false"]', template)
        self.assertIn("stroke: rgba(45, 45, 45, 0.28);", template)

    def test_renderer_hides_hidden_series_from_legacy_cache_content(self) -> None:
        html = build_chart_html(_legacy_cache_doc_with_hidden_series())

        self.assertNotIn("Risk band", html)
        self.assertNotIn("MX value, 1 yr only shock Standard&Tailored - for chart", html)


if __name__ == "__main__":
    unittest.main()
