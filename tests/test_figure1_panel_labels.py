from __future__ import annotations

import unittest

from lic_dsf.payload import CHART_SHEET, FIGURE1_PANELS, col_letters
from lic_dsf.workbook_payload import figure1_payload_from_chart_map
from web.charts import build_chart_html


def _chart_map_with_constant_values() -> dict[str, float]:
    chart_map: dict[str, float] = {}
    value = 1.0
    for panel in FIGURE1_PANELS:
        for spec in panel.series:
            for col in col_letters():
                chart_map[f"{CHART_SHEET}!{col}{spec.value_row}"] = value
                value += 1.0
    return chart_map


def _panel_annotations() -> list[dict[str, object]]:
    return [
        {
            "mostExtremeShockLabel": "PV of Debt to GDP Ratio MX Shock",
            "baselineBreaches": 1,
            "shockBreaches": 2,
        },
        {
            "mostExtremeShockLabel": "PV of Debt to Exports Ratio MX Shock",
            "baselineBreaches": 3,
            "shockBreaches": 4,
        },
        {
            "mostExtremeShockLabel": "Debt Service to Exports Ratio MX Shock",
            "baselineBreaches": 5,
            "shockBreaches": 6,
        },
        {
            "mostExtremeShockLabel": "Debt Service to Revenue MX Shock- Market",
            "baselineBreaches": 7,
            "shockBreaches": 8,
        },
    ]


class Figure1PanelLabelTests(unittest.TestCase):
    def test_workbook_payload_includes_panel_titles_labels_and_breach_counts(self) -> None:
        payload = figure1_payload_from_chart_map(
            categories=["2025", "2026", "2027", "2028", "2029", "2030", "2031", "2032", "2033", "2034", "2035"],
            chart_map=_chart_map_with_constant_values(),
            panel_annotations=_panel_annotations(),
        )

        panels = payload["panels"]
        self.assertEqual(
            [panel["title"] for panel in panels],
            [
                "PV of debt-to-GDP ratio",
                "PV of debt-to-exports ratio",
                "Debt service-to-exports ratio",
                "Debt service-to-revenue ratio",
            ],
        )
        self.assertEqual(
            [panel["mostExtremeShockLabel"] for panel in panels],
            [
                "PV of Debt to GDP Ratio MX Shock",
                "PV of Debt to Exports Ratio MX Shock",
                "Debt Service to Exports Ratio MX Shock",
                "Debt Service to Revenue MX Shock- Market",
            ],
        )
        self.assertEqual([panel["baselineBreaches"] for panel in panels], [1, 3, 5, 7])
        self.assertEqual([panel["shockBreaches"] for panel in panels], [2, 4, 6, 8])

    def test_chart_html_renders_most_extreme_shock_label_and_breach_counts(self) -> None:
        payload = figure1_payload_from_chart_map(
            categories=["2025", "2026", "2027", "2028", "2029", "2030", "2031", "2032", "2033", "2034", "2035"],
            chart_map=_chart_map_with_constant_values(),
            panel_annotations=_panel_annotations(),
        )
        cache_doc = {
            "schema": 3,
            "pct_min": -5.0,
            "pct_max": 5.0,
            "pct_step": 0.5,
            "default": {"pct": 0.0, "payload": payload},
            "shocks": [{"pct": 1.0, "payload": payload}],
        }

        html = build_chart_html(cache_doc)

        self.assertIn('class="panel-shock-label"', html)
        self.assertIn('class="panel-breach-counts"', html)
        self.assertIn("PV of Debt to GDP Ratio MX Shock", html)
        self.assertIn("Debt Service to Revenue MX Shock- Market", html)
        self.assertIn("Baseline breaches: <span class=\"count\">1</span>", html)
        self.assertIn("Shock breaches: <span class=\"count\">8</span>", html)


if __name__ == "__main__":
    unittest.main()
