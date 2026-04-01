from __future__ import annotations

import importlib.util
import re
import unittest
from pathlib import Path

from web.charts import build_chart_html


TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "templates" / "layout.html"
_BROWSER_CANDIDATES = (
    Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
    Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
)


def _sample_cache_doc() -> dict[str, object]:
    categories = ["2024", "2025", "2026", "2027"]
    default_series = [
        {"name": "Baseline", "data": [11.0, 12.0, 13.0, 14.0], "borderColor": "#4b82ad", "borderDash": []},
        {
            "name": "Historical scenario",
            "data": [6.0, 7.0, 8.0, 9.0],
            "borderColor": "#ff0000",
            "borderDash": [10, 5],
        },
        {
            "name": "MX shock Standard&Tailored",
            "data": [16.0, 15.0, 14.0, 13.0],
            "borderColor": "#000000",
            "borderDash": [],
        },
        {
            "name": "MX value, 1 yr only shock Standard&Tailored - for chart",
            "data": [4.0, 4.5, 5.0, 5.5],
            "borderColor": "#e46c0a",
            "borderDash": [],
        },
        {"name": "Threshold", "data": [17.0, 17.0, 17.0, 17.0], "borderColor": "#339966", "borderDash": [6, 4]},
    ]
    shock_series = [
        {"name": "Baseline", "data": [9.5, 10.0, 10.8, 11.7], "borderColor": "#4b82ad", "borderDash": []},
        {
            "name": "Historical scenario",
            "data": [4.5, 5.5, 6.2, 6.9],
            "borderColor": "#ff0000",
            "borderDash": [10, 5],
        },
        {
            "name": "MX shock Standard&Tailored",
            "data": [14.0, 13.0, 12.2, 11.4],
            "borderColor": "#000000",
            "borderDash": [],
        },
        {
            "name": "MX value, 1 yr only shock Standard&Tailored - for chart",
            "data": [5.2, 5.8, 6.4, 7.0],
            "borderColor": "#e46c0a",
            "borderDash": [],
        },
        {"name": "Threshold", "data": [18.0, 18.0, 18.0, 18.0], "borderColor": "#339966", "borderDash": [6, 4]},
    ]
    default_payload = {
        "categories": categories,
        "panels": [{"title": "PV of debt-to-GDP ratio", "series": default_series}],
    }
    shock_payload = {
        "categories": categories,
        "panels": [{"title": "PV of debt-to-GDP ratio", "series": shock_series}],
    }
    return {
        "schema": 3,
        "pct_min": -0.5,
        "pct_max": 0.5,
        "pct_step": 0.5,
        "default": {"pct": 0.0, "payload": default_payload},
        "shocks": [{"pct": -0.5, "payload": shock_payload}],
    }


def _style_block() -> str:
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    match = re.search(r"<style>\s*(.*?)\s*</style>", template, re.DOTALL)
    if not match:
        raise AssertionError("Could not extract layout stylesheet.")
    return match.group(1)


def _browser_executable() -> Path | None:
    for candidate in _BROWSER_CANDIDATES:
        if candidate.is_file():
            return candidate
    return None


@unittest.skipUnless(importlib.util.find_spec("playwright"), "playwright is not installed")
class ChartPixelTests(unittest.TestCase):
    def test_selected_focal_lines_rasterize_to_legend_colors(self) -> None:
        browser_path = _browser_executable()
        if browser_path is None:
            self.skipTest("No local Chromium-based browser found.")

        from playwright.sync_api import sync_playwright

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <style>{_style_block()}</style>
</head>
<body>
  {build_chart_html(_sample_cache_doc())}
</body>
</html>
"""

        with sync_playwright() as p:
            browser = p.chromium.launch(executable_path=str(browser_path), headless=True)
            try:
                page = browser.new_page(viewport={"width": 800, "height": 400, "device_scale_factor": 2})
                page.set_content(html, wait_until="load")
                result = page.locator("svg.figure-panel").first.evaluate(
                    """async (svg) => {
                        const expected = {
                          "Baseline": [75, 130, 173],
                          "Historical scenario": [255, 0, 0],
                          "MX shock Standard&Tailored": [0, 0, 0],
                          "Threshold": [51, 153, 102],
                        };
                        const nonFocalExpected = [196, 196, 196];
                        const tolerance = 40;
                        const radius = 3;

                        const clone = svg.cloneNode(true);
                        const sourceLines = Array.from(svg.querySelectorAll("polyline.shock-line"));
                        const cloneLines = Array.from(clone.querySelectorAll("polyline.shock-line"));
                        sourceLines.forEach((source, index) => {
                          const target = cloneLines[index];
                          const style = getComputedStyle(source);
                          target.setAttribute("stroke", style.stroke);
                          target.setAttribute("stroke-width", style.strokeWidth);
                          const dash = style.strokeDasharray;
                          if (dash && dash !== "none") {
                            target.setAttribute("stroke-dasharray", dash.replaceAll(",", " "));
                          } else {
                            target.removeAttribute("stroke-dasharray");
                          }
                        });

                        const svgText = new XMLSerializer().serializeToString(clone);
                        const img = await new Promise((resolve, reject) => {
                          const value = new Image();
                          value.onload = () => resolve(value);
                          value.onerror = () => reject(new Error("Could not rasterize SVG"));
                          value.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgText);
                        });

                        const width = Math.round(svg.viewBox.baseVal.width || 520);
                        const height = Math.round(svg.viewBox.baseVal.height || 260);
                        const canvas = document.createElement("canvas");
                        canvas.width = width;
                        canvas.height = height;
                        const ctx = canvas.getContext("2d");
                        if (!ctx) {
                          throw new Error("2D canvas not available");
                        }
                        ctx.drawImage(img, 0, 0, width, height);

                        const colorNear = (pixel, expectedRgb) =>
                          Math.abs(pixel[0] - expectedRgb[0]) <= tolerance &&
                          Math.abs(pixel[1] - expectedRgb[1]) <= tolerance &&
                          Math.abs(pixel[2] - expectedRgb[2]) <= tolerance &&
                          pixel[3] >= 64;

                        const sampleAlongSeries = (seriesName, expectedRgb) => {
                          const line = svg.querySelector(
                            `g.shock-layer.shock-selected polyline[data-series-name="${seriesName.replaceAll('"', '\\"')}"]`
                          );
                          if (!line) {
                            return false;
                          }
                          const points = (line.getAttribute("points") || "")
                            .trim()
                            .split(/\\s+/)
                            .map((pair) => pair.split(",").map(Number))
                            .filter((pair) => pair.length === 2 && Number.isFinite(pair[0]) && Number.isFinite(pair[1]));
                          for (let i = 0; i < points.length - 1; i += 1) {
                            const [x0, y0] = points[i];
                            const [x1, y1] = points[i + 1];
                            for (let step = 0; step <= 8; step += 1) {
                              const t = step / 8;
                              const x = x0 + (x1 - x0) * t;
                              const y = y0 + (y1 - y0) * t;
                              for (let dx = -radius; dx <= radius; dx += 1) {
                                for (let dy = -radius; dy <= radius; dy += 1) {
                                  const xi = Math.round(x + dx);
                                  const yi = Math.round(y + dy);
                                  if (xi < 0 || yi < 0 || xi >= width || yi >= height) {
                                    continue;
                                  }
                                  const pixel = Array.from(ctx.getImageData(xi, yi, 1, 1).data);
                                  if (colorNear(pixel, expectedRgb)) {
                                    return true;
                                  }
                                }
                              }
                            }
                          }
                          return false;
                        };

                        return {
                          baseline: sampleAlongSeries("Baseline", expected["Baseline"]),
                          historical: sampleAlongSeries("Historical scenario", expected["Historical scenario"]),
                          mxShock: sampleAlongSeries("MX shock Standard&Tailored", expected["MX shock Standard&Tailored"]),
                          threshold: sampleAlongSeries("Threshold", expected["Threshold"]),
                          nonFocalGray: sampleAlongSeries(
                            "MX value, 1 yr only shock Standard&Tailored - for chart",
                            nonFocalExpected
                          ),
                        };
                    }"""
                )
            finally:
                browser.close()

        self.assertEqual(
            result,
            {
                "baseline": True,
                "historical": True,
                "mxShock": True,
                "threshold": True,
                "nonFocalGray": True,
            },
        )

    def test_background_layers_keep_series_hues_with_transparency(self) -> None:
        browser_path = _browser_executable()
        if browser_path is None:
            self.skipTest("No local Chromium-based browser found.")

        from playwright.sync_api import sync_playwright

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8"/>
  <style>{_style_block()}</style>
</head>
<body>
  {build_chart_html(_sample_cache_doc())}
</body>
</html>
"""

        with sync_playwright() as p:
            browser = p.chromium.launch(executable_path=str(browser_path), headless=True)
            try:
                page = browser.new_page(viewport={"width": 800, "height": 400, "device_scale_factor": 2})
                page.set_content(html, wait_until="load")
                result = page.locator("svg.figure-panel").first.evaluate(
                    """async (svg) => {
                        const tolerance = 45;
                        const radius = 3;
                        const expectedBackground = {
                          "Baseline": [190, 210, 225],
                          "Historical scenario": [255, 163, 163],
                          "Threshold": [182, 218, 200],
                        };

                        const clone = svg.cloneNode(true);
                        const sourceLines = Array.from(svg.querySelectorAll("polyline.shock-line"));
                        const cloneLines = Array.from(clone.querySelectorAll("polyline.shock-line"));
                        sourceLines.forEach((source, index) => {
                          const target = cloneLines[index];
                          const style = getComputedStyle(source);
                          target.setAttribute("stroke", style.stroke);
                          target.setAttribute("stroke-width", style.strokeWidth);
                          target.setAttribute("stroke-opacity", style.strokeOpacity);
                          const dash = style.strokeDasharray;
                          if (dash && dash !== "none") {
                            target.setAttribute("stroke-dasharray", dash.replaceAll(",", " "));
                          } else {
                            target.removeAttribute("stroke-dasharray");
                          }
                        });

                        const svgText = new XMLSerializer().serializeToString(clone);
                        const img = await new Promise((resolve, reject) => {
                          const value = new Image();
                          value.onload = () => resolve(value);
                          value.onerror = () => reject(new Error("Could not rasterize SVG"));
                          value.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svgText);
                        });

                        const width = Math.round(svg.viewBox.baseVal.width || 520);
                        const height = Math.round(svg.viewBox.baseVal.height || 260);
                        const canvas = document.createElement("canvas");
                        canvas.width = width;
                        canvas.height = height;
                        const ctx = canvas.getContext("2d");
                        if (!ctx) {
                          throw new Error("2D canvas not available");
                        }
                        ctx.drawImage(img, 0, 0, width, height);

                        const colorNear = (pixel, expectedRgb) =>
                          Math.abs(pixel[0] - expectedRgb[0]) <= tolerance &&
                          Math.abs(pixel[1] - expectedRgb[1]) <= tolerance &&
                          Math.abs(pixel[2] - expectedRgb[2]) <= tolerance &&
                          pixel[3] >= 48;

                        const sampleAlongSeries = (layerPct, seriesName, expectedRgb) => {
                          const line = svg.querySelector(
                            `g.shock-layer[data-pct="${layerPct}"] polyline[data-series-name="${seriesName.replaceAll('"', '\\"')}"]`
                          );
                          if (!line) {
                            return false;
                          }
                          const points = (line.getAttribute("points") || "")
                            .trim()
                            .split(/\\s+/)
                            .map((pair) => pair.split(",").map(Number))
                            .filter((pair) => pair.length === 2 && Number.isFinite(pair[0]) && Number.isFinite(pair[1]));
                          for (let i = 0; i < points.length - 1; i += 1) {
                            const [x0, y0] = points[i];
                            const [x1, y1] = points[i + 1];
                            for (let step = 0; step <= 8; step += 1) {
                              const t = step / 8;
                              const x = x0 + (x1 - x0) * t;
                              const y = y0 + (y1 - y0) * t;
                              for (let dx = -radius; dx <= radius; dx += 1) {
                                for (let dy = -radius; dy <= radius; dy += 1) {
                                  const xi = Math.round(x + dx);
                                  const yi = Math.round(y + dy);
                                  if (xi < 0 || yi < 0 || xi >= width || yi >= height) {
                                    continue;
                                  }
                                  const pixel = Array.from(ctx.getImageData(xi, yi, 1, 1).data);
                                  if (colorNear(pixel, expectedRgb)) {
                                    return true;
                                  }
                                }
                              }
                            }
                          }
                          return false;
                        };

                        return {
                          baselineTint: sampleAlongSeries("-0.5", "Baseline", expectedBackground["Baseline"]),
                          historicalTint: sampleAlongSeries("-0.5", "Historical scenario", expectedBackground["Historical scenario"]),
                          thresholdTint: sampleAlongSeries("-0.5", "Threshold", expectedBackground["Threshold"]),
                        };
                    }"""
                )
            finally:
                browser.close()

        self.assertEqual(
            result,
            {
                "baselineTint": True,
                "historicalTint": True,
                "thresholdTint": True,
            },
        )


if __name__ == "__main__":
    unittest.main()
