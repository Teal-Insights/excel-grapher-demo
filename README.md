# py-lic-dsf-demo

Precomputed **Figure 1 — external stress** charts from an LIC-DSF-style Excel workbook. A FastAPI app serves static SVG layers and JSON so the GDP shock slider only toggles visibility in the browser (no workbook evaluation per request).

## Prerequisites

- [uv](https://docs.astral.sh/uv/) and Python 3.13+
- A compatible `.xlsm` workbook and dependency graph (see `lic_dsf/graph.py` for defaults and graph cache under `.cache/`, which is gitignored)

## Build the GDP shock cache

From the repo root, with `PYTHONPATH=.` so `lic_dsf` imports resolve:

```bash
PYTHONPATH=. uv run python scripts/precache.py --backend excel-grapher --workbook lic-dsf-template-2025-08-12.xlsm
```

`precache.py` now builds one backend-specific cache per run:

- `--backend excel-grapher` writes `.cache/gdp-shocks-excel-grapher.json`
- `--backend xlwings` writes `.cache/gdp-shocks-xlwings.json`
- `--backend libreoffice` writes `.cache/gdp-shocks-libreoffice.json`
- `xlwings` is optional and only needed when you build the `xlwings` cache.
- LibreOffice is optional and only needed when you build the `libreoffice` cache.

Precache schema v3 stores GDP shocks as **percent** (not basis points): default **0%**, and one entry per step from **-5%** to **+5%** in **0.5%** increments (`lic_dsf.payload.GDP_SHOCK_PCT_*`).

Common options:

| Flag | Purpose |
|------|---------|
| `--backend excel-grapher\|xlwings\|libreoffice` | Select the calculation engine used to build the cache |
| `--workbook PATH` | Source `.xlsm` (default: workbook path from `lic_dsf.graph`) |
| `--out PATH` | Output JSON (default: `.cache/gdp-shocks-<backend>.json`) |
| `--no-graph-cache` | Rebuild the dependency graph instead of using the pickle cache (`excel-grapher` only) |
| `--soffice PATH` | Explicit LibreOffice binary when `--backend libreoffice` |
| `--timeout N` | LibreOffice conversion timeout in seconds |

The web app’s default cache path is `.cache/gdp-shocks-excel-grapher.json` unless you override it (see below).

## Run the dashboard

**Option A — `main` entrypoint (recommended for local use)**

```bash
uv run python main.py
```

Optional cache file:

```bash
uv run python main.py --cache .cache/gdp-shocks-xlwings.json
```

**Option B — Uvicorn directly**

```bash
GDP_SHOCK_CACHE=.cache/gdp-shocks-excel-grapher.json uv run uvicorn main:app --host 127.0.0.1 --port 8000 --reload
```

Environment variables (used when `--cache` is not set and you are not using `python main.py`):

- `GDP_SHOCK_CACHE` — path to the precache JSON (alias: `FIGURE1_SHOCK_CACHE`)

Then open [http://127.0.0.1:8000/](http://127.0.0.1:8000/).

**API**

- `GET /api/figure1-data` — slim JSON for the slider and charts (or `503` with an `error` field if the cache is missing or invalid)

## Screenshot

Dashboard with a populated cache:

![Figure 1 external stress dashboard](README_files/dashboard.png)

## Regenerate the screenshot

One-time browser install for the dev toolchain:

```bash
uv sync --group dev
uv run playwright install chromium
```

Capture `README_files/dashboard.png` (requires an existing cache file):

```bash
uv run python scripts/screenshot_dashboard.py --cache .cache/gdp-shocks-excel-grapher.json
```

Use `--out` to write a different path.
