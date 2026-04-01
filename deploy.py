"""
Deploy the LIC-DSF Figure 1 FastAPI app to Modal.

From the repository root:

  uv run modal deploy deploy.py

Ensure the GDP shock cache exists (default path matches main.py), e.g.:

  uv run python scripts/precache.py --backend excel-grapher

Optional: attach ``modal.Secret.from_name(...)`` on ``app`` for env vars such as
``GDP_SHOCK_CACHE`` if you point the app at a different JSON path.
"""

from __future__ import annotations

import modal

# uv_sync respects git sources in uv.lock (e.g. excel-grapher). pip_install_from_pyproject
# does not read [tool.uv.sources], so plain pip would not resolve that dependency.
image = (
    modal.Image.debian_slim(python_version="3.13")
    .apt_install("git")
    .uv_sync()
    .add_local_python_source("deploy", "main", "web", "lic_dsf")
    .add_local_dir("templates", remote_path="/root/templates")
    .add_local_dir(
        ".cache",
        remote_path="/root/.cache",
        ignore=["*.pkl", "*.xlsm"],
    )
)

app = modal.App(
    name="lic-dsf-figure1",
    image=image,
    include_source=False,
)


@app.function()
@modal.asgi_app()
def serve():
    from main import app as web_app

    return web_app
