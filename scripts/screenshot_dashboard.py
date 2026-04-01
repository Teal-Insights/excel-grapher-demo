#!/usr/bin/env python3
"""Start the Figure 1 app briefly and save a full-page PNG (for README / docs)."""

from __future__ import annotations

import argparse
import os
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return int(s.getsockname()[1])


def _wait_http(url: str, timeout_s: float) -> None:
    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=2) as r:
                if r.status == 200:
                    return
        except (urllib.error.URLError, TimeoutError, OSError):
            time.sleep(0.15)
    raise TimeoutError(f"No HTTP 200 from {url} within {timeout_s}s")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument(
        "--cache",
        type=Path,
        default=_ROOT / ".cache" / "dsf-uga-gdp-shocks.json",
        help="GDP shock JSON passed as GDP_SHOCK_CACHE (default: .cache/dsf-uga-gdp-shocks.json).",
    )
    ap.add_argument(
        "--out",
        type=Path,
        default=_ROOT / "docs" / "dashboard.png",
        help="Output PNG path (default: docs/dashboard.png).",
    )
    ap.add_argument(
        "--timeout",
        type=float,
        default=45.0,
        help="Seconds to wait for the server to respond (default: 45).",
    )
    args = ap.parse_args()
    cache = args.cache.expanduser().resolve()
    if not cache.is_file():
        raise SystemExit(f"Cache file not found: {cache}")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise SystemExit(
            "Playwright is required: uv sync --group dev && uv run playwright install chromium"
        ) from e

    port = _free_port()
    url = f"http://127.0.0.1:{port}/"
    env = os.environ.copy()
    env["GDP_SHOCK_CACHE"] = str(cache)

    proc = subprocess.Popen(
        [sys.executable, "-m", "uvicorn", "main:app", "--host", "127.0.0.1", f"--port={port}"],
        cwd=_ROOT,
        env=env,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
        text=True,
    )
    try:
        _wait_http(url, args.timeout)
        args.out.parent.mkdir(parents=True, exist_ok=True)
        with sync_playwright() as p:
            browser = p.chromium.launch()
            try:
                page = browser.new_page(viewport={"width": 1280, "height": 900})
                page.goto(url, wait_until="networkidle", timeout=int(args.timeout * 1000))
                page.screenshot(path=str(args.out), full_page=True)
            finally:
                browser.close()
        print(args.out)
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()


if __name__ == "__main__":
    main()
