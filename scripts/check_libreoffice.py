#!/usr/bin/env python3
"""
Apply GDP forecast shocks to a workbook, recalculate via a workbook engine, and compare chart output cells.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from lic_dsf import graph
from lic_dsf.libreoffice import payloads_from_precache_json
from lic_dsf.workbook_backend import (
    print_check_report,
    resolve_backend,
    run_workbook_gdp_shock_check,
)


def main() -> None:
    ap = argparse.ArgumentParser(
        description="GDP shock via workbook recalc; optional FormulaEvaluator vs workbook-engine accuracy check."
    )
    ap.add_argument(
        "--workbook",
        type=Path,
        default=graph.WORKBOOK_PATH,
        help=f"Source .xlsm (default: {graph.WORKBOOK_PATH}).",
    )
    ap.add_argument(
        "--baseline-bps",
        type=int,
        default=0,
        help="bps applied in the reference copy (default 0).",
    )
    ap.add_argument(
        "--shock-bps",
        type=int,
        default=10,
        help="bps applied in the comparison copy (default 10).",
    )
    ap.add_argument(
        "--backend",
        choices=("auto", "libreoffice", "xlwings"),
        default="auto",
        help="Sanity-check engine: auto picks xlwings on Windows and LibreOffice on Linux.",
    )
    ap.add_argument(
        "--python-precache-json",
        type=Path,
        default=Path(".cache/gdp-shocks.json"),
        help="Precache JSON to compare FormulaEvaluator vs workbook backend.",
    )
    ap.add_argument(
        "--timeout",
        type=int,
        default=600,
        help="LibreOffice convert timeout per invocation (seconds) when that backend is used.",
    )
    ap.add_argument(
        "--soffice",
        type=str,
        default=None,
        help="Path or name of soffice/libreoffice (default: search PATH).",
    )
    ap.add_argument(
        "--top-n",
        type=int,
        default=15,
        help="Number of largest error rows to print / embed (default 15).",
    )
    ap.add_argument(
        "--keep-temps",
        action="store_true",
        help="Print temp directory path and do not delete it.",
    )
    args = ap.parse_args()

    py_base = py_shock = None
    if args.python_precache_json is not None:
        p = args.python_precache_json.resolve()
        if not p.is_file():
            raise SystemExit(f"Precache JSON not found: {p}")
        py_base, py_shock = payloads_from_precache_json(
            p,
            baseline_bps=args.baseline_bps,
            shock_bps=args.shock_bps,
        )

    result = run_workbook_gdp_shock_check(
        args.workbook,
        backend=resolve_backend(args.backend),
        baseline_bps=args.baseline_bps,
        shock_bps=args.shock_bps,
        timeout_s=args.timeout,
        soffice=args.soffice,
        keep_temps=args.keep_temps,
        top_n=args.top_n,
        python_baseline_payload=py_base,
        python_shock_payload=py_shock,
    )
    print_check_report(result)
    if not result.get("ok"):
        sys.exit(1)


if __name__ == "__main__":
    main()
