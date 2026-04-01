from __future__ import annotations

from pathlib import Path
from typing import Any

from web.charts import load_shock_json, payloads_by_shock

STANDARD_CACHE_BACKENDS = ("excel-grapher", "xlwings", "libreoffice")


def default_backend_cache_paths(repo_root: Path) -> dict[str, Path]:
    cache_dir = repo_root / ".cache"
    return {
        backend: cache_dir / f"gdp-shocks-{backend}.json"
        for backend in STANDARD_CACHE_BACKENDS
    }


def available_backend_cache_paths(repo_root: Path) -> dict[str, Path]:
    return {
        backend: path
        for backend, path in default_backend_cache_paths(repo_root).items()
        if path.is_file()
    }


def compare_cache_docs(left: dict[str, Any], right: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    for key in ("pct_min", "pct_max", "pct_step"):
        if left.get(key) != right.get(key):
            errors.append(
                f"{key} differs: left={left.get(key)!r}, right={right.get(key)!r}"
            )

    left_payloads = payloads_by_shock(left)
    right_payloads = payloads_by_shock(right)
    left_shocks = set(left_payloads)
    right_shocks = set(right_payloads)

    missing_from_right = sorted(left_shocks - right_shocks)
    if missing_from_right:
        errors.append(f"missing shocks on right: {missing_from_right}")

    missing_from_left = sorted(right_shocks - left_shocks)
    if missing_from_left:
        errors.append(f"missing shocks on left: {missing_from_left}")

    differing = [
        pct
        for pct in sorted(left_shocks & right_shocks)
        if left_payloads[pct] != right_payloads[pct]
    ]
    if differing:
        sample = ", ".join(f"{pct:g}" for pct in differing[:5])
        suffix = " ..." if len(differing) > 5 else ""
        errors.append(
            f"payload differs at {len(differing)} shock levels: {sample}{suffix}"
        )

    return errors


def compare_cache_files(left_path: Path, right_path: Path) -> list[str]:
    left = load_shock_json(left_path)
    right = load_shock_json(right_path)
    return compare_cache_docs(left, right)
