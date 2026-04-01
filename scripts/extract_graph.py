#!/usr/bin/env python3
"""CLI entry for dependency graph extraction."""

import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from lic_dsf.graph import main


if __name__ == "__main__":
    main()
