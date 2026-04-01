from __future__ import annotations

import unittest
from itertools import combinations
from pathlib import Path

from lic_dsf.cache_compare import available_backend_cache_paths, compare_cache_files


class BackendCacheParityTest(unittest.TestCase):
    def test_available_backend_caches_match(self) -> None:
        repo_root = Path(__file__).resolve().parents[1]
        caches = available_backend_cache_paths(repo_root)
        if len(caches) < 2:
            self.skipTest("Need cache files for at least two backends.")

        mismatches: list[str] = []
        for left_name, right_name in combinations(sorted(caches), 2):
            diffs = compare_cache_files(caches[left_name], caches[right_name])
            mismatches.extend(
                f"{left_name} vs {right_name}: {diff}" for diff in diffs
            )

        self.assertEqual([], mismatches, "\n".join(mismatches))


if __name__ == "__main__":
    unittest.main()
