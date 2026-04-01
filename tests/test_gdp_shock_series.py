from __future__ import annotations

import unittest

from lic_dsf.payload import gdp_forecast_series_from_percent


class GdpShockSeriesTests(unittest.TestCase):
    def test_zero_percent_reproduces_baseline_series(self) -> None:
        baselines = [100.0, 110.0, 121.0, 133.1]

        shocked = gdp_forecast_series_from_percent(baselines, 0.0)

        self.assertEqual(len(shocked), len(baselines))
        for actual, expected in zip(shocked, baselines, strict=True):
            self.assertAlmostEqual(actual, expected)

    def test_positive_shock_adds_percentage_points_to_growth_rates(self) -> None:
        baselines = [100.0, 110.0, 121.0]

        shocked = gdp_forecast_series_from_percent(baselines, 5.0)

        self.assertEqual(shocked[0], 100.0)
        self.assertAlmostEqual(shocked[1], 115.0)
        self.assertAlmostEqual(shocked[2], 132.25)

    def test_negative_shock_rebuilds_levels_from_shocked_prior_values(self) -> None:
        baselines = [200.0, 220.0, 242.0]

        shocked = gdp_forecast_series_from_percent(baselines, -5.0)

        self.assertEqual(shocked[0], 200.0)
        self.assertAlmostEqual(shocked[1], 210.0)
        self.assertAlmostEqual(shocked[2], 220.5)

    def test_zero_baseline_anchor_keeps_that_step_at_baseline(self) -> None:
        baselines = [0.0, 10.0, 15.0]

        shocked = gdp_forecast_series_from_percent(baselines, 5.0)

        self.assertEqual(shocked[0], 0.0)
        self.assertEqual(shocked[1], 10.0)
        self.assertAlmostEqual(shocked[2], 15.5)


if __name__ == "__main__":
    unittest.main()
