from __future__ import annotations

import unittest
from pathlib import Path
from datetime import datetime
from random import Random
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook

from outreach.batch_schedule import (
    MIN_RANDOM_SPACING_MINUTES,
    apply_batch_schedule,
    build_batch_schedule,
    build_random_batch_schedule,
    maximum_random_spacing_that_fits_window,
    maximum_spacing_that_fits_window,
)


class BatchScheduleTests(unittest.TestCase):
    def test_build_batch_schedule_assigns_spacing_inside_window(self) -> None:
        assignments = build_batch_schedule(
            ["TEST-001", "TEST-002", "TEST-003"],
            datetime(2026, 8, 18, 9, 0),
            datetime(2026, 8, 18, 10, 0),
            20,
        )

        self.assertEqual(
            [
                ("TEST-001", "2026-08-18 09:00"),
                ("TEST-002", "2026-08-18 09:20"),
                ("TEST-003", "2026-08-18 09:40"),
            ],
            [(item.test_id, item.scheduled_send_at) for item in assignments],
        )

    def test_build_batch_schedule_rejects_window_that_is_too_short(self) -> None:
        with self.assertRaisesRegex(ValueError, "do not fit"):
            build_batch_schedule(
                ["TEST-001", "TEST-002", "TEST-003"],
                datetime(2026, 8, 18, 9, 0),
                datetime(2026, 8, 18, 9, 30),
                20,
            )

    def test_build_random_batch_schedule_stays_inside_window_and_spacing_bounds(self) -> None:
        assignments = build_random_batch_schedule(
            ["TEST-001", "TEST-002", "TEST-003", "TEST-004"],
            datetime(2026, 8, 18, 9, 0),
            datetime(2026, 8, 18, 10, 0),
            20,
            rng=Random(7),
        )

        times = [
            datetime.strptime(item.scheduled_send_at, "%Y-%m-%d %H:%M")
            for item in assignments
        ]
        gaps = [
            int((times[index + 1] - times[index]).total_seconds() // 60)
            for index in range(len(times) - 1)
        ]
        self.assertEqual(datetime(2026, 8, 18, 9, 0), times[0])
        self.assertLessEqual(times[-1], datetime(2026, 8, 18, 10, 0))
        self.assertTrue(all(MIN_RANDOM_SPACING_MINUTES <= gap <= 20 for gap in gaps))

    def test_build_random_batch_schedule_rejects_max_spacing_that_cannot_fit(self) -> None:
        with self.assertRaisesRegex(ValueError, "Choose 27 minutes or less"):
            build_random_batch_schedule(
                ["TEST-001", "TEST-002", "TEST-003"],
                datetime(2026, 8, 18, 9, 0),
                datetime(2026, 8, 18, 9, 30),
                30,
                rng=Random(7),
            )

    def test_build_random_batch_schedule_rejects_window_below_minimum_spacing(self) -> None:
        with self.assertRaisesRegex(ValueError, "too short"):
            build_random_batch_schedule(
                ["TEST-001", "TEST-002", "TEST-003"],
                datetime(2026, 8, 18, 9, 0),
                datetime(2026, 8, 18, 9, 5),
                3,
                rng=Random(7),
            )

    def test_build_random_batch_schedule_rejects_cap_below_minimum_spacing(self) -> None:
        with self.assertRaisesRegex(ValueError, "at least 3"):
            build_random_batch_schedule(
                ["TEST-001", "TEST-002", "TEST-003"],
                datetime(2026, 8, 18, 9, 0),
                datetime(2026, 8, 18, 9, 30),
                2,
                rng=Random(7),
            )

    def test_build_random_batch_schedule_enforces_minimum_spacing(self) -> None:
        assignments = build_random_batch_schedule(
            ["TEST-001", "TEST-002", "TEST-003"],
            datetime(2026, 8, 18, 9, 0),
            datetime(2026, 8, 18, 9, 20),
            3,
            rng=Random(7),
        )

        times = [
            datetime.strptime(item.scheduled_send_at, "%Y-%m-%d %H:%M")
            for item in assignments
        ]
        gaps = [
            int((times[index + 1] - times[index]).total_seconds() // 60)
            for index in range(len(times) - 1)
        ]
        self.assertTrue(all(gap >= MIN_RANDOM_SPACING_MINUTES for gap in gaps))

    def test_build_random_batch_schedule_uses_full_window_when_cap_allows(self) -> None:
        assignments = build_random_batch_schedule(
            ["TEST-001", "TEST-002", "TEST-003"],
            datetime(2026, 8, 18, 9, 0),
            datetime(2026, 8, 18, 9, 30),
            27,
            rng=Random(7),
        )

        times = [
            datetime.strptime(item.scheduled_send_at, "%Y-%m-%d %H:%M")
            for item in assignments
        ]
        gaps = [
            int((times[index + 1] - times[index]).total_seconds() // 60)
            for index in range(len(times) - 1)
        ]
        self.assertEqual(datetime(2026, 8, 18, 9, 30), times[-1])
        self.assertEqual(30, sum(gaps))
        self.assertTrue(all(MIN_RANDOM_SPACING_MINUTES <= gap <= 27 for gap in gaps))

    def test_build_random_batch_schedule_accepts_dense_three_minute_gaps(self) -> None:
        test_ids = [f"TEST-{index:03d}" for index in range(1, 27)]

        assignments = build_random_batch_schedule(
            test_ids,
            datetime(2026, 8, 31, 15, 45),
            datetime(2026, 8, 31, 17, 0),
            3,
            rng=Random(7),
        )

        times = [
            datetime.strptime(item.scheduled_send_at, "%Y-%m-%d %H:%M")
            for item in assignments
        ]
        gaps = [
            int((times[index + 1] - times[index]).total_seconds() // 60)
            for index in range(len(times) - 1)
        ]
        self.assertEqual(26, len(assignments))
        self.assertLessEqual(times[-1], datetime(2026, 8, 31, 17, 0))
        self.assertTrue(all(gap == 3 for gap in gaps))

    def test_maximum_spacing_that_fits_window_uses_gap_count(self) -> None:
        self.assertEqual(
            20,
            maximum_spacing_that_fits_window(
                4,
                datetime(2026, 8, 18, 9, 0),
                datetime(2026, 8, 18, 10, 1),
            ),
        )

    def test_maximum_random_spacing_that_fits_window_allows_one_large_gap(self) -> None:
        self.assertEqual(
            54,
            maximum_random_spacing_that_fits_window(
                4,
                datetime(2026, 8, 18, 9, 0),
                datetime(2026, 8, 18, 10, 0),
            ),
        )

    def test_maximum_spacing_that_fits_window_allows_dense_sub_five_minute_caps(self) -> None:
        self.assertEqual(
            3,
            maximum_spacing_that_fits_window(
                26,
                datetime(2026, 8, 31, 15, 45),
                datetime(2026, 8, 31, 17, 0),
            ),
        )

    def test_maximum_spacing_that_fits_window_returns_zero_without_gaps(self) -> None:
        self.assertEqual(
            0,
            maximum_spacing_that_fits_window(
                1,
                datetime(2026, 8, 18, 9, 0),
                datetime(2026, 8, 18, 10, 0),
            ),
        )

    def test_apply_batch_schedule_writes_assignments_with_one_save(self) -> None:
        with TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "review.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Email Review Queue"
            sheet.append(["Test ID", "Scheduled Send At"])
            sheet.append(["TEST-001", ""])
            sheet.append(["TEST-002", ""])
            workbook.save(path)
            workbook.close()

            with patch("outreach.review_queue.save_workbook_atomically") as save:
                assignments = apply_batch_schedule(
                    path,
                    ["TEST-001", "TEST-002"],
                    datetime(2026, 8, 18, 9, 0),
                    datetime(2026, 8, 18, 10, 0),
                    10,
                )

            self.assertEqual(["2026-08-18 09:00", "2026-08-18 09:10"], [item.scheduled_send_at for item in assignments])
            save.assert_called_once()


if __name__ == "__main__":
    unittest.main()
