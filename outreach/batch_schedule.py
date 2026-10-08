from __future__ import annotations

import random
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

from outreach.review_queue import set_scheduled_send_at_many


MIN_RANDOM_SPACING_MINUTES = 3


@dataclass(frozen=True)
class ScheduleAssignment:
    test_id: str
    scheduled_send_at: str


def build_batch_schedule(
    test_ids: list[str],
    window_start: datetime,
    window_end: datetime,
    spacing_minutes: int,
) -> list[ScheduleAssignment]:
    if spacing_minutes <= 0:
        raise ValueError("Spacing must be greater than zero minutes.")
    if window_end < window_start:
        raise ValueError("Schedule window end must be after the start.")
    if not test_ids:
        return []

    gap_count = max(0, len(test_ids) - 1)
    max_spacing_that_fits = maximum_spacing_that_fits_window(len(test_ids), window_start, window_end)
    if gap_count and spacing_minutes > max_spacing_that_fits:
        raise ValueError(
            f"{len(test_ids)} email(s) do not fit in the selected window with "
            f"{spacing_minutes} minute spacing."
        )
    spacings = [spacing_minutes] * gap_count
    return _assign_schedule(test_ids, window_start, spacings)


def maximum_spacing_that_fits_window(
    item_count: int,
    window_start: datetime,
    window_end: datetime,
) -> int:
    if item_count < 0:
        raise ValueError("Item count must not be negative.")
    if window_end < window_start:
        raise ValueError("Schedule window end must be after the start.")

    gap_count = max(0, item_count - 1)
    if gap_count == 0:
        return 0

    available_minutes = int((window_end - window_start).total_seconds() // 60)
    return available_minutes // gap_count


def maximum_random_spacing_that_fits_window(
    item_count: int,
    window_start: datetime,
    window_end: datetime,
    min_spacing_minutes: int = MIN_RANDOM_SPACING_MINUTES,
) -> int:
    if item_count < 0:
        raise ValueError("Item count must not be negative.")
    if window_end < window_start:
        raise ValueError("Schedule window end must be after the start.")
    if min_spacing_minutes <= 0:
        raise ValueError("Minimum spacing must be greater than zero minutes.")

    gap_count = max(0, item_count - 1)
    if gap_count == 0:
        return 0

    available_minutes = int((window_end - window_start).total_seconds() // 60)
    minimum_total_spacing = gap_count * min_spacing_minutes
    if available_minutes < minimum_total_spacing:
        return 0
    return available_minutes - ((gap_count - 1) * min_spacing_minutes)


def build_random_batch_schedule(
    test_ids: list[str],
    window_start: datetime,
    window_end: datetime,
    max_spacing_minutes: int,
    min_spacing_minutes: int = MIN_RANDOM_SPACING_MINUTES,
    rng: random.Random | None = None,
) -> list[ScheduleAssignment]:
    if window_end < window_start:
        raise ValueError("Schedule window end must be after the start.")
    if min_spacing_minutes <= 0:
        raise ValueError("Minimum spacing must be greater than zero minutes.")
    if not test_ids:
        return []

    gap_count = len(test_ids) - 1
    if gap_count == 0:
        return _assign_schedule(test_ids, window_start, [])
    if max_spacing_minutes < min_spacing_minutes:
        raise ValueError(
            f"Maximum spacing must be at least {min_spacing_minutes} minute(s)."
        )

    available_minutes = int((window_end - window_start).total_seconds() // 60)
    max_spacing_that_fits = maximum_random_spacing_that_fits_window(
        len(test_ids),
        window_start,
        window_end,
        min_spacing_minutes,
    )
    if max_spacing_that_fits <= 0:
        raise ValueError("The selected window is too short for randomized spacing.")
    if max_spacing_minutes > max_spacing_that_fits:
        raise ValueError(
            f"{len(test_ids)} email(s) do not fit in the selected window with "
            f"maximum {max_spacing_minutes} minute spacing. Choose {max_spacing_that_fits} minutes or less."
        )

    random_source = rng or random.Random()
    total_spacing = min(available_minutes, gap_count * max_spacing_minutes)
    gaps = _random_bounded_gaps(
        gap_count,
        total_spacing,
        max_spacing_minutes,
        min_spacing_minutes,
        random_source,
    )
    return _assign_schedule(test_ids, window_start, gaps)


def _random_bounded_gaps(
    gap_count: int,
    total_minutes: int,
    max_spacing_minutes: int,
    min_spacing_minutes: int,
    random_source: random.Random,
) -> list[int]:
    if gap_count <= 0:
        return []
    if min_spacing_minutes <= 0:
        raise ValueError("Minimum spacing must be greater than zero minutes.")
    if max_spacing_minutes < min_spacing_minutes:
        raise ValueError("Maximum spacing must be at least the minimum spacing.")
    if total_minutes < gap_count * min_spacing_minutes:
        raise ValueError("The selected window is too short for randomized spacing.")
    if total_minutes > gap_count * max_spacing_minutes:
        raise ValueError("The selected window cannot fit within the selected maximum spacing.")

    gaps: list[int] = []
    remaining_total = total_minutes
    for index in range(gap_count):
        remaining_slots = gap_count - index - 1
        minimum_remaining = remaining_slots * min_spacing_minutes
        maximum_remaining = remaining_slots * max_spacing_minutes
        lower = max(min_spacing_minutes, remaining_total - maximum_remaining)
        upper = min(max_spacing_minutes, remaining_total - minimum_remaining)
        gap = random_source.randint(lower, upper)
        gaps.append(gap)
        remaining_total -= gap

    random_source.shuffle(gaps)
    return gaps


def _assign_schedule(
    test_ids: list[str],
    window_start: datetime,
    spacings: list[int],
) -> list[ScheduleAssignment]:
    assignments: list[ScheduleAssignment] = []
    scheduled_at = window_start
    for index, test_id in enumerate(test_ids):
        assignments.append(
            ScheduleAssignment(
                test_id=test_id,
                scheduled_send_at=scheduled_at.strftime("%Y-%m-%d %H:%M"),
            )
        )
        if index < len(spacings):
            scheduled_at += timedelta(minutes=spacings[index])
    return assignments


def apply_batch_schedule(
    review_workbook_path: Path,
    test_ids: list[str],
    window_start: datetime,
    window_end: datetime,
    spacing_minutes: int,
    randomize: bool = False,
    max_spacing_minutes: int | None = None,
) -> list[ScheduleAssignment]:
    if randomize:
        max_spacing = max_spacing_minutes
        if max_spacing is None:
            max_spacing = maximum_spacing_that_fits_window(len(test_ids), window_start, window_end)
        assignments = build_random_batch_schedule(
            test_ids,
            window_start,
            window_end,
            max_spacing,
        )
    else:
        assignments = build_batch_schedule(test_ids, window_start, window_end, spacing_minutes)
    set_scheduled_send_at_many(
        review_workbook_path,
        {assignment.test_id: assignment.scheduled_send_at for assignment in assignments},
    )
    return assignments
