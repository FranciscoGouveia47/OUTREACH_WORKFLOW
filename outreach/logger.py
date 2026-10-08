from __future__ import annotations

import csv
import re
import sys
from datetime import datetime
from pathlib import Path

from outreach.models import QueueEmail

csv.field_size_limit(sys.maxsize)

SENT_STATUS = "SENT"
UNKNOWN_STATUS = "SEND_STATUS_UNKNOWN"
DRAFT_CREATED_STATUS = "DRAFT_CREATED"
DRAFT_FAILED_STATUS = "DRAFT_FAILED"
OUTLOOK_SCHEDULED_STATUS = "OUTLOOK_SCHEDULED"
OUTLOOK_SCHEDULE_CANCELLED_STATUS = "OUTLOOK_SCHEDULE_CANCELLED"
OUTLOOK_OUTBOX_LOCKED_STATUS = "OUTLOOK_OUTBOX_LOCKED"
DELIVERY_FAILED_STATUS = "DELIVERY_FAILED"
DELIVERY_FAILURE_CLEARED_STATUS = "DELIVERY_FAILURE_CLEARED"
REFUSED_TO_ANSWER_STATUS = "REFUSED_TO_ANSWER"
REFUSED_TO_ANSWER_CLEARED_STATUS = "REFUSED_TO_ANSWER_CLEARED"
TRANSIENT_STATUSES = {"SENDING", "SCHEDULING"}
DRAFT_ENTRY_ID_RE = re.compile(r"^Outlook draft saved \((.+)\)$")
OUTLOOK_SCHEDULED_RE = re.compile(r"Outlook deferred send requested \((\d{4}-\d{2}-\d{2} \d{2}:\d{2})\)")
HEADERS = [
    "timestamp",
    "test_id",
    "target_name",
    "recipient_email",
    "recipient_type",
    "template_type",
    "language",
    "subject",
    "scheduled_send_at",
    "status",
    "graph_response",
    "error",
]


def already_sent(log_path: Path, test_id: str, recipient_email: str) -> bool:
    return has_status(log_path, test_id, recipient_email, {SENT_STATUS})


def has_unknown_send_status(log_path: Path, test_id: str, recipient_email: str) -> bool:
    return has_status(log_path, test_id, recipient_email, {UNKNOWN_STATUS})


def has_draft_created(log_path: Path, test_id: str, recipient_email: str) -> bool:
    return _active_draft_state(log_path, test_id, recipient_email)[0]


def has_outlook_scheduled(log_path: Path, test_id: str, recipient_email: str) -> bool:
    return latest_status(log_path, test_id, recipient_email) == OUTLOOK_SCHEDULED_STATUS


def has_outbox_locked(log_path: Path, test_id: str, recipient_email: str) -> bool:
    return latest_status(log_path, test_id, recipient_email) == OUTLOOK_OUTBOX_LOCKED_STATUS


def has_delivery_failed(log_path: Path, test_id: str, recipient_email: str) -> bool:
    return latest_status(log_path, test_id, recipient_email) == DELIVERY_FAILED_STATUS


def has_refused_to_answer(log_path: Path, test_id: str, recipient_email: str) -> bool:
    return latest_status(log_path, test_id, recipient_email) == REFUSED_TO_ANSWER_STATUS


def latest_status(log_path: Path, test_id: str, recipient_email: str) -> str:
    if not log_path.exists():
        return ""

    status = ""
    with log_path.open("r", encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            if (
                row.get("test_id") == test_id
                and row.get("recipient_email", "").lower() == recipient_email.lower()
            ):
                row_status = row.get("status", "") or ""
                if row_status not in TRANSIENT_STATUSES:
                    status = row_status
    return status


def latest_log_row(
    log_path: Path,
    test_id: str,
    recipient_email: str,
    statuses: set[str] | None = None,
) -> dict[str, str]:
    if not log_path.exists():
        return {}

    latest: dict[str, str] = {}
    with log_path.open("r", encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            if (
                row.get("test_id") == test_id
                and row.get("recipient_email", "").lower() == recipient_email.lower()
                and (statuses is None or row.get("status", "") in statuses)
            ):
                normalised = _normalise_row(row)
                if statuses is not None or normalised.get("status", "") not in TRANSIENT_STATUSES:
                    latest = normalised
    return latest


def latest_draft_entry_id(log_path: Path, test_id: str, recipient_email: str) -> str:
    return _active_draft_state(log_path, test_id, recipient_email)[1]


def latest_draft_entry_ids(log_path: Path, test_id: str, recipient_email: str = "") -> list[str]:
    if not log_path.exists():
        return []

    entry_ids: list[str] = []
    with log_path.open("r", encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            if row.get("test_id") != test_id:
                continue
            if recipient_email and row.get("recipient_email", "").lower() != recipient_email.lower():
                continue
            if row.get("status") != DRAFT_CREATED_STATUS:
                continue
            entry_id = draft_entry_id_from_response(row.get("graph_response", ""))
            if entry_id:
                entry_ids.append(entry_id)
    return entry_ids


def _active_draft_state(log_path: Path, test_id: str, recipient_email: str) -> tuple[bool, str]:
    if not log_path.exists():
        return False, ""

    active = False
    entry_id = ""
    draft_deactivating_statuses = {
        SENT_STATUS,
        OUTLOOK_SCHEDULED_STATUS,
        OUTLOOK_SCHEDULE_CANCELLED_STATUS,
    }
    with log_path.open("r", encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            if (
                row.get("test_id") != test_id
                or row.get("recipient_email", "").lower() != recipient_email.lower()
            ):
                continue

            status = row.get("status", "")
            if status == DRAFT_CREATED_STATUS:
                active = True
                entry_id = draft_entry_id_from_response(row.get("graph_response", ""))
            elif status in draft_deactivating_statuses:
                active = False
                entry_id = ""
    return active, entry_id


def remove_log_entries(log_path: Path, test_id: str, recipient_email: str = "") -> int:
    if not log_path.exists():
        return 0

    with log_path.open("r", encoding="utf-8", newline="") as handle:
        rows = [_normalise_row(row) for row in csv.DictReader(handle)]

    kept_rows = []
    removed = 0
    for row in rows:
        matches_test_id = row.get("test_id") == test_id
        matches_recipient = not recipient_email or row.get("recipient_email", "").lower() == recipient_email.lower()
        if matches_test_id and matches_recipient:
            removed += 1
        else:
            kept_rows.append(row)

    with log_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=HEADERS)
        writer.writeheader()
        writer.writerows(kept_rows)
    return removed


def draft_entry_id_from_response(response: str) -> str:
    match = DRAFT_ENTRY_ID_RE.match(response.strip())
    if not match:
        return ""
    return match.group(1)


def scheduled_send_at_from_response(response: str) -> str:
    match = OUTLOOK_SCHEDULED_RE.search(response.strip())
    if not match:
        return ""
    return match.group(1)


def clean_scheduled_send_at(value: str | None) -> str:
    cleaned = (value or "").strip()
    if cleaned.lower() in {"none", "null", "nat"}:
        return ""
    return cleaned


def has_status(
    log_path: Path,
    test_id: str,
    recipient_email: str,
    statuses: set[str],
) -> bool:
    if not log_path.exists():
        return False

    with log_path.open("r", encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            if (
                row.get("test_id") == test_id
                and row.get("recipient_email", "").lower() == recipient_email.lower()
                and row.get("status") in statuses
            ):
                return True
    return False


def append_send_log(
    log_path: Path,
    email: QueueEmail,
    status: str,
    graph_response: str = "",
    error: str = "",
) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    ensure_send_log_schema(log_path)
    file_exists = log_path.exists()
    with log_path.open("a", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=HEADERS)
        if not file_exists:
            writer.writeheader()
        writer.writerow(
            {
                "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
                "test_id": email.test_id,
                "target_name": email.target_name,
                "recipient_email": email.recipient_email,
                "recipient_type": email.recipient_type,
                "template_type": email.template_type,
                "language": email.language,
                "subject": email.subject,
                "scheduled_send_at": clean_scheduled_send_at(email.scheduled_send_at),
                "status": status,
                "graph_response": graph_response,
                "error": error,
            }
        )


def ensure_send_log_schema(log_path: Path) -> bool:
    if not log_path.exists():
        return False

    with log_path.open("r", encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        fieldnames = reader.fieldnames or []
        rows = [_normalise_row(row) for row in reader]

    if fieldnames == HEADERS:
        return False

    temp_path = log_path.with_suffix(f"{log_path.suffix}.tmp")
    with temp_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=HEADERS)
        writer.writeheader()
        writer.writerows(rows)
    temp_path.replace(log_path)
    return True


def _normalise_row(row: dict[str, str]) -> dict[str, str]:
    normalised = {header: row.get(header, "") or "" for header in HEADERS}
    normalised["scheduled_send_at"] = clean_scheduled_send_at(normalised["scheduled_send_at"])
    if not normalised["scheduled_send_at"]:
        normalised["scheduled_send_at"] = scheduled_send_at_from_response(
            normalised["graph_response"]
        )
    return normalised
