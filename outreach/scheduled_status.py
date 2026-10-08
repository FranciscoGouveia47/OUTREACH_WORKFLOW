from __future__ import annotations

from dataclasses import dataclass
from dataclasses import replace
from datetime import datetime

import config
from outreach.logger import (
    OUTLOOK_OUTBOX_LOCKED_STATUS,
    OUTLOOK_SCHEDULED_STATUS,
    SENT_STATUS,
    append_send_log,
    latest_log_row,
)
from outreach.models import QueueEmail
from outreach.outlook_client import run_outlook_find_sent_emails
from outreach.review_queue import parse_scheduled_send_at, read_review_email
from outreach.status import load_target_statuses


SENT_LOOKUP_BATCH_SIZE = 50


@dataclass(frozen=True)
class ScheduledStatusResult:
    checked: int
    marked_sent: int
    pending: int
    errors: tuple[str, ...]
    outbox_pending: int = 0
    not_yet_due: int = 0


def reconcile_scheduled_sends() -> ScheduledStatusResult:
    checked = 0
    marked_sent = 0
    pending = 0
    outbox_pending = 0
    not_yet_due = 0
    errors: list[str] = []
    due_emails: list[QueueEmail] = []

    for status in load_target_statuses():
        if status.last_send_status not in {OUTLOOK_SCHEDULED_STATUS, OUTLOOK_OUTBOX_LOCKED_STATUS}:
            continue
        if not status.scheduled_send_at.strip():
            continue
        if parse_scheduled_send_at(status.scheduled_send_at) > datetime.now():
            not_yet_due += 1
            pending += 1
            continue

        checked += 1
        scheduled_log = latest_log_row(
            config.SEND_LOG_PATH,
            status.test_id,
            status.recipient_email,
            {OUTLOOK_OUTBOX_LOCKED_STATUS, OUTLOOK_SCHEDULED_STATUS},
        )
        try:
            review_email = read_review_email(config.LOCAL_REVIEW_WORKBOOK_PATH, status.test_id)
            email = replace(
                review_email,
                subject=scheduled_log.get("subject", "") or review_email.subject,
                scheduled_send_at=status.scheduled_send_at,
            )
        except Exception as exc:
            if not scheduled_log:
                errors.append(f"{status.test_id}: {exc}")
                continue
            email = _email_from_scheduled_log(status, scheduled_log)
        due_emails.append(email)

    sent_matches: dict[tuple[str, str], str] = {}
    failed_lookup_keys: set[tuple[str, str]] = set()
    for batch in _email_batches(_sort_for_sent_lookup(due_emails), SENT_LOOKUP_BATCH_SIZE):
        try:
            sent_matches.update(run_outlook_find_sent_emails(batch, config.OUTLOOK_TIMEOUT_SECONDS))
        except Exception as exc:
            errors.append(f"Outlook sent lookup failed for {len(batch)} due scheduled row(s): {exc}")
            failed_lookup_keys.update(_email_key(email) for email in batch)
            pending += len(batch)

    for email in due_emails:
        key = _email_key(email)
        if key in failed_lookup_keys:
            continue
        response = sent_matches.get(key, "")
        if response:
            append_send_log(
                config.SEND_LOG_PATH,
                email,
                SENT_STATUS,
                graph_response=response,
            )
            marked_sent += 1
        else:
            pending += 1

    return ScheduledStatusResult(
        checked=checked,
        marked_sent=marked_sent,
        pending=pending,
        errors=tuple(errors),
        outbox_pending=outbox_pending,
        not_yet_due=not_yet_due,
    )


def _email_from_scheduled_log(status, scheduled_log: dict[str, str]) -> QueueEmail:
    return QueueEmail(
        test_id=status.test_id,
        target_name=scheduled_log.get("target_name", "") or status.target_name,
        recipient_email=scheduled_log.get("recipient_email", "") or status.recipient_email,
        recipient_type=scheduled_log.get("recipient_type", ""),
        template_type=scheduled_log.get("template_type", "") or status.template_type,
        language=scheduled_log.get("language", "") or status.language,
        subject=scheduled_log.get("subject", ""),
        body="",
        ready_to_send=status.ready_to_send == "Yes",
        validation_status=status.validation_status,
        scheduled_send_at=status.scheduled_send_at,
    )


def _sort_for_sent_lookup(emails: list[QueueEmail]) -> list[QueueEmail]:
    return sorted(emails, key=lambda email: parse_scheduled_send_at(email.scheduled_send_at), reverse=True)


def _email_batches(emails: list[QueueEmail], batch_size: int):
    for index in range(0, len(emails), batch_size):
        yield emails[index : index + batch_size]


def _email_key(email: QueueEmail) -> tuple[str, str]:
    return (email.test_id, email.recipient_email.lower())
