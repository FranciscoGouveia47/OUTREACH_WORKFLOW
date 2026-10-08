from __future__ import annotations

from dataclasses import replace
from datetime import datetime
from pathlib import Path

import config
from outreach.excel_reader import read_workbook
from outreach.handoff_workbook import read_prompt_packet, upsert_handoff_workbook_rows
from outreach.logger import (
    DELIVERY_FAILED_STATUS,
    DELIVERY_FAILURE_CLEARED_STATUS,
    OUTLOOK_OUTBOX_LOCKED_STATUS,
    OUTLOOK_SCHEDULE_CANCELLED_STATUS,
    OUTLOOK_SCHEDULED_STATUS,
    REFUSED_TO_ANSWER_CLEARED_STATUS,
    REFUSED_TO_ANSWER_STATUS,
    SENT_STATUS,
    already_sent,
    append_send_log,
    has_delivery_failed,
    has_outlook_scheduled,
    has_refused_to_answer,
    latest_log_row,
)
from outreach.models import EmailMessage, HandoffEmail, QueueEmail
from outreach.prompt_packet import build_prompt_packet
from outreach.review_queue import is_due, read_review_email
from outreach.review_queue import set_ready_to_send
from outreach.scheduler import (
    DELIVERY_OUTLOOK,
    DELIVERY_OUTLOOK_DRAFTS,
    DELIVERY_OUTLOOK_SCHEDULED,
    process_send_queue,
)
from outreach.status import load_target_statuses
from outreach.template_engine import render_email
from outreach.outlook_client import run_outlook_action


def get_base_email(test_id: str) -> EmailMessage:
    subjects, templates = read_workbook(config.EXCEL_PATH)
    subject = _find_subject(subjects, test_id)
    return render_email(subject, templates)


def get_ai_checked_email(test_id: str) -> QueueEmail:
    return read_review_email(config.LOCAL_REVIEW_WORKBOOK_PATH, test_id)


def get_prompt_packet(test_id: str) -> str:
    return read_prompt_packet(config.CHATGPT_HANDOFF_WORKBOOK_PATH, test_id)


def regenerate_handoff_for_entries(test_ids: list[str]) -> tuple[int, Path]:
    subjects, templates = read_workbook(config.EXCEL_PATH)
    selected_subjects = [_find_subject(subjects, test_id) for test_id in test_ids]
    generated_at = datetime.now().astimezone()
    handoff_emails: list[HandoffEmail] = []
    for subject in selected_subjects:
        if already_sent(config.SEND_LOG_PATH, subject.test_id, subject.recipient_email):
            continue
        if has_outlook_scheduled(config.SEND_LOG_PATH, subject.test_id, subject.recipient_email):
            continue
        if has_delivery_failed(config.SEND_LOG_PATH, subject.test_id, subject.recipient_email):
            continue
        if has_refused_to_answer(config.SEND_LOG_PATH, subject.test_id, subject.recipient_email):
            continue
        base_email = render_email(subject, templates)
        handoff_emails.append(
            HandoffEmail(
                stakeholder=subject,
                email=base_email,
                prompt_packet=build_prompt_packet(subject, base_email),
            )
        )
    path = upsert_handoff_workbook_rows(
        handoff_emails,
        config.CHATGPT_HANDOFF_WORKBOOK_PATH,
        generated_at,
    )
    return len(handoff_emails), path


def create_outlook_draft_for_entry(test_id: str) -> tuple[int, int, int]:
    email = _ready_review_email(test_id, require_due=False)
    return process_send_queue(
        [email],
        confirm_send=True,
        delivery=DELIVERY_OUTLOOK_DRAFTS,
        interval_minutes=0,
        max_daily=1,
        respect_window=False,
    )


def send_outlook_for_entry(test_id: str) -> tuple[int, int, int]:
    email = replace(_ready_review_email(test_id, require_due=False), scheduled_send_at="")
    return process_send_queue(
        [email],
        confirm_send=True,
        delivery=DELIVERY_OUTLOOK,
        interval_minutes=0,
        max_daily=1,
        respect_window=True,
    )


def schedule_outlook_for_entry(test_id: str) -> tuple[int, int, int]:
    email = _ready_review_email(test_id, require_due=False)
    scheduled_send_at = _scheduled_datetime(email)
    if scheduled_send_at <= datetime.now():
        raise ValueError(f"{test_id} is scheduled for {email.scheduled_send_at}, which is not in the future.")
    return process_send_queue(
        [email],
        confirm_send=True,
        delivery=DELIVERY_OUTLOOK_SCHEDULED,
        interval_minutes=0,
        max_daily=1,
        respect_window=False,
    )


def cancel_outlook_scheduled_for_entry(test_id: str) -> tuple[int, int, int]:
    status = next((row for row in load_target_statuses() if row.test_id == test_id), None)
    if status is None:
        raise ValueError(f"Test ID not found: {test_id}")

    if status.last_send_status != OUTLOOK_SCHEDULED_STATUS:
        set_ready_to_send(config.LOCAL_REVIEW_WORKBOOK_PATH, test_id, False)
        return 0, 1, 0

    email = read_review_email(config.LOCAL_REVIEW_WORKBOOK_PATH, test_id)
    response = run_outlook_action(
        email,
        "cancel-scheduled",
        config.OUTLOOK_TIMEOUT_SECONDS,
    )
    if "not found" in response.lower():
        set_ready_to_send(config.LOCAL_REVIEW_WORKBOOK_PATH, test_id, False)
        return 0, 0, 1

    append_send_log(
        config.SEND_LOG_PATH,
        email,
        OUTLOOK_SCHEDULE_CANCELLED_STATUS,
        graph_response=response,
        error="Scheduled Outlook send cancelled from UI.",
    )
    set_ready_to_send(config.LOCAL_REVIEW_WORKBOOK_PATH, test_id, False)
    return 1, 0, 0


def send_outbox_locked_for_entry(test_id: str) -> tuple[int, int, int]:
    status = next((row for row in load_target_statuses() if row.test_id == test_id), None)
    if status is None:
        raise ValueError(f"Test ID not found: {test_id}")
    if status.last_send_status != OUTLOOK_OUTBOX_LOCKED_STATUS:
        return 0, 1, 0

    email = _scheduled_outbox_email_for_status(status)
    append_send_log(
        config.SEND_LOG_PATH,
        email,
        "SENDING",
        error="Repairing OUTLOOK_OUTBOX_LOCKED by sending the existing Outlook Outbox item now.",
    )
    try:
        response = run_outlook_action(
            email,
            "send-outbox-now",
            config.OUTLOOK_TIMEOUT_SECONDS,
        )
    except Exception as exc:
        append_send_log(
            config.SEND_LOG_PATH,
            email,
            OUTLOOK_OUTBOX_LOCKED_STATUS,
            error=f"Outlook outbox repair failed; item may still be stuck in Outbox: {exc}",
        )
        return 0, 0, 1
    if "not found" in response.lower():
        append_send_log(
            config.SEND_LOG_PATH,
            email,
            OUTLOOK_OUTBOX_LOCKED_STATUS,
            graph_response=response,
            error="Outlook outbox repair could not find the scheduled item; refresh or cancel manually.",
        )
        return 0, 0, 1

    append_send_log(
        config.SEND_LOG_PATH,
        email,
        SENT_STATUS,
        graph_response=response,
    )
    return 1, 0, 0


def toggle_delivery_failed_for_entry(test_id: str) -> str:
    return _toggle_manual_status_for_entry(
        test_id,
        active_status=DELIVERY_FAILED_STATUS,
        cleared_status=DELIVERY_FAILURE_CLEARED_STATUS,
        active_error="Manual delivery-failure flag set in UI after non-delivery report.",
        cleared_error="Manual delivery-failure flag cleared in UI.",
    )


def toggle_refused_to_answer_for_entry(test_id: str) -> str:
    return _toggle_manual_status_for_entry(
        test_id,
        active_status=REFUSED_TO_ANSWER_STATUS,
        cleared_status=REFUSED_TO_ANSWER_CLEARED_STATUS,
        active_error="Manual refused-to-answer flag set in UI.",
        cleared_error="Manual refused-to-answer flag cleared in UI.",
    )


def _toggle_manual_status_for_entry(
    test_id: str,
    *,
    active_status: str,
    cleared_status: str,
    active_error: str,
    cleared_error: str,
) -> str:
    status = next((row for row in load_target_statuses() if row.test_id == test_id), None)
    if status is None:
        raise ValueError(f"Test ID not found: {test_id}")

    email = _queue_email_for_status(status)
    if status.last_send_status == active_status:
        append_send_log(
            config.SEND_LOG_PATH,
            email,
            cleared_status,
            error=cleared_error,
        )
        return cleared_status

    append_send_log(
        config.SEND_LOG_PATH,
        email,
        active_status,
        error=active_error,
    )
    return active_status


def _ready_review_email(test_id: str, require_due: bool = True) -> QueueEmail:
    email = get_ai_checked_email(test_id)
    if email.validation_status != "VALID":
        raise ValueError(f"{test_id} is not valid for delivery: {email.validation_status}")
    if not email.ready_to_send:
        raise ValueError(f"{test_id} is not marked Ready To Send = Yes.")
    if require_due and not is_due(email.scheduled_send_at):
        raise ValueError(f"{test_id} is scheduled for {email.scheduled_send_at}, which is not due yet.")
    return email


def _scheduled_datetime(email: QueueEmail) -> datetime:
    from outreach.review_queue import parse_scheduled_send_at

    if not email.scheduled_send_at.strip():
        raise ValueError(f"{email.test_id} has no Scheduled Send At value.")
    return parse_scheduled_send_at(email.scheduled_send_at)


def _queue_email_for_status(status) -> QueueEmail:
    try:
        return read_review_email(config.LOCAL_REVIEW_WORKBOOK_PATH, status.test_id)
    except Exception:
        return QueueEmail(
            test_id=status.test_id,
            target_name=status.target_name,
            recipient_email=status.recipient_email.lower(),
            recipient_type="",
            template_type=status.template_type,
            language=status.language,
            subject="",
            body="",
            ready_to_send=status.ready_to_send == "Yes",
            validation_status=status.validation_status,
            scheduled_send_at=status.scheduled_send_at,
        )


def _scheduled_outbox_email_for_status(status) -> QueueEmail:
    scheduled_log = latest_log_row(
        config.SEND_LOG_PATH,
        status.test_id,
        status.recipient_email,
        {OUTLOOK_OUTBOX_LOCKED_STATUS, OUTLOOK_SCHEDULED_STATUS},
    )
    if not scheduled_log:
        raise ValueError(f"No scheduled Outlook log entry found for {status.test_id}.")
    subject = scheduled_log.get("subject", "")
    if not subject:
        raise ValueError(f"No scheduled Outlook subject found for {status.test_id}.")
    return QueueEmail(
        test_id=status.test_id,
        target_name=scheduled_log.get("target_name", "") or status.target_name,
        recipient_email=scheduled_log.get("recipient_email", "") or status.recipient_email,
        recipient_type=scheduled_log.get("recipient_type", ""),
        template_type=scheduled_log.get("template_type", "") or status.template_type,
        language=scheduled_log.get("language", "") or status.language,
        subject=subject,
        body="",
        ready_to_send=status.ready_to_send == "Yes",
        validation_status=status.validation_status,
        scheduled_send_at=scheduled_log.get("scheduled_send_at", "") or status.scheduled_send_at,
    )


def _find_subject(subjects, test_id: str):
    for subject in subjects:
        if subject.test_id == test_id:
            return subject
    raise ValueError(f"Test ID not found in source workbook: {test_id}")
