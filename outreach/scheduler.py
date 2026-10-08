from __future__ import annotations

import time
from datetime import datetime, time as clock_time

import config
from outreach.console import detail
from outreach.graph_client import GraphClient
from outreach.logger import (
    DRAFT_CREATED_STATUS,
    DRAFT_FAILED_STATUS,
    OUTLOOK_SCHEDULED_STATUS,
    OUTLOOK_OUTBOX_LOCKED_STATUS,
    REFUSED_TO_ANSWER_STATUS,
    already_sent,
    append_send_log,
    has_draft_created,
    has_delivery_failed,
    has_outlook_scheduled,
    has_outbox_locked,
    has_refused_to_answer,
    has_unknown_send_status,
    latest_draft_entry_id,
)
from outreach.models import QueueEmail
from outreach.outlook_client import run_outlook_action


DELIVERY_GRAPH = "graph"
DELIVERY_OUTLOOK = "outlook"
DELIVERY_OUTLOOK_DRAFTS = "outlook-drafts"
DELIVERY_OUTLOOK_SCHEDULED = "outlook-scheduled"


def process_send_queue(
    ready_emails: list[QueueEmail],
    *,
    confirm_send: bool,
    delivery: str = DELIVERY_GRAPH,
    interval_minutes: int = config.SEND_INTERVAL_MINUTES,
    max_daily: int = config.MAX_DAILY_EMAILS,
    respect_window: bool = True,
    outlook_timeout_seconds: int = config.OUTLOOK_TIMEOUT_SECONDS,
) -> tuple[int, int, int]:
    sender = _build_sender(delivery, confirm_send)
    draft_mode = delivery == DELIVERY_OUTLOOK_DRAFTS
    scheduled_mode = delivery == DELIVERY_OUTLOOK_SCHEDULED

    completed_or_planned = 0
    skipped = 0
    blocked = 0

    for email in ready_emails:
        if completed_or_planned >= max_daily:
            detail(f"Daily/run cap reached at {max_daily}.")
            break

        detail(f"Checking {email.test_id} -> {email.recipient_email}")
        if already_sent(config.SEND_LOG_PATH, email.test_id, email.recipient_email):
            detail("Skipped: already logged as SENT.")
            skipped += 1
            continue
        if has_outlook_scheduled(config.SEND_LOG_PATH, email.test_id, email.recipient_email):
            detail("Skipped: already logged as OUTLOOK_SCHEDULED.")
            skipped += 1
            continue
        if draft_mode and has_draft_created(config.SEND_LOG_PATH, email.test_id, email.recipient_email):
            detail("Skipped: Outlook draft already logged.")
            skipped += 1
            continue
        if has_unknown_send_status(config.SEND_LOG_PATH, email.test_id, email.recipient_email):
            detail("Blocked: previous SEND_STATUS_UNKNOWN requires manual review.")
            append_send_log(
                config.SEND_LOG_PATH,
                email,
                "BLOCKED",
                error="Previous SEND_STATUS_UNKNOWN requires manual review before retry.",
            )
            blocked += 1
            continue
        if has_outbox_locked(config.SEND_LOG_PATH, email.test_id, email.recipient_email):
            detail("Blocked: scheduled Outlook message is stuck in Outbox; use the outbox repair action.")
            append_send_log(
                config.SEND_LOG_PATH,
                email,
                "BLOCKED",
                error=f"Previous {OUTLOOK_OUTBOX_LOCKED_STATUS} requires outbox repair before retry.",
            )
            blocked += 1
            continue
        if has_delivery_failed(config.SEND_LOG_PATH, email.test_id, email.recipient_email):
            detail("Blocked: previous DELIVERY_FAILED requires recipient review.")
            append_send_log(
                config.SEND_LOG_PATH,
                email,
                "BLOCKED",
                error="Previous DELIVERY_FAILED requires recipient review before retry.",
            )
            blocked += 1
            continue
        if has_refused_to_answer(config.SEND_LOG_PATH, email.test_id, email.recipient_email):
            detail("Blocked: previous REFUSED_TO_ANSWER requires manual review.")
            append_send_log(
                config.SEND_LOG_PATH,
                email,
                "BLOCKED",
                error=f"Previous {REFUSED_TO_ANSWER_STATUS} requires manual review before retry.",
            )
            blocked += 1
            continue

        if email.prompt_version and email.prompt_version != config.PROMPT_PACKET_VERSION:
            detail(
                "Blocked: review row was generated with stale prompt version "
                f"{email.prompt_version}; current version is {config.PROMPT_PACKET_VERSION}."
            )
            append_send_log(
                config.SEND_LOG_PATH,
                email,
                "BLOCKED",
                error=(
                    f"Stale prompt version {email.prompt_version}; regenerate handoff and "
                    f"reimport with {config.PROMPT_PACKET_VERSION}."
                ),
            )
            blocked += 1
            continue

        if respect_window and not scheduled_mode and not _inside_send_window():
            detail("Blocked: outside configured send window.")
            append_send_log(config.SEND_LOG_PATH, email, "BLOCKED", error="Outside configured send window.")
            blocked += 1
            continue

        if scheduled_mode and not email.scheduled_send_at.strip():
            detail("Blocked: scheduled Outlook delivery requires Scheduled Send At.")
            append_send_log(
                config.SEND_LOG_PATH,
                email,
                "BLOCKED",
                error="Scheduled Outlook delivery requires Scheduled Send At.",
            )
            blocked += 1
            continue

        if not confirm_send:
            detail("Dry-run logged; no delivery attempted.")
            append_send_log(config.SEND_LOG_PATH, email, "DRY_RUN")
            completed_or_planned += 1
            continue

        try:
            if draft_mode:
                detail("Creating Outlook draft.")
                response = run_outlook_action(email, "draft", outlook_timeout_seconds)
                append_send_log(config.SEND_LOG_PATH, email, DRAFT_CREATED_STATUS, graph_response=response)
                detail("Outlook draft created.")
                completed_or_planned += 1
                continue

            action_label = "Scheduling email in Outlook" if scheduled_mode else "Sending email"
            detail(f"{action_label}.")
            append_send_log(config.SEND_LOG_PATH, email, "SCHEDULING" if scheduled_mode else "SENDING")
            if delivery == DELIVERY_OUTLOOK:
                draft_entry_id = latest_draft_entry_id(
                    config.SEND_LOG_PATH,
                    email.test_id,
                    email.recipient_email,
                )
                if draft_entry_id:
                    detail("Sending existing Outlook draft.")
                    response = run_outlook_action(
                        email,
                        "send-draft",
                        outlook_timeout_seconds,
                        draft_entry_id=draft_entry_id,
                    )
                else:
                    response = run_outlook_action(email, "send", outlook_timeout_seconds)
            elif scheduled_mode:
                draft_entry_id = latest_draft_entry_id(
                    config.SEND_LOG_PATH,
                    email.test_id,
                    email.recipient_email,
                )
                if draft_entry_id:
                    detail("Scheduling existing Outlook draft.")
                    response = run_outlook_action(
                        email,
                        "schedule-draft",
                        outlook_timeout_seconds,
                        draft_entry_id=draft_entry_id,
                    )
                else:
                    response = run_outlook_action(email, "schedule", outlook_timeout_seconds)
            else:
                response = sender.send_email(email)
        except TimeoutError as exc:
            detail(f"Blocked/failed: {exc}")
            status = DRAFT_FAILED_STATUS if draft_mode else "SEND_STATUS_UNKNOWN"
            append_send_log(config.SEND_LOG_PATH, email, status, error=str(exc))
            blocked += 1
            continue
        except Exception as exc:
            detail(f"Failed: {exc}")
            status = DRAFT_FAILED_STATUS if draft_mode else "FAILED"
            append_send_log(config.SEND_LOG_PATH, email, status, error=str(exc))
            blocked += 1
            continue

        final_status = OUTLOOK_SCHEDULED_STATUS if scheduled_mode else "SENT"
        append_send_log(config.SEND_LOG_PATH, email, final_status, graph_response=response)
        detail("Scheduled and logged." if scheduled_mode else "Sent and logged.")
        completed_or_planned += 1

        if interval_minutes > 0 and completed_or_planned < max_daily:
            detail(f"Waiting {interval_minutes} minute(s) before next message.")
            time.sleep(interval_minutes * 60)

    return completed_or_planned, skipped, blocked


def _build_sender(delivery: str, confirm_send: bool):
    if not confirm_send:
        return None
    if delivery == DELIVERY_GRAPH:
        return GraphClient()
    if delivery in {DELIVERY_OUTLOOK, DELIVERY_OUTLOOK_DRAFTS, DELIVERY_OUTLOOK_SCHEDULED}:
        return None
    raise ValueError(f"Unsupported delivery backend: {delivery}")


def _inside_send_window() -> bool:
    now = datetime.now().time()
    start = _parse_time(config.SEND_WINDOW_START)
    end = _parse_time(config.SEND_WINDOW_END)
    return start <= now <= end


def _parse_time(value: str) -> clock_time:
    hour, minute = value.split(":", 1)
    return clock_time(hour=int(hour), minute=int(minute))
