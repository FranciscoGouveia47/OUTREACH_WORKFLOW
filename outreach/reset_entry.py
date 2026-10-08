from __future__ import annotations

from dataclasses import dataclass

import config
from outreach.excel_reader import read_workbook
from outreach.handoff_workbook import remove_handoff_email
from outreach.logger import latest_draft_entry_ids, remove_log_entries
from outreach.models import QueueEmail
from outreach.outlook_client import run_outlook_action
from outreach.review_queue import remove_review_email


@dataclass(frozen=True)
class ResetResult:
    review_removed: bool
    handoff_removed: bool
    log_entries_removed: int
    outlook_drafts_deleted: int
    outlook_errors: tuple[str, ...]


def reset_target_entry(test_id: str) -> ResetResult:
    recipient = _recipient_for_test_id(test_id)
    draft_entry_ids = latest_draft_entry_ids(config.SEND_LOG_PATH, test_id, recipient)

    deleted = 0
    errors: list[str] = []
    placeholder_email = QueueEmail(
        test_id=test_id,
        target_name="",
        recipient_email=recipient,
        recipient_type="",
        template_type="",
        language="",
        subject="",
        body="",
        ready_to_send=False,
        validation_status="",
    )

    for draft_entry_id in draft_entry_ids:
        try:
            run_outlook_action(
                placeholder_email,
                "delete-draft",
                config.OUTLOOK_TIMEOUT_SECONDS,
                draft_entry_id=draft_entry_id,
            )
            deleted += 1
        except Exception as exc:
            errors.append(f"{draft_entry_id}: {exc}")

    review_removed = remove_review_email(config.LOCAL_REVIEW_WORKBOOK_PATH, test_id)
    handoff_removed = remove_handoff_email(config.CHATGPT_HANDOFF_WORKBOOK_PATH, test_id)
    log_entries_removed = remove_log_entries(config.SEND_LOG_PATH, test_id, recipient)

    return ResetResult(
        review_removed=review_removed,
        handoff_removed=handoff_removed,
        log_entries_removed=log_entries_removed,
        outlook_drafts_deleted=deleted,
        outlook_errors=tuple(errors),
    )


def _recipient_for_test_id(test_id: str) -> str:
    subjects, _ = read_workbook(config.EXCEL_PATH)
    for subject in subjects:
        if subject.test_id == test_id:
            return subject.recipient_email
    return ""
