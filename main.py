from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import config
from outreach.console import detail, result, run_header, stage
from outreach.excel_reader import read_workbook
from outreach.handoff_workbook import read_handoff_test_ids, write_handoff_workbook
from outreach.logger import already_sent, has_delivery_failed, has_outlook_scheduled, has_refused_to_answer
from outreach.models import HandoffEmail, Stakeholder
from outreach.preview_writer import clear_previews, save_preview
from outreach.prompt_packet import build_prompt_packet
from outreach.review_importer import import_completed_handoff
from outreach.review_queue import read_ready_queue
from outreach.scheduler import DELIVERY_GRAPH, DELIVERY_OUTLOOK, DELIVERY_OUTLOOK_DRAFTS, process_send_queue
from outreach.template_engine import render_email
from outreach.validator import validate_stakeholder, validate_templates, validate_unique_ids
from outreach.workbook_state import assert_workbooks_writable


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Configurable stakeholder outreach handoff generator."
    )
    parser.add_argument(
        "command",
        nargs="?",
        default="ui",
        choices=(
            "ui",
            "preview",
            "validate",
            "generate-handoff",
            "import-review",
            "send-ready",
            "create-outlook-drafts",
            "preflight",
            "migrate-files",
        ),
        help="Default is ui. Delivery commands require explicit confirmation.",
    )
    parser.add_argument(
        "--include-generated",
        action="store_true",
        help="For generate-handoff only: include rows already marked as handed off.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="For generate-handoff only: maximum number of handoff rows to generate.",
    )
    parser.add_argument(
        "--completed-handoff",
        type=Path,
        default=config.CHATGPT_HANDOFF_WORKBOOK_PATH,
        help="For import-review only: completed workbook returned by ChatGPT.",
    )
    parser.add_argument(
        "--confirm-send",
        action="store_true",
        help="For send-ready only: actually send through the selected delivery backend.",
    )
    parser.add_argument(
        "--delivery",
        choices=(DELIVERY_GRAPH, DELIVERY_OUTLOOK),
        default=DELIVERY_GRAPH,
        help="For send-ready only: delivery backend for actual sending.",
    )
    parser.add_argument(
        "--interval-minutes",
        type=int,
        default=config.SEND_INTERVAL_MINUTES,
        help="For send-ready only: delay between sent emails.",
    )
    parser.add_argument(
        "--max-daily",
        type=int,
        default=config.MAX_DAILY_EMAILS,
        help="For send-ready only: maximum emails to send or dry-run in this run.",
    )
    parser.add_argument(
        "--ignore-send-window",
        action="store_true",
        help="For send-ready only: ignore the configured working-hours send window.",
    )
    parser.add_argument(
        "--outlook-timeout-seconds",
        type=int,
        default=config.OUTLOOK_TIMEOUT_SECONDS,
        help="For Outlook delivery only: timeout for each Outlook COM operation.",
    )
    parser.add_argument(
        "--preflight-mode",
        choices=("draft", "schedule_outlook", "send_now"),
        default="draft",
        help="For preflight only: delivery action to check.",
    )
    return parser


def run(
    command: str,
    include_generated: bool = False,
    limit: int | None = None,
    completed_handoff: Path = config.CHATGPT_HANDOFF_WORKBOOK_PATH,
    confirm_send: bool = False,
    delivery: str = DELIVERY_GRAPH,
    interval_minutes: int = config.SEND_INTERVAL_MINUTES,
    max_daily: int = config.MAX_DAILY_EMAILS,
    ignore_send_window: bool = False,
    outlook_timeout_seconds: int = config.OUTLOOK_TIMEOUT_SECONDS,
    preflight_mode: str = "draft",
) -> int:
    if command == "ui":
        import ui

        ui.main()
        return 0

    if command == "migrate-files":
        return _migrate_files()

    if command == "preflight":
        return _preflight(preflight_mode)

    run_header(command)
    stage(1, "Load Source Workbook")
    detail(f"Workbook: {config.EXCEL_PATH}")
    subjects, templates = read_workbook(config.EXCEL_PATH)
    enabled_count = sum(subject.enabled for subject in subjects)
    detail(f"Loaded {len(subjects)} subject row(s), {enabled_count} enabled row(s).")
    detail(f"Loaded {len(templates)} email template(s).")

    stage(2, "Validate Configuration And Enabled Rows")
    errors = _validate_configuration_and_rows(subjects, templates)
    if errors:
        detail(f"Found {len(errors)} validation/configuration issue(s).")
    else:
        detail("Validation passed.")

    if command == "validate":
        return _print_validation_result(subjects, errors)

    if command == "generate-handoff":
        return _generate_handoff(subjects, templates, errors, include_generated, limit)

    if command == "import-review":
        return _import_review(subjects, errors, completed_handoff)

    if command == "send-ready":
        return _send_ready(
            confirm_send=confirm_send,
            delivery=delivery,
            requested_delivery=delivery,
            interval_minutes=interval_minutes,
            max_daily=max_daily,
            respect_window=not ignore_send_window,
            outlook_timeout_seconds=outlook_timeout_seconds,
        )

    if command == "create-outlook-drafts":
        return _send_ready(
            confirm_send=True,
            delivery=DELIVERY_OUTLOOK_DRAFTS,
            requested_delivery=DELIVERY_OUTLOOK_DRAFTS,
            interval_minutes=interval_minutes,
            max_daily=max_daily,
            respect_window=False,
            outlook_timeout_seconds=outlook_timeout_seconds,
        )

    return _preview(subjects, templates, errors)


def _migrate_files() -> int:
    from outreach.service import OutreachService
    from outreach.logger import ensure_send_log_schema

    run_header("migrate-files")
    stage(1, "Migrate Runtime Workbook Schemas")
    changed_count = 0
    for migration in OutreachService().migrate_runtime_files():
        added = ", ".join(migration.added_columns) if migration.added_columns else "none"
        detail(f"{migration.path}: added column(s): {added}")
        changed_count += 1 if migration.changed else 0
    stage(2, "Migrate Send Log Schema")
    log_changed = ensure_send_log_schema(config.SEND_LOG_PATH)
    detail(f"{config.SEND_LOG_PATH}: {'updated' if log_changed else 'already current or not present'}")
    changed_count += 1 if log_changed else 0
    result(f"Migration complete. Updated {changed_count} runtime file(s).")
    return 0


def _preflight(preflight_mode: str) -> int:
    from outreach.service import OutreachService

    run_header("preflight")
    stage(1, "Load Pipeline State")
    service = OutreachService()
    preflight = service.delivery_preflight(preflight_mode)
    stage(2, "Global Checks")
    if preflight.global_issues:
        for issue in preflight.global_issues:
            detail(f"Blocked: {issue}")
    else:
        detail("No global delivery blockers found.")

    stage(3, "Target Checks")
    detail(f"Eligible target(s): {len(preflight.eligible)}")
    detail(f"Blocked target(s): {len(preflight.blocked)}")
    max_blocked_details = 80
    for target, issue in preflight.blocked[:max_blocked_details]:
        detail(f"{target.test_id}: {issue}")
    if len(preflight.blocked) > max_blocked_details:
        detail(f"...and {len(preflight.blocked) - max_blocked_details} more blocked target(s).")

    if preflight.global_issues or preflight.blocked:
        result("Preflight completed with blockers.")
        return 1
    result("Preflight passed.")
    return 0


def _validate_configuration_and_rows(subjects, templates) -> list[str]:
    errors: list[str] = []
    errors.extend(validate_templates(templates))
    errors.extend(validate_unique_ids(subjects))

    for subject in subjects:
        if not subject.enabled:
            continue
        row_errors = validate_stakeholder(subject)
        errors.extend(f"{subject.test_id}: {error}" for error in row_errors)

    return errors


def _print_validation_result(subjects: list[Stakeholder], errors: list[str]) -> int:
    stage(3, "Validation Result")
    if errors:
        print("Validation failed:")
        for error in errors:
            print(f"- {error}")
        return 1
    result(f"Validation passed for {sum(subject.enabled for subject in subjects)} enabled rows.")
    return 0


def _preview(subjects, templates, errors: list[str]) -> int:
    stage(3, "Render Base Email Previews")
    generated = []
    for subject in subjects:
        if not subject.enabled:
            continue
        try:
            generated.append(render_email(subject, templates))
        except ValueError as exc:
            errors.append(f"{subject.test_id}: {exc}")
    detail(f"Rendered {len(generated)} preview email(s).")

    stage(4, "Write Preview Text Files")
    clear_previews(config.PREVIEWS_DIR)
    for email in generated:
        save_preview(email, config.PREVIEWS_DIR)
        detail(f"Saved preview for {email.stakeholder_id}: {config.PREVIEWS_DIR / f'{email.stakeholder_id}.txt'}")

    result(
        f"Dry run complete: generated {len(generated)} preview(s), "
        f"{len(errors)} error(s)."
    )
    if errors:
        print("Errors:")
        for error in errors:
            print(f"- {error}")
    return 1 if errors else 0


def _generate_handoff(
    subjects: list[Stakeholder],
    templates,
    errors: list[str],
    include_generated: bool,
    limit: int | None,
) -> int:
    stage(3, "Prepare ChatGPT Handoff")
    if errors:
        print("Validation failed. No ChatGPT handoff workbook was generated.")
        for error in errors:
            print(f"- {error}")
        return 1

    try:
        detail("Checking source and handoff workbook write access.")
        assert_workbooks_writable(config.CHATGPT_HANDOFF_WORKBOOK_PATH)
    except PermissionError as exc:
        print(str(exc))
        print("Close the workbook in Excel and rerun the command.")
        return 1

    handoff_emails: list[HandoffEmail] = []
    skipped_generated = 0
    skipped_failed = 0
    skipped_refused = 0
    skipped_scheduled = 0
    skipped_sent = 0
    already_handed_off = read_handoff_test_ids(config.CHATGPT_HANDOFF_WORKBOOK_PATH)

    detail("Selecting enabled rows that have not already been handed off.")
    for subject in subjects:
        if limit is not None and len(handoff_emails) >= limit:
            break
        if not subject.enabled:
            continue
        if already_sent(config.SEND_LOG_PATH, subject.test_id, subject.recipient_email):
            skipped_sent += 1
            continue
        if has_outlook_scheduled(config.SEND_LOG_PATH, subject.test_id, subject.recipient_email):
            skipped_scheduled += 1
            continue
        if has_delivery_failed(config.SEND_LOG_PATH, subject.test_id, subject.recipient_email):
            skipped_failed += 1
            continue
        if has_refused_to_answer(config.SEND_LOG_PATH, subject.test_id, subject.recipient_email):
            skipped_refused += 1
            continue
        if subject.test_id in already_handed_off and not include_generated:
            skipped_generated += 1
            continue

        try:
            base_email = render_email(subject, templates)
        except ValueError as exc:
            errors.append(f"{subject.test_id}: {exc}")
            continue

        handoff_emails.append(
            HandoffEmail(
                stakeholder=subject,
                email=base_email,
                prompt_packet=build_prompt_packet(subject, base_email),
            )
        )

    detail(f"Selected {len(handoff_emails)} row(s) for ChatGPT handoff.")
    if not handoff_emails:
        print(
            "No new ChatGPT handoff rows were generated. "
            f"Skipped {skipped_generated} already-handed-off row(s) and "
            f"{skipped_sent} already-sent row(s), "
            f"{skipped_scheduled} Outlook-scheduled row(s), "
            f"{skipped_failed} delivery-failed row(s), "
            f"{skipped_refused} refused-to-answer row(s)."
        )
        return 1 if errors else 0

    generated_at = datetime.now().astimezone()
    stage(4, "Write ChatGPT Handoff Workbook")
    handoff_path = write_handoff_workbook(
        handoff_emails,
        config.CHATGPT_HANDOFF_WORKBOOK_PATH,
        generated_at,
    )

    result(f"Generated {len(handoff_emails)} ChatGPT handoff row(s): {handoff_path}")
    detail("Upload this workbook to ChatGPT Plus, then fill the AI output columns.")
    if skipped_generated:
        print(f"Skipped {skipped_generated} already-handed-off row(s).")
    if skipped_sent:
        print(f"Skipped {skipped_sent} already-sent row(s).")
    if skipped_scheduled:
        print(f"Skipped {skipped_scheduled} Outlook-scheduled row(s).")
    if skipped_failed:
        print(f"Skipped {skipped_failed} delivery-failed row(s).")
    if skipped_refused:
        print(f"Skipped {skipped_refused} refused-to-answer row(s).")
    return 1 if errors else 0


def _import_review(
    subjects: list[Stakeholder],
    errors: list[str],
    completed_handoff: Path,
) -> int:
    stage(3, "Import Completed ChatGPT Handoff")
    detail(f"Completed handoff workbook: {completed_handoff}")
    if errors:
        print("Validation failed. Completed handoff was not imported.")
        for error in errors:
            print(f"- {error}")
        return 1

    try:
        detail("Checking source and local review workbook write access.")
        assert_workbooks_writable(config.LOCAL_REVIEW_WORKBOOK_PATH)
    except PermissionError as exc:
        print(str(exc))
        print("Close the workbook in Excel and rerun the command.")
        return 1

    imported_at = datetime.now().astimezone()
    try:
        stage(4, "Validate And Write Local Review Queue")
        review_path, review_emails = import_completed_handoff(
            completed_handoff,
            subjects,
            config.LOCAL_REVIEW_WORKBOOK_PATH,
            imported_at,
            send_log_path=config.SEND_LOG_PATH,
        )
    except (FileNotFoundError, ValueError) as exc:
        print(str(exc))
        return 1

    valid_subjects = [
        item.stakeholder
        for item in review_emails
        if item.validation_status == "VALID" and item.stakeholder.source_row
    ]

    invalid_count = sum(1 for item in review_emails if item.validation_status != "VALID")
    result(f"Imported {len(review_emails)} completed handoff row(s): {review_path}")
    detail(f"Valid review email(s): {len(valid_subjects)}")
    detail(f"Invalid review email(s): {invalid_count}")
    if invalid_count:
        print("Open the local review workbook and check Validation Errors.")
    return 1 if invalid_count else 0


def _send_ready(
    *,
    confirm_send: bool,
    delivery: str,
    requested_delivery: str,
    interval_minutes: int,
    max_daily: int,
    respect_window: bool,
    outlook_timeout_seconds: int,
) -> int:
    stage(3, "Read Local Review Queue")
    detail(f"Review queue workbook: {config.LOCAL_REVIEW_WORKBOOK_PATH}")
    try:
        ready_emails = read_ready_queue(config.LOCAL_REVIEW_WORKBOOK_PATH)
    except (FileNotFoundError, ValueError) as exc:
        print(str(exc))
        return 1

    if not ready_emails:
        result("No review queue rows are marked Ready To Send = Yes with Validation Status = VALID.")
        return 0

    mode = "SEND" if confirm_send else "DRY RUN"
    if delivery == DELIVERY_OUTLOOK_DRAFTS:
        mode = "OUTLOOK DRAFT"
    detail(f"{mode}: found {len(ready_emails)} ready review queue row(s).")
    if not confirm_send:
        detail("No email will be sent. Add --confirm-send to use a delivery backend.")

    try:
        stage(4, "Process Send Queue")
        detail(f"Delivery backend: {delivery}")
        detail(f"Max emails this run: {max_daily}")
        detail(f"Interval minutes: {interval_minutes}")
        detail(f"Respect send window: {respect_window}")
        sent_or_planned, skipped, blocked = process_send_queue(
            ready_emails,
            confirm_send=confirm_send,
            delivery=delivery,
            interval_minutes=interval_minutes,
            max_daily=max_daily,
            respect_window=respect_window,
            outlook_timeout_seconds=outlook_timeout_seconds,
        )
    except FileNotFoundError as exc:
        print(str(exc))
        return 1
    except RuntimeError as exc:
        print(str(exc))
        return 1

    action = "Sent" if confirm_send else "Dry-run planned"
    if requested_delivery == DELIVERY_OUTLOOK_DRAFTS:
        action = "Outlook drafts created"
    stage(5, "Send Queue Result")
    result(f"{action}: {sent_or_planned}")
    detail(f"Skipped already sent/existing draft: {skipped}")
    detail(f"Blocked/failed: {blocked}")
    detail(f"Send log: {config.SEND_LOG_PATH}")
    return 1 if blocked else 0


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return run(
        args.command,
        include_generated=args.include_generated,
        limit=args.limit,
        completed_handoff=args.completed_handoff,
        confirm_send=args.confirm_send,
        delivery=args.delivery,
        interval_minutes=args.interval_minutes,
        max_daily=args.max_daily,
        ignore_send_window=args.ignore_send_window,
        outlook_timeout_seconds=args.outlook_timeout_seconds,
        preflight_mode=args.preflight_mode,
    )


if __name__ == "__main__":
    sys.exit(main())
