from __future__ import annotations

from dataclasses import dataclass

import config
from outreach.handoff_workbook import HEADERS as HANDOFF_HEADERS
from outreach.handoff_workbook import SHEET_NAME as HANDOFF_SHEET_NAME
from outreach.review_importer import LOCAL_REVIEW_HEADERS, LOCAL_REVIEW_SHEET_NAME
from outreach.status import FileLineageStatus, TargetStatus, load_file_lineage_status, load_target_statuses
from outreach.workbook_schema import WorkbookMigrationResult, ensure_sheet_columns


@dataclass(frozen=True)
class PipelineState:
    lineage: FileLineageStatus
    targets: tuple[TargetStatus, ...]


@dataclass(frozen=True)
class DeliveryPreflightResult:
    eligible: tuple[TargetStatus, ...]
    blocked: tuple[tuple[TargetStatus, str], ...]
    global_issues: tuple[str, ...]


class OutreachService:
    def load_pipeline_state(self) -> PipelineState:
        return PipelineState(
            lineage=load_file_lineage_status(),
            targets=tuple(load_target_statuses()),
        )

    def migrate_runtime_files(self) -> tuple[WorkbookMigrationResult, ...]:
        return (
            ensure_sheet_columns(
                config.CHATGPT_HANDOFF_WORKBOOK_PATH,
                HANDOFF_SHEET_NAME,
                HANDOFF_HEADERS,
            ),
            ensure_sheet_columns(
                config.LOCAL_REVIEW_WORKBOOK_PATH,
                LOCAL_REVIEW_SHEET_NAME,
                LOCAL_REVIEW_HEADERS,
            ),
        )

    def delivery_preflight(
        self,
        mode: str = "draft",
        test_ids: set[str] | None = None,
    ) -> DeliveryPreflightResult:
        state = self.load_pipeline_state()
        global_issues = self._global_delivery_issues(state.lineage)

        eligible: list[TargetStatus] = []
        blocked: list[tuple[TargetStatus, str]] = []
        selected = test_ids or set()
        for target in state.targets:
            if selected and target.test_id not in selected:
                continue
            issue = self._delivery_issue(target, mode)
            if issue:
                blocked.append((target, issue))
            else:
                eligible.append(target)

        return DeliveryPreflightResult(
            eligible=tuple(eligible),
            blocked=tuple(blocked),
            global_issues=tuple(global_issues),
        )

    def _global_delivery_issues(self, lineage: FileLineageStatus) -> list[str]:
        issues: list[str] = []
        if config.CTA_URL == "CTA_URL_TO_BE_CONFIRMED":
            issues.append("CTA_URL is still the placeholder value.")
        if lineage.status != "OK":
            issues.append(f"File lineage is {lineage.status}: {lineage.message}")
        return issues

    def _delivery_issue(self, target: TargetStatus, mode: str) -> str:
        if not target.recipient_email.strip():
            return "recipient email is missing"
        if target.enabled != "Yes":
            return "disabled in source workbook"
        if target.last_send_status == "SENT":
            return "already logged as SENT"
        if target.last_send_status == "OUTLOOK_SCHEDULED":
            return "already logged as OUTLOOK_SCHEDULED"
        if target.last_send_status == "SEND_STATUS_UNKNOWN":
            return "previous send status is unknown and needs manual review"
        if target.last_send_status == "OUTLOOK_OUTBOX_LOCKED":
            return "scheduled Outlook message is stuck in Outbox; repair or cancel it"
        if target.last_send_status == "DELIVERY_FAILED":
            return "previous delivery failed; review or correct the recipient email"
        if target.last_send_status == "REFUSED_TO_ANSWER":
            return "recipient refused to answer; manual review required"
        if target.handoff_generated != "Yes":
            return "handoff row has not been generated"
        if target.review_imported != "Yes":
            return "completed handoff has not been imported"
        if target.ready_to_send != "Yes":
            return "not marked Ready To Send"
        if target.validation_status != "VALID":
            return f"validation is {target.validation_status or '<blank>'}"
        if target.prompt_version and target.prompt_version != config.PROMPT_PACKET_VERSION:
            return (
                f"review prompt version is stale ({target.prompt_version}); "
                f"regenerate handoff and reimport with {config.PROMPT_PACKET_VERSION}"
            )
        if mode == "draft" and target.last_send_status == "DRAFT_CREATED":
            return "Outlook draft is already logged"
        if mode == "schedule_outlook":
            if not target.scheduled_send_at.strip():
                return "Scheduled Send At is blank"
            if target.last_send_status != "DRAFT_CREATED":
                return "Outlook draft has not been created"
        return ""
