from __future__ import annotations

import unittest
from unittest.mock import patch

from outreach.service import OutreachService, PipelineState
from outreach.status import FileLineageStatus, TargetStatus


class OutreachServiceTests(unittest.TestCase):
    def test_delivery_preflight_allows_eligible_recipient_without_allowlist(self) -> None:
        target = TargetStatus(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="recipient@example.com",
            enabled="Yes",
            handoff_generated="Yes",
            review_imported="Yes",
            ready_to_send="Yes",
            validation_status="VALID",
            scheduled_send_at="",
            last_send_status="",
            last_send_time="",
            template_type="General / Fallback",
            language="English",
        )
        service = OutreachService()

        with (
            patch.object(
                service,
                "load_pipeline_state",
                return_value=PipelineState(
                    lineage=FileLineageStatus("OK", "ok"),
                    targets=(target,),
                ),
            ),
            patch("config.SURVEY_URL", "https://example.com/survey"),
        ):
            result = service.delivery_preflight("draft")

        self.assertEqual((target,), result.eligible)
        self.assertEqual((), result.blocked)

    def test_delivery_preflight_blocks_refused_to_answer(self) -> None:
        target = TargetStatus(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="recipient@example.com",
            enabled="Yes",
            handoff_generated="Yes",
            review_imported="Yes",
            ready_to_send="Yes",
            validation_status="VALID",
            scheduled_send_at="",
            last_send_status="REFUSED_TO_ANSWER",
            last_send_time="",
            template_type="General / Fallback",
            language="English",
        )
        service = OutreachService()

        with (
            patch.object(
                service,
                "load_pipeline_state",
                return_value=PipelineState(
                    lineage=FileLineageStatus("OK", "ok"),
                    targets=(target,),
                ),
            ),
            patch("config.SURVEY_URL", "https://example.com/survey"),
        ):
            result = service.delivery_preflight("draft")

        self.assertEqual((), result.eligible)
        self.assertEqual(((target, "recipient refused to answer; manual review required"),), result.blocked)

    def test_delivery_preflight_blocks_stale_prompt_version(self) -> None:
        target = TargetStatus(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="recipient@example.com",
            enabled="Yes",
            handoff_generated="Yes",
            review_imported="Yes",
            ready_to_send="Yes",
            validation_status="VALID",
            scheduled_send_at="",
            last_send_status="",
            last_send_time="",
            template_type="General / Fallback",
            language="English",
            prompt_version="old-version",
        )
        service = OutreachService()

        with (
            patch.object(
                service,
                "load_pipeline_state",
                return_value=PipelineState(
                    lineage=FileLineageStatus("OK", "ok"),
                    targets=(target,),
                ),
            ),
            patch("config.SURVEY_URL", "https://example.com/survey"),
            patch("config.PROMPT_PACKET_VERSION", "current-version"),
        ):
            result = service.delivery_preflight("draft")

        self.assertEqual((), result.eligible)
        self.assertEqual(
            (
                (
                    target,
                    "review prompt version is stale (old-version); "
                    "regenerate handoff and reimport with current-version",
                ),
            ),
            result.blocked,
        )


if __name__ == "__main__":
    unittest.main()
