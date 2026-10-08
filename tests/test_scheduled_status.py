from __future__ import annotations

import unittest
from datetime import datetime, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from outreach.logger import OUTLOOK_SCHEDULED_STATUS, append_send_log
from outreach.models import QueueEmail
import outreach.scheduled_status as scheduled_status
from outreach.scheduled_status import reconcile_scheduled_sends
from outreach.status import TargetStatus


class ScheduledStatusTests(unittest.TestCase):
    def test_reconcile_scheduled_sends_marks_found_sent_item_as_sent(self) -> None:
        status = TargetStatus(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="test@example.com",
            enabled="Yes",
            handoff_generated="Yes",
            review_imported="Yes",
            ready_to_send="Yes",
            validation_status="VALID",
            scheduled_send_at="2000-01-01 09:00",
            last_send_status="OUTLOOK_SCHEDULED",
            last_send_time="2000-01-01T08:55:00+00:00",
            template_type="General / Fallback",
            language="English",
        )
        email = QueueEmail(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="test@example.com",
            recipient_type="Individual",
            template_type="General / Fallback",
            language="English",
            subject="Subject",
            body="Body",
            ready_to_send=True,
            validation_status="VALID",
            scheduled_send_at="2000-01-01 09:00",
        )

        with (
            patch("outreach.scheduled_status.load_target_statuses", return_value=[status]),
            patch("outreach.scheduled_status.read_review_email", return_value=email),
            patch(
                "outreach.scheduled_status.run_outlook_find_sent_emails",
                return_value={("TEST-001", "test@example.com"): "Found in Sent Items"},
            ),
            patch("outreach.scheduled_status.append_send_log") as append_send_log,
        ):
            result = reconcile_scheduled_sends()

        self.assertEqual(1, result.checked)
        self.assertEqual(1, result.marked_sent)
        self.assertEqual(0, result.pending)
        self.assertEqual(0, result.outbox_pending)
        self.assertFalse(hasattr(scheduled_status, "run_outlook_find_outbox_emails"))
        append_send_log.assert_called_once()
        self.assertEqual("SENT", append_send_log.call_args.args[2])

    def test_reconcile_scheduled_sends_uses_status_schedule_when_review_is_blank(self) -> None:
        status = TargetStatus(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="test@example.com",
            enabled="Yes",
            handoff_generated="Yes",
            review_imported="Yes",
            ready_to_send="Yes",
            validation_status="VALID",
            scheduled_send_at="2000-01-01 09:00",
            last_send_status="OUTLOOK_SCHEDULED",
            last_send_time="2000-01-01T08:55:00+00:00",
            template_type="General / Fallback",
            language="English",
        )
        email = QueueEmail(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="test@example.com",
            recipient_type="Individual",
            template_type="General / Fallback",
            language="English",
            subject="Subject",
            body="Body",
            ready_to_send=True,
            validation_status="VALID",
            scheduled_send_at="",
        )

        with (
            patch("outreach.scheduled_status.load_target_statuses", return_value=[status]),
            patch("outreach.scheduled_status.read_review_email", return_value=email),
            patch(
                "outreach.scheduled_status.run_outlook_find_sent_emails",
                return_value={("TEST-001", "test@example.com"): "Found in Sent Items"},
            ) as outlook_action,
            patch("outreach.scheduled_status.append_send_log"),
        ):
            result = reconcile_scheduled_sends()

        self.assertEqual(1, result.checked)
        queued_email = outlook_action.call_args.args[0][0]
        self.assertEqual("2000-01-01 09:00", queued_email.scheduled_send_at)

    def test_reconcile_scheduled_sends_uses_logged_subject_when_review_changed(self) -> None:
        status = TargetStatus(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="test@example.com",
            enabled="Yes",
            handoff_generated="Yes",
            review_imported="Yes",
            ready_to_send="Yes",
            validation_status="VALID",
            scheduled_send_at="2000-01-01 09:00",
            last_send_status="OUTLOOK_SCHEDULED",
            last_send_time="2000-01-01T08:55:00+00:00",
            template_type="General / Fallback",
            language="English",
        )
        review_email = QueueEmail(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="test@example.com",
            recipient_type="Individual",
            template_type="General / Fallback",
            language="English",
            subject="Current review subject",
            body="Body",
            ready_to_send=True,
            validation_status="VALID",
            scheduled_send_at="",
        )
        scheduled_email = QueueEmail(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="test@example.com",
            recipient_type="Individual",
            template_type="General / Fallback",
            language="English",
            subject="Subject handed to Outlook",
            body="Body",
            ready_to_send=True,
            validation_status="VALID",
            scheduled_send_at="2000-01-01 09:00",
        )

        with TemporaryDirectory() as temp_dir:
            log_path = Path(temp_dir) / "send_log.csv"
            append_send_log(log_path, scheduled_email, OUTLOOK_SCHEDULED_STATUS)

            with (
                patch("config.SEND_LOG_PATH", log_path),
                patch("outreach.scheduled_status.load_target_statuses", return_value=[status]),
                patch("outreach.scheduled_status.read_review_email", return_value=review_email),
                patch("outreach.scheduled_status.run_outlook_find_sent_emails", return_value={}) as outlook_action,
                patch("outreach.scheduled_status.append_send_log"),
            ):
                result = reconcile_scheduled_sends()

        self.assertEqual(1, result.checked)
        queued_email = outlook_action.call_args.args[0][0]
        self.assertEqual("Subject handed to Outlook", queued_email.subject)

    def test_reconcile_scheduled_sends_uses_log_when_review_row_is_missing(self) -> None:
        status = TargetStatus(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="test@example.com",
            enabled="Yes",
            handoff_generated="Yes",
            review_imported="No",
            ready_to_send="Yes",
            validation_status="VALID",
            scheduled_send_at="2000-01-01 09:00",
            last_send_status="OUTLOOK_SCHEDULED",
            last_send_time="2000-01-01T08:55:00+00:00",
            template_type="General / Fallback",
            language="English",
        )
        scheduled_email = QueueEmail(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="test@example.com",
            recipient_type="Individual",
            template_type="General / Fallback",
            language="English",
            subject="Subject handed to Outlook",
            body="Body",
            ready_to_send=True,
            validation_status="VALID",
            scheduled_send_at="2000-01-01 09:00",
        )

        with TemporaryDirectory() as temp_dir:
            log_path = Path(temp_dir) / "send_log.csv"
            append_send_log(log_path, scheduled_email, OUTLOOK_SCHEDULED_STATUS)

            with (
                patch("config.SEND_LOG_PATH", log_path),
                patch("outreach.scheduled_status.load_target_statuses", return_value=[status]),
                patch("outreach.scheduled_status.read_review_email", side_effect=ValueError("missing row")),
                patch(
                    "outreach.scheduled_status.run_outlook_find_sent_emails",
                    return_value={("TEST-001", "test@example.com"): "Found in Sent Items"},
                ) as outlook_action,
                patch("outreach.scheduled_status.append_send_log") as append_send_log_mock,
            ):
                result = reconcile_scheduled_sends()

        self.assertEqual(1, result.checked)
        self.assertEqual(1, result.marked_sent)
        self.assertEqual((), result.errors)
        self.assertEqual(0, result.pending)
        self.assertEqual(0, result.outbox_pending)
        queued_email = outlook_action.call_args.args[0][0]
        self.assertEqual("Subject handed to Outlook", queued_email.subject)
        append_send_log_mock.assert_called_once()
        self.assertEqual("SENT", append_send_log_mock.call_args.args[2])

    def test_reconcile_scheduled_sends_does_not_inspect_outbox_when_sent_item_is_missing(self) -> None:
        status = TargetStatus(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="test@example.com",
            enabled="Yes",
            handoff_generated="Yes",
            review_imported="Yes",
            ready_to_send="Yes",
            validation_status="VALID",
            scheduled_send_at="2000-01-01 09:00",
            last_send_status="OUTLOOK_SCHEDULED",
            last_send_time="2000-01-01T08:55:00+00:00",
            template_type="General / Fallback",
            language="English",
        )
        email = QueueEmail(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="test@example.com",
            recipient_type="Individual",
            template_type="General / Fallback",
            language="English",
            subject="Subject",
            body="Body",
            ready_to_send=True,
            validation_status="VALID",
            scheduled_send_at="2000-01-01 09:00",
        )

        with (
            patch("outreach.scheduled_status.load_target_statuses", return_value=[status]),
            patch("outreach.scheduled_status.read_review_email", return_value=email),
            patch("outreach.scheduled_status.run_outlook_find_sent_emails", return_value={}),
            patch("outreach.scheduled_status.append_send_log") as append_send_log_mock,
        ):
            result = reconcile_scheduled_sends()

        self.assertEqual(1, result.checked)
        self.assertEqual(0, result.marked_sent)
        self.assertEqual(1, result.pending)
        self.assertEqual(0, result.outbox_pending)
        self.assertFalse(hasattr(scheduled_status, "run_outlook_find_outbox_emails"))
        append_send_log_mock.assert_not_called()

    def test_reconcile_scheduled_sends_batches_due_outlook_lookup(self) -> None:
        statuses = [
            TargetStatus(
                test_id=f"TEST-00{index}",
                target_name="Target",
                recipient_email=f"test{index}@example.com",
                enabled="Yes",
                handoff_generated="Yes",
                review_imported="Yes",
                ready_to_send="Yes",
                validation_status="VALID",
                scheduled_send_at="2000-01-01 09:00",
                last_send_status="OUTLOOK_SCHEDULED",
                last_send_time="2000-01-01T08:55:00+00:00",
                template_type="General / Fallback",
                language="English",
            )
            for index in (1, 2)
        ]
        emails = [
            QueueEmail(
                test_id=f"TEST-00{index}",
                target_name="Target",
                recipient_email=f"test{index}@example.com",
                recipient_type="Individual",
                template_type="General / Fallback",
                language="English",
                subject=f"Subject {index}",
                body="Body",
                ready_to_send=True,
                validation_status="VALID",
                scheduled_send_at="",
            )
            for index in (1, 2)
        ]

        with (
            patch("outreach.scheduled_status.load_target_statuses", return_value=statuses),
            patch("outreach.scheduled_status.read_review_email", side_effect=emails),
            patch(
                "outreach.scheduled_status.run_outlook_find_sent_emails",
                return_value={
                    ("TEST-001", "test1@example.com"): "Found in Sent Items",
                    ("TEST-002", "test2@example.com"): "Found in Sent Items",
                },
            ) as outlook_lookup,
            patch("outreach.scheduled_status.append_send_log") as append_send_log,
        ):
            result = reconcile_scheduled_sends()

        self.assertEqual(2, result.checked)
        self.assertEqual(2, result.marked_sent)
        outlook_lookup.assert_called_once()
        self.assertEqual(2, len(outlook_lookup.call_args.args[0]))
        self.assertEqual(2, append_send_log.call_count)

    def test_reconcile_scheduled_sends_splits_large_due_lookup_into_batches(self) -> None:
        statuses = [
            TargetStatus(
                test_id=f"TEST-00{index}",
                target_name="Target",
                recipient_email=f"test{index}@example.com",
                enabled="Yes",
                handoff_generated="Yes",
                review_imported="Yes",
                ready_to_send="Yes",
                validation_status="VALID",
                scheduled_send_at="2000-01-01 09:00",
                last_send_status="OUTLOOK_SCHEDULED",
                last_send_time="2000-01-01T08:55:00+00:00",
                template_type="General / Fallback",
                language="English",
            )
            for index in (1, 2, 3)
        ]
        emails = [
            QueueEmail(
                test_id=f"TEST-00{index}",
                target_name="Target",
                recipient_email=f"test{index}@example.com",
                recipient_type="Individual",
                template_type="General / Fallback",
                language="English",
                subject=f"Subject {index}",
                body="Body",
                ready_to_send=True,
                validation_status="VALID",
                scheduled_send_at="",
            )
            for index in (1, 2, 3)
        ]

        def fake_lookup(batch, _timeout_seconds):
            return {
                (email.test_id, email.recipient_email.lower()): "Found in Sent Items"
                for email in batch
            }

        with (
            patch.object(scheduled_status, "SENT_LOOKUP_BATCH_SIZE", 2),
            patch("outreach.scheduled_status.load_target_statuses", return_value=statuses),
            patch("outreach.scheduled_status.read_review_email", side_effect=emails),
            patch("outreach.scheduled_status.run_outlook_find_sent_emails", side_effect=fake_lookup) as outlook_lookup,
            patch("outreach.scheduled_status.append_send_log") as append_send_log,
        ):
            result = reconcile_scheduled_sends()

        self.assertEqual(3, result.checked)
        self.assertEqual(3, result.marked_sent)
        self.assertEqual(0, result.pending)
        self.assertEqual([2, 1], [len(call.args[0]) for call in outlook_lookup.call_args_list])
        self.assertEqual(3, append_send_log.call_count)

    def test_reconcile_scheduled_sends_continues_after_lookup_batch_failure(self) -> None:
        statuses = [
            TargetStatus(
                test_id=f"TEST-00{index}",
                target_name="Target",
                recipient_email=f"test{index}@example.com",
                enabled="Yes",
                handoff_generated="Yes",
                review_imported="Yes",
                ready_to_send="Yes",
                validation_status="VALID",
                scheduled_send_at="2000-01-01 09:00",
                last_send_status="OUTLOOK_SCHEDULED",
                last_send_time="2000-01-01T08:55:00+00:00",
                template_type="General / Fallback",
                language="English",
            )
            for index in (1, 2, 3)
        ]
        emails = [
            QueueEmail(
                test_id=f"TEST-00{index}",
                target_name="Target",
                recipient_email=f"test{index}@example.com",
                recipient_type="Individual",
                template_type="General / Fallback",
                language="English",
                subject=f"Subject {index}",
                body="Body",
                ready_to_send=True,
                validation_status="VALID",
                scheduled_send_at="",
            )
            for index in (1, 2, 3)
        ]

        with (
            patch.object(scheduled_status, "SENT_LOOKUP_BATCH_SIZE", 2),
            patch("outreach.scheduled_status.load_target_statuses", return_value=statuses),
            patch("outreach.scheduled_status.read_review_email", side_effect=emails),
            patch(
                "outreach.scheduled_status.run_outlook_find_sent_emails",
                side_effect=[
                    TimeoutError("slow lookup"),
                    {("TEST-003", "test3@example.com"): "Found in Sent Items"},
                ],
            ),
            patch("outreach.scheduled_status.append_send_log") as append_send_log,
        ):
            result = reconcile_scheduled_sends()

        self.assertEqual(3, result.checked)
        self.assertEqual(1, result.marked_sent)
        self.assertEqual(2, result.pending)
        self.assertEqual(1, len(result.errors))
        self.assertIn("2 due scheduled row(s)", result.errors[0])
        append_send_log.assert_called_once()

    def test_reconcile_scheduled_sends_skips_outlook_when_nothing_is_due(self) -> None:
        future_schedule = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d %H:%M")
        status = TargetStatus(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="test@example.com",
            enabled="Yes",
            handoff_generated="Yes",
            review_imported="Yes",
            ready_to_send="Yes",
            validation_status="VALID",
            scheduled_send_at=future_schedule,
            last_send_status="OUTLOOK_SCHEDULED",
            last_send_time="2000-01-01T08:55:00+00:00",
            template_type="General / Fallback",
            language="English",
        )

        with (
            patch("outreach.scheduled_status.load_target_statuses", return_value=[status]),
            patch("outreach.scheduled_status.run_outlook_find_sent_emails") as outlook_lookup,
        ):
            result = reconcile_scheduled_sends()

        self.assertEqual(0, result.checked)
        self.assertEqual(1, result.pending)
        self.assertEqual(1, result.not_yet_due)
        outlook_lookup.assert_not_called()
        self.assertFalse(hasattr(scheduled_status, "run_outlook_find_outbox_emails"))


if __name__ == "__main__":
    unittest.main()
