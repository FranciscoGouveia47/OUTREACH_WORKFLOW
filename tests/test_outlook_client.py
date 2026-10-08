from __future__ import annotations

import unittest
from datetime import datetime, timezone
from unittest.mock import patch

from outreach.models import QueueEmail
from outreach.outlook_client import (
    OUTLOOK_DELETED_ITEMS_FOLDER,
    OUTLOOK_OUTBOX_FOLDER,
    OUTLOOK_SENT_MAIL_FOLDER,
    PR_DEFERRED_DELIVERY_TIME,
    PR_DEFERRED_SEND_NUMBER,
    PR_DEFERRED_SEND_TIME,
    PR_DEFERRED_SEND_UNITS,
    PR_SMTP_ADDRESS,
    OutlookClient,
)


class FakePropertyAccessor:
    def __init__(self) -> None:
        self.deleted: list[str] = []
        self.properties: dict[str, object] = {}

    def DeleteProperty(self, property_name: str) -> None:
        self.deleted.append(property_name)

    def SetProperty(self, property_name: str, value) -> None:
        self.properties[property_name] = value


class FakeMessage:
    def __init__(self) -> None:
        self.PropertyAccessor = FakePropertyAccessor()
        self.DeferredDeliveryTime = None
        self.saved = False

    def Save(self) -> None:
        self.saved = True


class FakeOutboxMessage(FakeMessage):
    def __init__(self, subject: str, to: str) -> None:
        super().__init__()
        self.Subject = subject
        self.To = to
        self.Recipients = []
        self.sent = False

    def Send(self) -> None:
        self.sent = True


class ExplodingText:
    def __str__(self) -> str:
        raise AssertionError("Outbox metadata should not be read during refresh lookup")


class FakeDeletedMessage:
    def __init__(self) -> None:
        self.deleted = False

    def Delete(self) -> None:
        self.deleted = True


class FakeDraftMessage:
    def __init__(self, moved_message: FakeDeletedMessage) -> None:
        self.moved_message = moved_message
        self.move_target = None

    def Move(self, folder):
        self.move_target = folder
        return self.moved_message


class FakeSession:
    def __init__(self, draft_message: FakeDraftMessage) -> None:
        self.draft_message = draft_message
        self.requested_folder = None

    def GetItemFromID(self, _entry_id: str):
        return self.draft_message

    def GetDefaultFolder(self, folder_id: int):
        self.requested_folder = folder_id
        return "deleted-items-folder"


class FakeOutlook:
    def __init__(self, session: FakeSession) -> None:
        self.Session = session


class FakeItems(list):
    def __init__(self, items) -> None:
        super().__init__(items)
        self.sorted = False
        self.restrict_filter = ""

    def Sort(self, _field: str, _descending: bool) -> None:
        self.sorted = True

    def Restrict(self, filter_value: str):
        self.restrict_filter = filter_value
        return self

    @property
    def Count(self) -> int:
        return len(self)

    def Item(self, index: int):
        return self[index - 1]


class FakeSentItem:
    def __init__(self, subject: str, to: str, sent_on: datetime, recipients=None) -> None:
        self.Subject = subject
        self.To = to
        self.SentOn = sent_on
        self.Recipients = recipients or []


class FakeExchangeUser:
    def __init__(self, smtp_address: str) -> None:
        self.PrimarySmtpAddress = smtp_address


class FakeAddressEntry:
    def __init__(self, smtp_address: str) -> None:
        self.Address = "/o=Exchange/ou=Recipients/cn=Declan"
        self.Name = "Declan Devine"
        self.PropertyAccessor = FakeAddressPropertyAccessor(smtp_address)
        self.smtp_address = smtp_address

    def GetExchangeUser(self):
        return FakeExchangeUser(self.smtp_address)


class FakeAddressPropertyAccessor:
    def __init__(self, smtp_address: str) -> None:
        self.smtp_address = smtp_address

    def GetProperty(self, property_name: str):
        if property_name == PR_SMTP_ADDRESS:
            return self.smtp_address
        raise KeyError(property_name)


class FakeRecipient:
    def __init__(self, smtp_address: str) -> None:
        self.Address = "/o=Exchange/ou=Recipients/cn=Declan"
        self.Name = "Declan Devine"
        self.AddressEntry = FakeAddressEntry(smtp_address)


class FakeFolder:
    def __init__(self, name: str, items=None, folders=None) -> None:
        self.Name = name
        self.Items = FakeItems(items or [])
        self.Folders = folders or []


class FakeSearchSession:
    def __init__(self, default_folder: FakeFolder, folders) -> None:
        self.default_folder = default_folder
        self.Folders = folders
        self.requested_folder = None

    def GetDefaultFolder(self, folder_id: int):
        self.requested_folder = folder_id
        return self.default_folder


class FakeOutboxSession:
    def __init__(self, outbox: FakeFolder) -> None:
        self.outbox = outbox
        self.requested_folder = None

    def GetDefaultFolder(self, folder_id: int):
        self.requested_folder = folder_id
        return self.outbox


class OutlookClientTests(unittest.TestCase):
    def test_set_deferred_delivery_sets_delivery_and_send_mapi_properties(self) -> None:
        message = FakeMessage()
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
            scheduled_send_at="2999-01-01 09:00",
        )

        OutlookClient._set_deferred_delivery(None, message, email)

        self.assertIsNotNone(message.DeferredDeliveryTime)
        self.assertIn(PR_DEFERRED_DELIVERY_TIME, message.PropertyAccessor.properties)
        self.assertIn(PR_DEFERRED_SEND_TIME, message.PropertyAccessor.properties)
        self.assertIn(PR_DEFERRED_SEND_NUMBER, message.PropertyAccessor.deleted)
        self.assertIn(PR_DEFERRED_SEND_UNITS, message.PropertyAccessor.deleted)
        self.assertTrue(message.saved)

    def test_delete_draft_permanently_deletes_after_moving_to_deleted_items(self) -> None:
        moved_message = FakeDeletedMessage()
        draft_message = FakeDraftMessage(moved_message)
        session = FakeSession(draft_message)
        client = OutlookClient.__new__(OutlookClient)
        client.outlook = FakeOutlook(session)

        result = client.delete_draft("draft-entry-id-123")

        self.assertEqual(OUTLOOK_DELETED_ITEMS_FOLDER, session.requested_folder)
        self.assertEqual("deleted-items-folder", draft_message.move_target)
        self.assertTrue(moved_message.deleted)
        self.assertIn("permanently deleted", result)

    def test_find_sent_email_checks_configured_project_folder(self) -> None:
        email = QueueEmail(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="recipient@example.com",
            recipient_type="Individual",
            template_type="General / Fallback",
            language="English",
            subject="Subject",
            body="Body",
            ready_to_send=True,
            validation_status="VALID",
            scheduled_send_at="2026-08-31 15:45",
        )
        default_sent = FakeFolder("Sent Items")
        project_folder = FakeFolder(
            "Commercial Outreach",
            items=[
                FakeSentItem(
                    "Subject",
                    "recipient@example.com",
                    datetime(2026, 8, 31, 15, 56),
                )
            ],
        )
        session = FakeSearchSession(
            default_sent,
            folders=[FakeFolder("Mailbox", folders=[project_folder])],
        )
        client = OutlookClient.__new__(OutlookClient)
        client.outlook = FakeOutlook(session)

        with patch("config.OUTLOOK_SENT_SEARCH_FOLDER_NAMES", ("Commercial Outreach",)):
            result = client.find_sent_email(email)

        self.assertEqual(OUTLOOK_SENT_MAIL_FOLDER, session.requested_folder)
        self.assertIn("Commercial Outreach", result)
        self.assertIn("2026-08-31 15:56:00", result)

    def test_find_sent_email_matches_exchange_recipient_smtp_address(self) -> None:
        email = QueueEmail(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="declan.devine@tus.ie",
            recipient_type="Individual",
            template_type="General / Fallback",
            language="English",
            subject="Invitation to contribute",
            body="Body",
            ready_to_send=True,
            validation_status="VALID",
            scheduled_send_at="2026-08-31 15:45",
        )
        default_sent = FakeFolder(
            "Sent Items",
            items=[
                FakeSentItem(
                    "Invitation to contribute",
                    "Declan Devine",
                    datetime(2026, 8, 31, 15, 56),
                    recipients=[FakeRecipient("declan.devine@tus.ie")],
                )
            ],
        )
        session = FakeSearchSession(default_sent, folders=[])
        client = OutlookClient.__new__(OutlookClient)
        client.outlook = FakeOutlook(session)

        result = client.find_sent_email(email)

        self.assertIn("Sent Items", result)
        self.assertIn("2026-08-31 15:56:00", result)

    def test_find_sent_email_compares_timezone_aware_sent_time_in_local_time(self) -> None:
        scheduled_at = datetime(2026, 9, 17, 15, 32)
        sent_on = scheduled_at.replace(second=30).astimezone(timezone.utc)
        email = QueueEmail(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="recipient@example.com",
            recipient_type="Individual",
            template_type="General / Fallback",
            language="English",
            subject="Invitation to contribute",
            body="Body",
            ready_to_send=True,
            validation_status="VALID",
            scheduled_send_at="2026-09-17 15:32",
        )
        default_sent = FakeFolder(
            "Sent Items",
            items=[
                FakeSentItem(
                    "Invitation to contribute",
                    "recipient@example.com",
                    sent_on,
                )
            ],
        )
        session = FakeSearchSession(default_sent, folders=[])
        client = OutlookClient.__new__(OutlookClient)
        client.outlook = FakeOutlook(session)

        result = client.find_sent_email(email)

        self.assertIn("Sent Items", result)

    def test_find_sent_emails_batches_multiple_matches_in_one_scan(self) -> None:
        emails = [
            QueueEmail(
                test_id=f"TEST-00{index}",
                target_name="Target",
                recipient_email=f"recipient{index}@example.com",
                recipient_type="Individual",
                template_type="General / Fallback",
                language="English",
                subject=f"Subject {index}",
                body="Body",
                ready_to_send=True,
                validation_status="VALID",
                scheduled_send_at="2026-08-31 15:45",
            )
            for index in (1, 2)
        ]
        default_sent = FakeFolder(
            "Sent Items",
            items=[
                FakeSentItem(
                    "Subject 1",
                    "recipient1@example.com",
                    datetime(2026, 8, 31, 15, 56),
                ),
                FakeSentItem(
                    "Subject 2",
                    "recipient2@example.com",
                    datetime(2026, 8, 31, 15, 57),
                ),
            ],
        )
        session = FakeSearchSession(default_sent, folders=[])
        client = OutlookClient.__new__(OutlookClient)
        client.outlook = FakeOutlook(session)

        result = client.find_sent_emails(emails)

        self.assertEqual(
            {
                ("TEST-001", "recipient1@example.com"),
                ("TEST-002", "recipient2@example.com"),
            },
            set(result),
        )
        self.assertTrue(default_sent.Items.sorted)
        self.assertIn("[SentOn] >=", default_sent.Items.restrict_filter)
        self.assertIn("15:56:00", result[("TEST-001", "recipient1@example.com")])
        self.assertIn("15:57:00", result[("TEST-002", "recipient2@example.com")])

    def test_find_outbox_emails_reads_without_sorting_collection(self) -> None:
        email = QueueEmail(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="recipient@example.com",
            recipient_type="Individual",
            template_type="General / Fallback",
            language="English",
            subject="Subject",
            body="Body",
            ready_to_send=True,
            validation_status="VALID",
            scheduled_send_at="2000-01-01 09:00",
        )
        message = FakeOutboxMessage("Subject", "recipient@example.com")
        message.DeferredDeliveryTime = ExplodingText()
        message.Submitted = ExplodingText()
        message.CreationTime = ExplodingText()
        outbox = FakeFolder("Outbox", items=[message])
        session = FakeOutboxSession(outbox)
        client = OutlookClient.__new__(OutlookClient)
        client.outlook = FakeOutlook(session)

        result = client.find_outbox_emails([email])

        self.assertEqual(OUTLOOK_OUTBOX_FOLDER, session.requested_folder)
        self.assertEqual({("TEST-001", "recipient@example.com"): "Found in Outlook Outbox"}, result)
        self.assertFalse(outbox.Items.sorted)
        self.assertFalse(message.saved)
        self.assertFalse(message.sent)

    def test_send_outbox_email_now_clears_deferred_delivery_and_sends_match(self) -> None:
        email = QueueEmail(
            test_id="TEST-001",
            target_name="Target",
            recipient_email="recipient@example.com",
            recipient_type="Individual",
            template_type="General / Fallback",
            language="English",
            subject="Subject",
            body="Body",
            ready_to_send=True,
            validation_status="VALID",
            scheduled_send_at="2000-01-01 09:00",
        )
        message = FakeOutboxMessage("Subject", "recipient@example.com")
        outbox = FakeFolder("Outbox", items=[message])
        session = FakeOutboxSession(outbox)
        client = OutlookClient.__new__(OutlookClient)
        client.outlook = FakeOutlook(session)

        result = client.send_outbox_email_now(email)

        self.assertIn("sent now", result)
        self.assertEqual(4, session.requested_folder)
        self.assertTrue(message.saved)
        self.assertTrue(message.sent)
        self.assertIn(PR_DEFERRED_DELIVERY_TIME, message.PropertyAccessor.deleted)
        self.assertIn(PR_DEFERRED_SEND_TIME, message.PropertyAccessor.deleted)


if __name__ == "__main__":
    unittest.main()
