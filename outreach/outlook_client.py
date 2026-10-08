from __future__ import annotations

import os
from datetime import datetime, timezone
from collections.abc import Iterable
from multiprocessing import get_context

import config
from outreach.models import QueueEmail
from outreach.review_queue import parse_scheduled_send_at


OUTLOOK_MAIL_ITEM = 0
OUTLOOK_OUTBOX_FOLDER = 4
OUTLOOK_DELETED_ITEMS_FOLDER = 3
OUTLOOK_SENT_MAIL_FOLDER = 5
PR_DEFERRED_DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x000F0040"
PR_DEFERRED_SEND_TIME = "http://schemas.microsoft.com/mapi/proptag/0x3FEF0040"
PR_DEFERRED_SEND_NUMBER = "http://schemas.microsoft.com/mapi/proptag/0x3FEB0003"
PR_DEFERRED_SEND_UNITS = "http://schemas.microsoft.com/mapi/proptag/0x3FEC0003"
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"


class OutlookClient:
    def __init__(self) -> None:
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise RuntimeError(
                "pywin32 is not installed. Run: .\\.venv\\Scripts\\python.exe -m pip install pywin32"
            ) from exc

        pythoncom.CoInitialize()
        try:
            self.outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception as exc:
            raise RuntimeError(
                "Classic Outlook COM automation is not available. The visible app may be "
                "new Outlook for Windows, which does not support this automation path. "
                "Open classic Outlook and make sure the mailbox is loaded, or use the "
                "handoff/review workflow without Outlook automation."
            ) from exc

    def create_draft(self, email: QueueEmail) -> str:
        message = self._build_message(email)
        message.Save()
        entry_id = getattr(message, "EntryID", "")
        return f"Outlook draft saved{f' ({entry_id})' if entry_id else ''}"

    def send_email(self, email: QueueEmail) -> str:
        message = self._build_message(email)
        message.Send()
        return "Outlook send requested"

    def schedule_email(self, email: QueueEmail) -> str:
        message = self._build_message(email)
        self._set_deferred_delivery(message, email)
        message.Send()
        return f"Outlook deferred send requested ({email.scheduled_send_at})"

    def send_draft(self, email: QueueEmail, draft_entry_id: str) -> str:
        message = self._build_message(email)
        message.Send()
        delete_result = self.delete_draft(draft_entry_id)
        return f"Outlook send requested; {delete_result}"

    def schedule_draft(self, email: QueueEmail, draft_entry_id: str) -> str:
        message = self._build_message(email)
        self._set_deferred_delivery(message, email)
        message.Send()
        delete_result = self.delete_draft(draft_entry_id)
        return f"Outlook deferred send requested ({email.scheduled_send_at}); {delete_result}"

    def delete_draft(self, draft_entry_id: str) -> str:
        message = self.outlook.Session.GetItemFromID(draft_entry_id)
        if message is None:
            return f"Outlook draft not found ({draft_entry_id})"
        deleted_items = self.outlook.Session.GetDefaultFolder(OUTLOOK_DELETED_ITEMS_FOLDER)
        moved = message.Move(deleted_items)
        moved.Delete()
        return f"Outlook draft permanently deleted ({draft_entry_id})"

    def cancel_scheduled_email(self, email: QueueEmail) -> str:
        outbox = self.outlook.Session.GetDefaultFolder(OUTLOOK_OUTBOX_FOLDER)
        items = outbox.Items
        try:
            items.Sort("[DeferredDeliveryTime]", True)
        except Exception:
            pass

        checked = 0
        for index in range(int(getattr(items, "Count", 0) or 0), 0, -1):
            checked += 1
            if checked > 500:
                break
            try:
                item = items.Item(index)
            except Exception:
                continue
            if _scheduled_outbox_item_matches(item, email):
                item.Delete()
                return f"Outlook scheduled send cancelled ({email.scheduled_send_at or 'no scheduled time'})"
        return "Outlook scheduled send not found"

    def send_outbox_email_now(self, email: QueueEmail) -> str:
        outbox = self.outlook.Session.GetDefaultFolder(OUTLOOK_OUTBOX_FOLDER)
        items = outbox.Items
        try:
            items.Sort("[DeferredDeliveryTime]", True)
        except Exception:
            pass

        checked = 0
        for index in range(int(getattr(items, "Count", 0) or 0), 0, -1):
            checked += 1
            if checked > 1000:
                break
            try:
                item = items.Item(index)
            except Exception:
                continue
            if _scheduled_outbox_item_matches(item, email):
                self._clear_deferred_delivery(item)
                item.Send()
                return f"Outlook outbox message sent now ({email.scheduled_send_at or 'no scheduled time'})"
        return "Outlook scheduled outbox message not found"

    def find_sent_email(self, email: QueueEmail) -> str:
        scheduled_at = parse_scheduled_send_at(email.scheduled_send_at) if email.scheduled_send_at else None
        for folder_name, sent_items in self._sent_search_item_collections():
            if scheduled_at:
                sent_items = _restrict_sent_items_since(sent_items, scheduled_at)
            result = _find_sent_item_in_collection(sent_items, email, scheduled_at)
            if result:
                return f"Found in Outlook {folder_name}; SentOn={result}"
        return ""

    def find_sent_emails(self, emails: list[QueueEmail]) -> dict[tuple[str, str], str]:
        pending = {
            _email_key(email): (email, parse_scheduled_send_at(email.scheduled_send_at))
            for email in emails
        }
        found: dict[tuple[str, str], str] = {}
        if not pending:
            return found

        oldest_scheduled_at = min(scheduled_at for _email, scheduled_at in pending.values())
        pending_by_subject: dict[str, list[tuple[tuple[str, str], QueueEmail, object]]] = {}
        for key, (email, scheduled_at) in pending.items():
            pending_by_subject.setdefault(email.subject, []).append((key, email, scheduled_at))

        for folder_name, sent_items in self._sent_search_item_collections():
            if not pending:
                break
            sent_items = _restrict_sent_items_since(sent_items, oldest_scheduled_at)
            sorted_descending = _sort_sent_items(sent_items)
            checked = 0
            for item in sent_items:
                checked += 1
                if checked > 500 or not pending:
                    break

                sent_on = getattr(item, "SentOn", None)
                sent_on_naive = _naive_datetime(sent_on)
                if sorted_descending and sent_on_naive and sent_on_naive < oldest_scheduled_at:
                    break

                subject = str(getattr(item, "Subject", "") or "")
                candidates = pending_by_subject.get(subject, [])
                if not candidates:
                    continue

                for key, email, scheduled_at in list(candidates):
                    if key not in pending:
                        continue
                    if sent_on_naive and sent_on_naive < scheduled_at:
                        continue
                    if _sent_item_recipient_matches(item, email):
                        sent_on_text = str(sent_on or "")
                        found[key] = f"Found in Outlook {folder_name}; SentOn={sent_on_text}"
                        del pending[key]
        return found

    def find_outbox_emails(self, emails: list[QueueEmail]) -> dict[tuple[str, str], str]:
        pending = {_email_key(email): email for email in emails}
        found: dict[tuple[str, str], str] = {}
        if not pending:
            return found

        outbox = self.outlook.Session.GetDefaultFolder(OUTLOOK_OUTBOX_FOLDER)
        items = outbox.Items

        checked = 0
        for item in items:
            checked += 1
            if checked > 1000 or not pending:
                break
            for key, email in list(pending.items()):
                if _scheduled_outbox_item_matches(item, email):
                    found[key] = "Found in Outlook Outbox"
                    del pending[key]
                    break
        return found

    def _sent_search_item_collections(self):
        session = self.outlook.Session
        for folder in _find_named_folders(session, config.OUTLOOK_SENT_SEARCH_FOLDER_NAMES):
            try:
                yield getattr(folder, "Name", "configured sent folder"), folder.Items
            except Exception:
                continue

        default_sent = session.GetDefaultFolder(OUTLOOK_SENT_MAIL_FOLDER)
        yield "Sent Items", default_sent.Items

    def _build_message(self, email: QueueEmail):
        message = self.outlook.CreateItem(OUTLOOK_MAIL_ITEM)
        message.To = email.recipient_email
        message.Subject = email.subject
        message.Body = email.body
        return message

    def _set_deferred_delivery(self, message, email: QueueEmail) -> None:
        if not email.scheduled_send_at:
            raise RuntimeError("Scheduled Send At is blank.")
        scheduled_send_at = parse_scheduled_send_at(email.scheduled_send_at)
        scheduled_send_utc = scheduled_send_at.astimezone().astimezone(timezone.utc).replace(tzinfo=None)
        try:
            import pywintypes

            local_com_time = pywintypes.Time(scheduled_send_at)
            utc_com_time = pywintypes.Time(scheduled_send_utc)
        except ImportError:
            local_com_time = scheduled_send_at
            utc_com_time = scheduled_send_utc

        message.DeferredDeliveryTime = local_com_time
        accessor = message.PropertyAccessor
        _delete_mapi_property(accessor, PR_DEFERRED_SEND_NUMBER)
        _delete_mapi_property(accessor, PR_DEFERRED_SEND_UNITS)
        accessor.SetProperty(PR_DEFERRED_DELIVERY_TIME, local_com_time)
        accessor.SetProperty(PR_DEFERRED_SEND_TIME, utc_com_time)
        message.Save()

    def _clear_deferred_delivery(self, message) -> None:
        try:
            message.DeferredDeliveryTime = datetime(4501, 1, 1)
        except Exception:
            pass
        accessor = message.PropertyAccessor
        for property_name in (
            PR_DEFERRED_DELIVERY_TIME,
            PR_DEFERRED_SEND_TIME,
            PR_DEFERRED_SEND_NUMBER,
            PR_DEFERRED_SEND_UNITS,
        ):
            _delete_mapi_property(accessor, property_name)
        message.Save()


def _delete_mapi_property(accessor, property_name: str) -> None:
    try:
        accessor.DeleteProperty(property_name)
    except Exception:
        pass


def _find_sent_item_in_collection(items, email: QueueEmail, scheduled_at) -> str:
    _sort_sent_items(items)

    checked = 0
    for item in items:
        checked += 1
        if checked > 500:
            break
        if _sent_item_matches(item, email, scheduled_at):
            return str(getattr(item, "SentOn", "") or "")
    return ""


def _sort_sent_items(items) -> bool:
    try:
        items.Sort("[SentOn]", True)
        return True
    except Exception:
        return False


def _restrict_sent_items_since(items, sent_since: datetime):
    try:
        filter_value = sent_since.strftime("%m/%d/%Y %I:%M %p")
        return items.Restrict(f"[SentOn] >= '{filter_value}'")
    except Exception:
        return items


def _email_key(email: QueueEmail) -> tuple[str, str]:
    return (email.test_id, email.recipient_email.lower())


def _find_named_folders(session, folder_names: Iterable[str]):
    wanted = {folder_name.strip().lower() for folder_name in folder_names if folder_name.strip()}
    if not wanted:
        return

    seen: set[int] = set()
    for parent in _likely_search_parents(session):
        for folder in _folder_children(parent):
            if id(folder) in seen:
                continue
            if str(getattr(folder, "Name", "") or "").strip().lower() in wanted:
                seen.add(id(folder))
                yield folder


def _likely_search_parents(session) -> list[object]:
    parents: list[object] = []
    try:
        default_sent = session.GetDefaultFolder(OUTLOOK_SENT_MAIL_FOLDER)
        parent = getattr(default_sent, "Parent", None)
        if parent is not None:
            parents.append(parent)
    except Exception:
        pass

    try:
        default_store = getattr(session, "DefaultStore", None)
        if default_store is not None:
            parents.append(default_store.GetRootFolder())
    except Exception:
        pass

    parents.extend(_iter_com_collection(getattr(session, "Folders", ())))
    return parents


def _folder_children(folder) -> list[object]:
    direct_children = list(_iter_com_collection(getattr(folder, "Folders", ())))
    grandchildren: list[object] = []
    for child in direct_children:
        grandchildren.extend(_iter_com_collection(getattr(child, "Folders", ())))
    return direct_children + grandchildren


def run_outlook_action(
    email: QueueEmail,
    action: str,
    timeout_seconds: int,
    draft_entry_id: str = "",
) -> str:
    context = get_context("spawn")
    parent_connection, child_connection = context.Pipe(duplex=False)
    process = context.Process(
        target=_outlook_worker,
        args=(email, action, child_connection, draft_entry_id),
    )
    process.start()
    child_connection.close()
    process.join(timeout_seconds)

    if process.is_alive():
        process.terminate()
        process.join(5)
        raise TimeoutError(
            f"Outlook {action} timed out after {timeout_seconds} seconds. "
            "Check Outlook for profile, login, security, or modal prompts."
        )

    if not parent_connection.poll():
        raise RuntimeError(f"Outlook {action} failed without returning a result.")

    status, payload = parent_connection.recv()
    if status == "ok":
        return payload
    raise RuntimeError(payload)


def run_outlook_find_sent_emails(
    emails: list[QueueEmail],
    timeout_seconds: int,
) -> dict[tuple[str, str], str]:
    if not emails:
        return {}

    context = get_context("spawn")
    parent_connection, child_connection = context.Pipe(duplex=False)
    process = context.Process(
        target=_outlook_find_sent_worker,
        args=(emails, child_connection),
    )
    process.start()
    child_connection.close()
    process.join(timeout_seconds)

    if process.is_alive():
        process.terminate()
        process.join(5)
        raise TimeoutError(
            f"Outlook sent lookup timed out after {timeout_seconds} seconds. "
            "Check Outlook for profile, login, security, or modal prompts."
        )

    if not parent_connection.poll():
        raise RuntimeError("Outlook sent lookup failed without returning a result.")

    status, payload = parent_connection.recv()
    if status == "ok":
        return payload
    raise RuntimeError(payload)


def run_outlook_find_outbox_emails(
    emails: list[QueueEmail],
    timeout_seconds: int,
) -> dict[tuple[str, str], str]:
    if not emails:
        return {}

    context = get_context("spawn")
    parent_connection, child_connection = context.Pipe(duplex=False)
    process = context.Process(
        target=_outlook_find_outbox_worker,
        args=(emails, child_connection),
    )
    process.start()
    child_connection.close()
    process.join(timeout_seconds)

    if process.is_alive():
        process.terminate()
        process.join(5)
        raise TimeoutError(
            f"Outlook outbox lookup timed out after {timeout_seconds} seconds. "
            "Check Outlook for profile, login, security, or modal prompts."
        )

    if not parent_connection.poll():
        raise RuntimeError("Outlook outbox lookup failed without returning a result.")

    status, payload = parent_connection.recv()
    if status == "ok":
        return payload
    raise RuntimeError(payload)


def _outlook_find_sent_worker(emails: list[QueueEmail], connection) -> None:
    try:
        client = OutlookClient()
        connection.send(("ok", client.find_sent_emails(emails)))
        os._exit(0)
    except Exception as exc:
        connection.send(("error", str(exc)))
        os._exit(1)


def _outlook_find_outbox_worker(emails: list[QueueEmail], connection) -> None:
    try:
        client = OutlookClient()
        connection.send(("ok", client.find_outbox_emails(emails)))
        os._exit(0)
    except Exception as exc:
        connection.send(("error", str(exc)))
        os._exit(1)


def _outlook_worker(email: QueueEmail, action: str, connection, draft_entry_id: str) -> None:
    try:
        client = OutlookClient()
        if action == "draft":
            connection.send(("ok", client.create_draft(email)))
            os._exit(0)
        if action == "send-draft":
            connection.send(("ok", client.send_draft(email, draft_entry_id)))
            os._exit(0)
        if action == "schedule-draft":
            connection.send(("ok", client.schedule_draft(email, draft_entry_id)))
            os._exit(0)
        if action == "schedule":
            connection.send(("ok", client.schedule_email(email)))
            os._exit(0)
        if action == "delete-draft":
            connection.send(("ok", client.delete_draft(draft_entry_id)))
            os._exit(0)
        if action == "cancel-scheduled":
            connection.send(("ok", client.cancel_scheduled_email(email)))
            os._exit(0)
        if action == "send-outbox-now":
            connection.send(("ok", client.send_outbox_email_now(email)))
            os._exit(0)
        if action == "find-sent":
            connection.send(("ok", client.find_sent_email(email)))
            os._exit(0)
        if action == "send":
            connection.send(("ok", client.send_email(email)))
            os._exit(0)
        connection.send(("error", f"Unknown Outlook action: {action}"))
        os._exit(1)
    except Exception as exc:
        connection.send(("error", str(exc)))
        os._exit(1)


def _sent_item_matches(item, email: QueueEmail, scheduled_at) -> bool:
    if getattr(item, "Subject", "") != email.subject:
        return False

    sent_on = getattr(item, "SentOn", None)
    sent_on_naive = _naive_datetime(sent_on)
    if scheduled_at and sent_on_naive and sent_on_naive < scheduled_at:
        return False

    return _sent_item_recipient_matches(item, email)


def _scheduled_outbox_item_matches(item, email: QueueEmail) -> bool:
    if str(getattr(item, "Subject", "") or "").strip() != email.subject:
        return False
    return _sent_item_recipient_matches(item, email)


def _sent_item_recipient_matches(item, email: QueueEmail) -> bool:
    recipient = email.recipient_email.lower()

    to_text = str(getattr(item, "To", "") or "").lower()
    if recipient in to_text:
        return True

    try:
        for outlook_recipient in _iter_com_collection(item.Recipients):
            if _recipient_matches(outlook_recipient, recipient):
                return True
    except Exception:
        pass

    return not to_text


def _naive_datetime(value) -> datetime | None:
    if value is None:
        return None
    try:
        if getattr(value, "tzinfo", None) is not None:
            return value.astimezone().replace(tzinfo=None)
        return value.replace(tzinfo=None)
    except Exception:
        return None


def _recipient_matches(outlook_recipient, recipient_email: str) -> bool:
    values = [
        getattr(outlook_recipient, "Address", ""),
        getattr(outlook_recipient, "Name", ""),
    ]
    address_entry = getattr(outlook_recipient, "AddressEntry", None)
    if address_entry is not None:
        values.extend(
            [
                getattr(address_entry, "Address", ""),
                getattr(address_entry, "Name", ""),
            ]
        )
        try:
            exchange_user = address_entry.GetExchangeUser()
            values.append(getattr(exchange_user, "PrimarySmtpAddress", ""))
        except Exception:
            pass

        accessor = getattr(address_entry, "PropertyAccessor", None)
        if accessor is not None:
            try:
                values.append(accessor.GetProperty(PR_SMTP_ADDRESS))
            except Exception:
                pass

    return any(recipient_email in str(value or "").lower() for value in values)


def _iter_com_collection(collection):
    try:
        yield from collection
        return
    except TypeError:
        pass

    count = int(getattr(collection, "Count", 0) or 0)
    for index in range(1, count + 1):
        try:
            yield collection.Item(index)
        except Exception:
            continue
