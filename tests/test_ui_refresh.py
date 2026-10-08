from __future__ import annotations

import random
import unittest
from unittest.mock import Mock, patch

from outreach.scheduled_status import ScheduledStatusResult
from outreach.status import TargetStatus
from ui import OutreachApp, _email_progress_text, _initial_window_size, _visible_rows


class ImmediateThread:
    def __init__(self, target, args=(), daemon: bool = False) -> None:
        self.target = target
        self.args = args
        self.daemon = daemon

    def start(self) -> None:
        self.target(*self.args)


class NoopThread:
    def __init__(self, target, args=(), daemon: bool = False) -> None:
        self.target = target
        self.args = args
        self.daemon = daemon

    def start(self) -> None:
        return None


class FakeStringVar:
    def __init__(self, value="") -> None:
        self.value = value

    def set(self, value: str) -> None:
        self.value = value

    def get(self) -> str:
        return self.value


class FakeTable:
    def __init__(self) -> None:
        self.rows: dict[str, tuple[str, ...]] = {}
        self.next_id = 1
        self.deleted: list[str] = []
        self.sort_calls = 0

    def get_children(self) -> tuple[str, ...]:
        return tuple(self.rows)

    def delete(self, item: str) -> None:
        self.deleted.append(item)
        self.rows.pop(item, None)

    def insert(self, _parent: str, _index, values: tuple[str, ...], tags: tuple[str, ...] = ()) -> str:
        item = f"I{self.next_id:04d}"
        self.next_id += 1
        self.rows[item] = tuple(values)
        return item

    def apply_sort(self) -> None:
        self.sort_calls += 1


class FakeButton:
    def __init__(self) -> None:
        self.states: list[str] = []

    def configure(self, *, state: str) -> None:
        self.states.append(state)


class FakeText:
    def __init__(self) -> None:
        self.content = ""

    def delete(self, _start: str, _end: str) -> None:
        self.content = ""

    def insert(self, _index: str, text: str) -> None:
        self.content += text

    def see(self, _index: str) -> None:
        return None


class UiRefreshTests(unittest.TestCase):
    def test_initial_window_size_fits_target_resolutions(self) -> None:
        self.assertEqual((1600, 920), _initial_window_size(1920, 1080))
        self.assertEqual((1600, 980), _initial_window_size(1920, 1200))

    def _target(
        self,
        test_id: str,
        last_send_status: str = "",
        ready: str = "",
        enabled: str = "Yes",
    ) -> TargetStatus:
        return TargetStatus(
            test_id=test_id,
            target_name=f"Target {test_id}",
            recipient_email=f"{test_id.lower()}@example.com",
            enabled=enabled,
            handoff_generated="Yes" if last_send_status else "No",
            review_imported="Yes" if last_send_status else "No",
            ready_to_send=ready,
            validation_status="VALID" if last_send_status else "",
            scheduled_send_at="",
            last_send_status=last_send_status,
            last_send_time="",
            template_type="General / Fallback",
            language="English",
        )

    def test_email_progress_text_counts_sent_not_sent_and_failed_rows(self) -> None:
        rows = [
            self._target("TEST-001", "SENT", "Yes"),
            self._target("TEST-002", "DRAFT_CREATED", "No"),
            self._target("TEST-003"),
            self._target("TEST-004", "DELIVERY_FAILED", "Yes"),
            self._target("TEST-005", "REFUSED_TO_ANSWER", "Yes"),
        ]

        self.assertEqual("Sent: 0 | Not sent: 0 | Failed: 0", _email_progress_text([]))
        self.assertEqual("Sent: 1 | Not sent: 3 | Failed: 1", _email_progress_text(rows))
        self.assertEqual(rows, _visible_rows(rows, hide_sent=False))
        self.assertEqual(["TEST-002", "TEST-003"], [row.test_id for row in _visible_rows(rows, hide_sent=True)])
        self.assertEqual([], _visible_rows([self._target("TEST-004", "SENT")], hide_sent=True))
        self.assertEqual([], _visible_rows([self._target("TEST-005", "DELIVERY_FAILED")], hide_sent=True))
        self.assertEqual([], _visible_rows([self._target("TEST-006", "REFUSED_TO_ANSWER")], hide_sent=True))

    def test_visible_rows_can_hide_disabled_rows(self) -> None:
        rows = [
            self._target("TEST-001", enabled="Yes"),
            self._target("TEST-002", enabled="No"),
            self._target("TEST-003", "SENT", "Yes", enabled="No"),
            self._target("TEST-004", "DRAFT_CREATED", "No", enabled="Yes"),
        ]

        self.assertEqual(["TEST-001", "TEST-004"], [row.test_id for row in _visible_rows(rows, hide_sent=False, hide_disabled=True)])
        self.assertEqual(["TEST-001", "TEST-004"], [row.test_id for row in _visible_rows(rows, hide_sent=True, hide_disabled=True)])
        self.assertEqual(["TEST-001", "TEST-002", "TEST-004"], [row.test_id for row in _visible_rows(rows, hide_sent=True, hide_disabled=False)])

    def test_email_progress_text_fuzzes_status_combinations(self) -> None:
        rng = random.Random(20260921)
        statuses = [
            "",
            "SENT",
            "DELIVERY_FAILED",
            "REFUSED_TO_ANSWER",
            "DRAFT_CREATED",
            "OUTLOOK_SCHEDULED",
            "SEND_STATUS_UNKNOWN",
        ]

        for _case in range(200):
            rows = [
                self._target(
                    f"FUZZ-{index:04d}",
                    rng.choice(statuses),
                    rng.choice(["", "Yes", "No"]),
                    enabled=rng.choice(["Yes", "No"]),
                )
                for index in range(rng.randint(0, 250))
            ]

            sent = sum(1 for row in rows if row.last_send_status == "SENT")
            failed = sum(1 for row in rows if row.last_send_status == "DELIVERY_FAILED")
            not_sent = len(rows) - sent - failed

            self.assertEqual(
                f"Sent: {sent} | Not sent: {not_sent} | Failed: {failed}",
                _email_progress_text(rows),
            )

    def test_visible_rows_fuzzes_filter_combinations(self) -> None:
        rng = random.Random(20260921)
        terminal_statuses = {"SENT", "DELIVERY_FAILED", "REFUSED_TO_ANSWER"}
        statuses = [
            "",
            "SENT",
            "DELIVERY_FAILED",
            "REFUSED_TO_ANSWER",
            "DRAFT_CREATED",
            "OUTLOOK_SCHEDULED",
            "SEND_STATUS_UNKNOWN",
        ]

        for _case in range(200):
            rows = [
                self._target(
                    f"FILTER-{index:04d}",
                    rng.choice(statuses),
                    enabled=rng.choice(["Yes", "No"]),
                )
                for index in range(rng.randint(0, 250))
            ]
            hide_sent = rng.choice([False, True])
            hide_disabled = rng.choice([False, True])

            expected = [
                row
                for row in rows
                if (not hide_sent or row.last_send_status not in terminal_statuses)
                and (not hide_disabled or row.enabled == "Yes")
            ]

            self.assertEqual(expected, _visible_rows(rows, hide_sent, hide_disabled))

    def test_render_status_table_survives_repeated_hide_sent_toggles(self) -> None:
        app = OutreachApp.__new__(OutreachApp)
        app.rows = [
            self._target("TEST-001", "SENT", "Yes"),
            self._target("TEST-002", "DRAFT_CREATED", "No"),
            self._target("TEST-003"),
            self._target("TEST-004", "SENT", "Yes"),
            self._target("TEST-005", "DELIVERY_FAILED", "Yes"),
            self._target("TEST-006", enabled="No"),
            self._target("TEST-007", "REFUSED_TO_ANSWER", "Yes"),
        ]
        app.table = FakeTable()
        app.status_var = FakeStringVar()
        app.email_progress_var = FakeStringVar()
        app.hide_sent_var = FakeStringVar(False)
        app.hide_disabled_var = FakeStringVar(False)

        app.render_status_table()
        self.assertEqual(["TEST-001", "TEST-002", "TEST-003", "TEST-004", "TEST-005", "TEST-006", "TEST-007"], [values[0] for values in app.table.rows.values()])
        self.assertEqual("Sent: 2 | Not sent: 4 | Failed: 1", app.email_progress_var.get())
        self.assertIn("Loaded 7 target(s)", app.status_var.get())
        self.assertIn("Failed: 1", app.status_var.get())

        app.hide_sent_var.set(True)
        app.render_status_table()
        self.assertEqual(["TEST-002", "TEST-003", "TEST-006"], [values[0] for values in app.table.rows.values()])
        self.assertIn("Showing 3 active target(s)", app.status_var.get())

        app.render_status_table()
        self.assertEqual(["TEST-002", "TEST-003", "TEST-006"], [values[0] for values in app.table.rows.values()])

        app.hide_disabled_var.set(True)
        app.render_status_table()
        self.assertEqual(["TEST-002", "TEST-003"], [values[0] for values in app.table.rows.values()])

        app.hide_sent_var.set(False)
        app.render_status_table()
        self.assertEqual(["TEST-001", "TEST-002", "TEST-003", "TEST-004", "TEST-005", "TEST-007"], [values[0] for values in app.table.rows.values()])
        self.assertEqual(5, app.table.sort_calls)

    def test_render_status_table_stress_handles_large_status_sets(self) -> None:
        statuses = ["", "SENT", "DELIVERY_FAILED", "REFUSED_TO_ANSWER", "DRAFT_CREATED"]
        app = OutreachApp.__new__(OutreachApp)
        app.rows = [
            self._target(
                f"STRESS-{index:04d}",
                statuses[index % len(statuses)],
                ready="Yes" if index % 3 == 0 else "No",
                enabled="No" if index % 11 == 0 else "Yes",
            )
            for index in range(2500)
        ]
        app.table = FakeTable()
        app.status_var = FakeStringVar()
        app.email_progress_var = FakeStringVar()
        app.hide_sent_var = FakeStringVar(False)
        app.hide_disabled_var = FakeStringVar(False)

        app.render_status_table()

        self.assertEqual(2500, len(app.table.rows))
        self.assertEqual("Sent: 500 | Not sent: 1500 | Failed: 500", app.email_progress_var.get())
        self.assertIn("Loaded 2500 target(s)", app.status_var.get())
        self.assertIn("Failed: 500", app.status_var.get())

        app.hide_sent_var.set(True)
        app.hide_disabled_var.set(True)
        app.render_status_table()

        expected_visible = [
            row
            for row in app.rows
            if row.last_send_status not in {"SENT", "DELIVERY_FAILED", "REFUSED_TO_ANSWER"}
            and row.enabled == "Yes"
        ]
        self.assertEqual(len(expected_visible), len(app.table.rows))
        self.assertIn(f"Showing {len(expected_visible)} active target(s)", app.status_var.get())

    def test_refresh_button_checks_scheduled_sends_before_reloading_table(self) -> None:
        app = OutreachApp.__new__(OutreachApp)
        app.status_var = FakeStringVar()
        app.refresh_button = FakeButton()
        app.output = FakeText()
        app.after = lambda _delay, callback, *args: callback(*args)
        app.refresh_status = Mock()

        with (
            patch("ui.threading.Thread", ImmediateThread),
            patch(
                "ui.reconcile_scheduled_sends",
                return_value=ScheduledStatusResult(
                    checked=2,
                    marked_sent=1,
                    pending=1,
                    errors=(),
                    outbox_pending=1,
                    not_yet_due=0,
                ),
            ) as reconcile,
        ):
            app.refresh_with_scheduled_check()

        reconcile.assert_called_once_with()
        app.refresh_status.assert_called_once_with()
        self.assertEqual(["disabled", "normal"], app.refresh_button.states)
        self.assertIn("Checking Outlook sent folders", app.output.content)
        self.assertIn("(Please wait)", app.output.content)
        self.assertIn("Due scheduled rows checked: 2", app.output.content)
        self.assertIn("Marked SENT: 1", app.output.content)
        self.assertIn("Still pending or not yet due: 1", app.output.content)
        self.assertIn("Not yet due: 0", app.output.content)

    def test_run_entry_action_outputs_please_wait_while_worker_runs(self) -> None:
        app = OutreachApp.__new__(OutreachApp)
        app.status_var = FakeStringVar()
        app.output = FakeText()
        app.after = lambda _delay, callback, *args: callback(*args)
        app.refresh_status = Mock()

        with patch("ui.threading.Thread", ImmediateThread):
            app.run_entry_action("Creating Outlook draft for TEST-001", lambda: (1, 0, 0))

        self.assertIn("Creating Outlook draft for TEST-001 (Please wait)", app.output.content)
        self.assertIn("Completed: 1", app.output.content)

    def test_run_command_outputs_please_wait_while_worker_runs(self) -> None:
        app = OutreachApp.__new__(OutreachApp)
        app.status_var = FakeStringVar()
        app.output = FakeText()

        with patch("ui.threading.Thread", NoopThread):
            app.run_command(["validate"])

        self.assertEqual("Running: validate", app.status_var.get())
        self.assertIn("Running: python main.py validate (Please wait)", app.output.content)

    def test_set_selected_ready_uses_one_bulk_queue_update(self) -> None:
        app = OutreachApp.__new__(OutreachApp)
        app.status_var = FakeStringVar()
        app.undo_stack = []
        app.refresh_status = Mock()
        rows = [
            self._target("TEST-001", ready="No"),
            self._target("TEST-002", ready="No"),
        ]
        app.selected_statuses = Mock(return_value=rows)

        with patch("ui.set_ready_to_send_many") as set_ready:
            app.set_selected_ready(True)

        set_ready.assert_called_once()
        self.assertEqual({"TEST-001": True, "TEST-002": True}, set_ready.call_args.args[1])
        self.assertEqual(("ready", [("TEST-001", "No"), ("TEST-002", "No")]), app.undo_stack[-1])
        app.refresh_status.assert_called_once_with()
        self.assertIn("2 selected targets", app.status_var.get())

    def test_toggle_delivery_failed_confirms_and_refreshes(self) -> None:
        app = OutreachApp.__new__(OutreachApp)
        app.status_var = FakeStringVar()
        app.refresh_status = Mock()
        app.selected_status = Mock(
            return_value=TargetStatus(
                test_id="TEST-001",
                target_name="Target",
                recipient_email="test@example.com",
                enabled="Yes",
                handoff_generated="Yes",
                review_imported="Yes",
                ready_to_send="Yes",
                validation_status="VALID",
                scheduled_send_at="",
                last_send_status="SENT",
                last_send_time="",
                template_type="General / Fallback",
                language="English",
            )
        )

        with (
            patch("ui.messagebox.askyesno", return_value=True),
            patch("ui.toggle_delivery_failed_for_entry", return_value="DELIVERY_FAILED") as toggle,
        ):
            app.confirm_toggle_delivery_failed()

        toggle.assert_called_once_with("TEST-001")
        app.refresh_status.assert_called_once_with()
        self.assertEqual("TEST-001: logged DELIVERY_FAILED.", app.status_var.get())

    def test_toggle_refused_to_answer_confirms_and_refreshes(self) -> None:
        app = OutreachApp.__new__(OutreachApp)
        app.status_var = FakeStringVar()
        app.refresh_status = Mock()
        app.selected_status = Mock(
            return_value=TargetStatus(
                test_id="TEST-001",
                target_name="Target",
                recipient_email="test@example.com",
                enabled="Yes",
                handoff_generated="Yes",
                review_imported="Yes",
                ready_to_send="Yes",
                validation_status="VALID",
                scheduled_send_at="",
                last_send_status="SENT",
                last_send_time="",
                template_type="General / Fallback",
                language="English",
            )
        )

        with (
            patch("ui.messagebox.askyesno", return_value=True),
            patch("ui.toggle_refused_to_answer_for_entry", return_value="REFUSED_TO_ANSWER") as toggle,
        ):
            app.confirm_toggle_refused_to_answer()

        toggle.assert_called_once_with("TEST-001")
        app.refresh_status.assert_called_once_with()
        self.assertEqual("TEST-001: logged REFUSED_TO_ANSWER.", app.status_var.get())

    def test_cancel_outbox_confirms_and_runs_only_scheduled_rows(self) -> None:
        app = OutreachApp.__new__(OutreachApp)
        scheduled = self._target("TEST-001", "OUTLOOK_SCHEDULED", "Yes")
        not_scheduled = self._target("TEST-002", "DRAFT_CREATED", "Yes")
        app.selected_statuses = Mock(return_value=[scheduled, not_scheduled])
        app.run_entry_action = Mock()

        with patch("ui.messagebox.askyesno", return_value=True) as ask:
            app.confirm_cancel_outbox_selected()

        ask.assert_called_once()
        app.run_entry_action.assert_called_once()
        title, action = app.run_entry_action.call_args.args
        self.assertIn("TEST-001", title)
        with patch("ui.cancel_outlook_scheduled_for_entry", return_value=(1, 0, 0)) as cancel:
            result = action()
        self.assertEqual((1, 0, 0), result)
        cancel.assert_called_once_with("TEST-001")

    def test_fix_outbox_locked_confirms_and_runs_only_locked_rows(self) -> None:
        app = OutreachApp.__new__(OutreachApp)
        locked = self._target("TEST-001", "OUTLOOK_OUTBOX_LOCKED", "Yes")
        not_locked = self._target("TEST-002", "OUTLOOK_SCHEDULED", "Yes")
        app.selected_statuses = Mock(return_value=[locked, not_locked])
        app.run_entry_action = Mock()

        with patch("ui.messagebox.askyesno", return_value=True) as ask:
            app.confirm_fix_outbox_locked_selected()

        ask.assert_called_once()
        app.run_entry_action.assert_called_once()
        title, action = app.run_entry_action.call_args.args
        self.assertIn("TEST-001", title)
        with patch("ui.send_outbox_locked_for_entry", return_value=(1, 0, 0)) as repair:
            result = action()
        self.assertEqual((1, 0, 0), result)
        repair.assert_called_once_with("TEST-001")


if __name__ == "__main__":
    unittest.main()
