from __future__ import annotations

import os
import re
import subprocess
import sys
import threading
import tkinter as tk
import difflib
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import config
from outreach.batch_schedule import (
    MIN_RANDOM_SPACING_MINUTES,
    apply_batch_schedule,
    maximum_random_spacing_that_fits_window,
    maximum_spacing_that_fits_window,
)
from outreach.entry_actions import (
    cancel_outlook_scheduled_for_entry,
    create_outlook_draft_for_entry,
    get_ai_checked_email,
    get_base_email,
    get_prompt_packet,
    regenerate_handoff_for_entries,
    schedule_outlook_for_entry,
    send_outbox_locked_for_entry,
    send_outlook_for_entry,
    toggle_delivery_failed_for_entry,
    toggle_refused_to_answer_for_entry,
)
from outreach.review_queue import (
    set_ready_to_send_many,
    set_scheduled_send_at,
    set_scheduled_send_at_many,
)
from outreach.reset_entry import reset_target_entry
from outreach.scheduled_status import ScheduledStatusResult, reconcile_scheduled_sends
from outreach.status import TargetStatus, load_file_lineage_status, load_target_statuses


BASE_DIR = Path(__file__).resolve().parent
HIDE_SENT_STATUSES = {"SENT", "DELIVERY_FAILED", "REFUSED_TO_ANSWER"}
MIN_BATCH_SCHEDULE_LEAD_MINUTES = 5


def _initial_window_size(screen_width: int, screen_height: int) -> tuple[int, int]:
    width = min(1600, max(1180, screen_width - 160))
    height = min(980, max(760, screen_height - 160))
    return width, height


class Tooltip:
    def __init__(self, widget, text: str) -> None:
        self.widget = widget
        self.text = text
        self.window: tk.Toplevel | None = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _event=None) -> None:
        if self.window or not self.text:
            return
        self.window = tk.Toplevel(self.widget)
        self.window.wm_overrideredirect(True)
        label = ttk.Label(
            self.window,
            text=self.text,
            padding=8,
            relief=tk.SOLID,
            borderwidth=1,
            background="#ffffe0",
            wraplength=360,
        )
        label.pack()
        self.window.update_idletasks()

        width = self.window.winfo_reqwidth()
        height = self.window.winfo_reqheight()
        screen_width = self.widget.winfo_screenwidth()
        screen_height = self.widget.winfo_screenheight()

        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        if x + width > screen_width:
            x = max(0, screen_width - width - 8)
        if y + height > screen_height:
            y = max(0, self.widget.winfo_rooty() - height - 8)
        self.window.wm_geometry(f"+{x}+{y}")

    def hide(self, _event=None) -> None:
        if self.window:
            self.window.destroy()
            self.window = None


class ScheduleDialog(tk.Toplevel):
    def __init__(self, parent, initial_value: str = "") -> None:
        super().__init__(parent)
        self.title("Scheduled Send")
        self.resizable(False, False)
        self.result: str | None = None
        self.transient(parent)
        self.grab_set()

        initial = self._parse_initial(initial_value)
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        self.date_var = tk.StringVar(value=initial.strftime("%Y-%m-%d"))
        self.hour_var = tk.StringVar(value=initial.strftime("%H"))
        self.minute_var = tk.StringVar(value=self._nearest_five_minute(initial.minute))

        dates = [
            (datetime.now() + timedelta(days=offset)).strftime("%Y-%m-%d")
            for offset in range(90)
        ]
        hours = [f"{hour:02d}" for hour in range(24)]
        minutes = [f"{minute:02d}" for minute in range(0, 60, 5)]

        ttk.Label(frame, text="Date").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(frame, text="Hour").grid(row=0, column=1, sticky="w", padx=(0, 8))
        ttk.Label(frame, text="Minute").grid(row=0, column=2, sticky="w")

        date_box = ttk.Combobox(frame, textvariable=self.date_var, values=dates, width=14, state="readonly")
        hour_box = ttk.Combobox(frame, textvariable=self.hour_var, values=hours, width=6, state="readonly")
        minute_box = ttk.Combobox(frame, textvariable=self.minute_var, values=minutes, width=8, state="readonly")
        date_box.grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=(4, 12))
        hour_box.grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=(4, 12))
        minute_box.grid(row=1, column=2, sticky="ew", pady=(4, 12))

        buttons = ttk.Frame(frame)
        buttons.grid(row=2, column=0, columnspan=3, sticky="e")
        ttk.Button(buttons, text="Clear", command=self.clear).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Cancel", command=self.cancel).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Save", command=self.save).pack(side=tk.LEFT)

        self.protocol("WM_DELETE_WINDOW", self.cancel)
        date_box.focus_set()
        self.bind("<Return>", lambda _event: self.save())
        self.bind("<Escape>", lambda _event: self.cancel())
        self.wait_window(self)

    def save(self) -> None:
        self.result = f"{self.date_var.get()} {self.hour_var.get()}:{self.minute_var.get()}"
        self.destroy()

    def clear(self) -> None:
        self.result = ""
        self.destroy()

    def cancel(self) -> None:
        self.result = None
        self.destroy()

    def _parse_initial(self, value: str) -> datetime:
        if value.strip():
            try:
                return datetime.strptime(value.strip().replace("T", " "), "%Y-%m-%d %H:%M")
            except ValueError:
                pass
        now = datetime.now() + timedelta(minutes=10)
        minute = ((now.minute + 4) // 5) * 5
        if minute == 60:
            now = now.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
        else:
            now = now.replace(minute=minute, second=0, microsecond=0)
        return now

    def _nearest_five_minute(self, minute: int) -> str:
        return f"{min(55, (minute // 5) * 5):02d}"


class BatchScheduleDialog(tk.Toplevel):
    def __init__(self, parent, count: int) -> None:
        super().__init__(parent)
        self.title("Batch Schedule")
        self.resizable(False, False)
        self.result: tuple[datetime, datetime, int, bool, int | None] | None = None
        self.count = count
        self.fixed_spacing_values = [value for value in range(5, 125, 5)]
        self.max_spacing_limit = 120
        self.transient(parent)
        self.grab_set()

        start = self._rounded_start()
        end = start + timedelta(hours=1)
        dates = [
            (datetime.now() + timedelta(days=offset)).strftime("%Y-%m-%d")
            for offset in range(90)
        ]
        hours = [f"{hour:02d}" for hour in range(24)]
        minutes = [f"{minute:02d}" for minute in range(0, 60, 5)]
        spacings = [str(value) for value in self.fixed_spacing_values]

        self.date_var = tk.StringVar(value=start.strftime("%Y-%m-%d"))
        self.start_hour_var = tk.StringVar(value=start.strftime("%H"))
        self.start_minute_var = tk.StringVar(value=start.strftime("%M"))
        self.end_hour_var = tk.StringVar(value=end.strftime("%H"))
        self.end_minute_var = tk.StringVar(value=end.strftime("%M"))
        self.spacing_var = tk.StringVar(value="10")
        self.randomize_var = tk.BooleanVar(value=False)
        self.max_spacing_var = tk.StringVar(value="30")

        frame = ttk.Frame(self, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=f"Selected ready emails: {count}").grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 10))

        labels = ["Date", "Start Hour", "Start Min", "End Hour", "End Min", "Spacing"]
        for column, label in enumerate(labels):
            ttk.Label(frame, text=label).grid(row=1, column=column, sticky="w", padx=(0, 8))

        self.date_box = ttk.Combobox(frame, textvariable=self.date_var, values=dates, width=14, state="readonly")
        self.start_hour_box = ttk.Combobox(frame, textvariable=self.start_hour_var, values=hours, width=8, state="readonly")
        self.start_minute_box = ttk.Combobox(frame, textvariable=self.start_minute_var, values=minutes, width=8, state="readonly")
        self.end_hour_box = ttk.Combobox(frame, textvariable=self.end_hour_var, values=hours, width=8, state="readonly")
        self.end_minute_box = ttk.Combobox(frame, textvariable=self.end_minute_var, values=minutes, width=8, state="readonly")
        self.spacing_box = ttk.Combobox(frame, textvariable=self.spacing_var, values=spacings, width=9, state="readonly")
        self.date_box.grid(row=2, column=0, sticky="ew", padx=(0, 8), pady=(4, 12))
        self.start_hour_box.grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=(4, 12))
        self.start_minute_box.grid(row=2, column=2, sticky="ew", padx=(0, 8), pady=(4, 12))
        self.end_hour_box.grid(row=2, column=3, sticky="ew", padx=(0, 8), pady=(4, 12))
        self.end_minute_box.grid(row=2, column=4, sticky="ew", padx=(0, 8), pady=(4, 12))
        self.spacing_box.grid(row=2, column=5, sticky="ew", pady=(4, 12))
        ttk.Checkbutton(frame, text="Randomize spacing", variable=self.randomize_var, command=self._sync_random_controls).grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 12))
        ttk.Label(frame, text="Max Spacing").grid(row=3, column=2, sticky="w", padx=(0, 8), pady=(0, 12))
        self.max_spacing_box = ttk.Combobox(frame, textvariable=self.max_spacing_var, values=spacings, width=9, state="disabled")
        self.max_spacing_box.grid(row=3, column=3, sticky="ew", padx=(0, 8), pady=(0, 12))

        buttons = ttk.Frame(frame)
        buttons.grid(row=4, column=0, columnspan=6, sticky="e")
        ttk.Button(buttons, text="Cancel", command=self.cancel).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Apply", command=self.apply).pack(side=tk.LEFT)

        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.bind("<Return>", lambda _event: self.apply())
        self.bind("<Escape>", lambda _event: self.cancel())
        for control in (
            self.date_box,
            self.start_hour_box,
            self.start_minute_box,
            self.end_hour_box,
            self.end_minute_box,
        ):
            control.bind("<<ComboboxSelected>>", lambda _event: self._sync_random_controls())
        self._sync_random_controls()
        self.wait_window(self)

    def apply(self) -> None:
        start, end = self._selected_window()
        minimum_start = datetime.now() + timedelta(minutes=MIN_BATCH_SCHEDULE_LEAD_MINUTES)
        if start < minimum_start:
            messagebox.showerror(
                "Schedule starts too soon",
                (
                    "Start the batch at least "
                    f"{MIN_BATCH_SCHEDULE_LEAD_MINUTES} minutes from now so Outlook "
                    "has time to keep delayed-send messages submitted."
                ),
            )
            return
        spacing = int(self.spacing_var.get())
        randomize = self.randomize_var.get()
        max_spacing = None
        if randomize:
            self._sync_random_controls()
            if not self.max_spacing_var.get():
                messagebox.showerror(
                    "Random spacing unavailable",
                    "The selected window is too short for randomized spacing.",
                )
                return
            max_spacing = int(self.max_spacing_var.get())
        self.result = (start, end, spacing, randomize, max_spacing)
        self.destroy()

    def cancel(self) -> None:
        self.result = None
        self.destroy()

    def _rounded_start(self) -> datetime:
        now = datetime.now() + timedelta(minutes=MIN_BATCH_SCHEDULE_LEAD_MINUTES)
        minute = ((now.minute + 4) // 5) * 5
        if minute == 60:
            return now.replace(minute=0, second=0, microsecond=0) + timedelta(hours=1)
        return now.replace(minute=minute, second=0, microsecond=0)

    def _sync_random_controls(self) -> None:
        randomize = self.randomize_var.get()
        self.spacing_box.configure(state="disabled" if randomize else "readonly")
        values = self._max_spacing_options()
        self.max_spacing_box.configure(
            values=values,
            state="readonly" if randomize and values else "disabled",
        )
        if self.max_spacing_var.get() not in values:
            self.max_spacing_var.set(values[-1] if values else "")

    def _selected_window(self) -> tuple[datetime, datetime]:
        date = self.date_var.get()
        start = datetime.strptime(
            f"{date} {self.start_hour_var.get()}:{self.start_minute_var.get()}",
            "%Y-%m-%d %H:%M",
        )
        end = datetime.strptime(
            f"{date} {self.end_hour_var.get()}:{self.end_minute_var.get()}",
            "%Y-%m-%d %H:%M",
        )
        return start, end

    def _max_spacing_options(self) -> list[str]:
        try:
            start, end = self._selected_window()
            max_fit = maximum_random_spacing_that_fits_window(self.count, start, end)
        except ValueError:
            return []
        return [
            str(value)
            for value in range(MIN_RANDOM_SPACING_MINUTES, min(max_fit, self.max_spacing_limit) + 1)
        ]


def _scheduled_check_output(result: ScheduledStatusResult) -> str:
    output = (
        "Checked scheduled sends\n"
        f"Due scheduled rows checked: {result.checked}\n"
        f"Marked SENT: {result.marked_sent}\n"
        f"Still pending or not yet due: {result.pending}\n"
        f"Not yet due: {result.not_yet_due}\n"
    )
    if result.errors:
        output += "Errors:\n"
        output += "\n".join(f"- {error}" for error in result.errors) + "\n"
    return output


def _please_wait(message: str) -> str:
    return f"{message} (Please wait)"


def _email_progress_text(rows: list[TargetStatus]) -> str:
    sent = sum(1 for row in rows if row.last_send_status == "SENT")
    failed = sum(1 for row in rows if row.last_send_status == "DELIVERY_FAILED")
    not_sent = len(rows) - sent - failed
    return f"Sent: {sent} | Not sent: {not_sent} | Failed: {failed}"


def _visible_rows(
    rows: list[TargetStatus],
    hide_sent: bool,
    hide_disabled: bool = False,
) -> list[TargetStatus]:
    visible = rows
    if hide_sent:
        visible = [row for row in visible if row.last_send_status not in HIDE_SENT_STATUSES]
    if hide_disabled:
        visible = [row for row in visible if row.enabled == "Yes"]
    return visible


class ColoredStatusTable(ttk.Frame):
    STATUS_COLUMNS = {"enabled", "handoff", "review", "ready"}
    SHIFT_MASK = 0x0001
    CONTROL_MASK = 0x0004

    def __init__(self, parent, columns: tuple[str, ...], height: int = 10, selectmode: str = "extended") -> None:
        super().__init__(parent)
        self.columns = columns
        self.selectmode = selectmode
        self.header_height = 28
        self.row_height = 26
        self.headings: dict[str, str] = {column: column for column in columns}
        self.widths: dict[str, int] = {column: 100 for column in columns}
        self.items: list[str] = []
        self.rows: dict[str, tuple[str, ...]] = {}
        self.row_tags: dict[str, tuple[str, ...]] = {}
        self.tag_styles: dict[str, dict[str, str]] = {}
        self.selected_items: set[str] = set()
        self.focus_item = ""
        self.anchor_item = ""
        self.next_item_id = 1
        self.render_pending = False
        self.sort_column: str | None = None
        self.sort_reverse = False

        self.header = tk.Canvas(self, height=self.header_height, highlightthickness=1, highlightbackground="#9aa0a6", background="#ffffff")
        self.body = tk.Canvas(self, height=height * self.row_height, highlightthickness=1, highlightbackground="#9aa0a6", background="#ffffff")
        self.header.grid(row=0, column=0, sticky="ew")
        self.body.grid(row=1, column=0, sticky="nsew")
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        self.header.bind("<Button-1>", self._on_header_click)
        self.body.bind("<Button-1>", self._on_click)
        self.body.bind("<MouseWheel>", self._on_mousewheel)
        self.body.bind("<Configure>", lambda _event: self._render())

    def configure(self, cnf=None, **kw):
        yscrollcommand = kw.pop("yscrollcommand", None)
        xscrollcommand = kw.pop("xscrollcommand", None)
        if yscrollcommand is not None:
            self.body.configure(yscrollcommand=yscrollcommand)
        if xscrollcommand is not None:
            self.body.configure(xscrollcommand=xscrollcommand)
        if cnf or kw:
            return super().configure(cnf, **kw)
        return None

    config = configure

    def heading(self, column: str, text: str) -> None:
        self.headings[column] = text
        self._schedule_render()

    def column(self, column: str, width: int, minwidth: int = 60, stretch: bool = False) -> None:
        self.widths[column] = max(width, minwidth)
        self._schedule_render()

    def tag_configure(self, tag: str, **kw) -> None:
        self.tag_styles[tag] = kw
        self._schedule_render()

    def insert(self, parent: str, index, values: tuple[str, ...], tags: tuple[str, ...] = ()) -> str:
        item = f"I{self.next_item_id:04d}"
        self.next_item_id += 1
        self.items.append(item)
        self.rows[item] = tuple(values)
        self.row_tags[item] = tuple(tags)
        self._schedule_render()
        return item

    def apply_sort(self) -> None:
        if self.sort_column is None:
            return
        self._sort_items(self.sort_column, self.sort_reverse)

    def delete(self, item: str) -> None:
        if item in self.items:
            self.items.remove(item)
        self.rows.pop(item, None)
        self.row_tags.pop(item, None)
        self.selected_items.discard(item)
        if self.focus_item == item:
            self.focus_item = ""
        self._schedule_render()

    def get_children(self, item: str = "") -> tuple[str, ...]:
        return tuple(self.items)

    def item(self, item: str, option: str | None = None):
        if option == "values":
            return self.rows.get(item, ())
        return {"values": self.rows.get(item, ()), "tags": self.row_tags.get(item, ())}

    def selection(self) -> tuple[str, ...]:
        return tuple(item for item in self.items if item in self.selected_items)

    def selection_set(self, items) -> None:
        if isinstance(items, str):
            items = [items]
        self.selected_items = {item for item in items if item in self.rows}
        self.anchor_item = next((item for item in items if item in self.rows), "")
        self._render_body()
        self.event_generate("<<TreeviewSelect>>")

    def selection_add(self, item: str) -> None:
        if item in self.rows:
            self.selected_items.add(item)
            if not self.anchor_item:
                self.anchor_item = item
            self._render_body()
            self.event_generate("<<TreeviewSelect>>")

    def selection_remove(self, item: str) -> None:
        self.selected_items.discard(item)
        self._render_body()
        self.event_generate("<<TreeviewSelect>>")

    def focus(self, item: str | None = None) -> str:
        if item is not None:
            self.focus_item = item if item in self.rows else ""
        return self.focus_item

    def identify_row(self, y: int) -> str:
        index = int(self.body.canvasy(y) // self.row_height)
        if 0 <= index < len(self.items):
            return self.items[index]
        return ""

    def yview(self, *args) -> None:
        self.body.yview(*args)

    def xview(self, *args) -> None:
        self.body.xview(*args)
        self.header.xview_moveto(self.body.xview()[0])

    def _on_header_click(self, event) -> str:
        column = self._identify_column(self.header.canvasx(event.x))
        if not column:
            return "break"
        reverse = not self.sort_reverse if self.sort_column == column else False
        self._sort_items(column, reverse)
        return "break"

    def _identify_column(self, canvas_x: float) -> str:
        x = 0
        for column in self.columns:
            x += self.widths[column]
            if canvas_x < x:
                return column
        return ""

    def _sort_items(self, column: str, reverse: bool) -> None:
        if column not in self.columns:
            return
        self.sort_column = column
        self.sort_reverse = reverse
        column_index = self.columns.index(column)

        nonblank: list[str] = []
        blank: list[str] = []
        for item in self.items:
            value = self._sort_cell_value(item, column_index)
            if value.strip():
                nonblank.append(item)
            else:
                blank.append(item)

        nonblank.sort(
            key=lambda item: self._sort_key(self._sort_cell_value(item, column_index)),
            reverse=reverse,
        )
        self.items = nonblank + blank
        self._render()

    def _sort_cell_value(self, item: str, column_index: int) -> str:
        values = self.rows.get(item, ())
        if column_index >= len(values):
            return ""
        return str(values[column_index])

    def _sort_key(self, value: str) -> tuple:
        cleaned = value.strip().lower()
        if cleaned in {"yes", "no"}:
            return (0, 0 if cleaned == "no" else 1)

        test_id_match = re.fullmatch(r"but-(\d+)", cleaned)
        if test_id_match:
            return (1, int(test_id_match.group(1)))

        number_parts = re.split(r"(\d+)", cleaned)
        natural_parts = tuple(int(part) if part.isdigit() else part for part in number_parts)
        return (2, natural_parts)

    def _on_click(self, event) -> str:
        item = self.identify_row(event.y)
        if not item:
            return "break"
        if event.state & self.SHIFT_MASK and self.selectmode == "extended":
            anchor = self.anchor_item or self.focus_item or item
            range_items = set(self._range_items(anchor, item))
            if event.state & self.CONTROL_MASK:
                self.selected_items.update(range_items)
            else:
                self.selected_items = range_items
        elif event.state & self.CONTROL_MASK and self.selectmode == "extended":
            if item in self.selected_items:
                self.selected_items.remove(item)
            else:
                self.selected_items.add(item)
            self.anchor_item = item
        else:
            self.selected_items = {item}
            self.anchor_item = item
        self.focus_item = item
        self._render_body()
        self.event_generate("<<TreeviewSelect>>")
        return "break"

    def _range_items(self, anchor: str, item: str) -> tuple[str, ...]:
        if anchor not in self.items or item not in self.items:
            return (item,)
        start = self.items.index(anchor)
        end = self.items.index(item)
        if start > end:
            start, end = end, start
        return tuple(self.items[start : end + 1])

    def _on_mousewheel(self, event) -> str:
        if event.delta:
            steps = max(1, abs(event.delta) // 120)
            direction = -1 if event.delta > 0 else 1
            self.body.yview_scroll(direction * steps, "units")
        return "break"

    def _render(self) -> None:
        self.render_pending = False
        self._render_header()
        self._render_body()

    def _schedule_render(self) -> None:
        if self.render_pending:
            return
        self.render_pending = True
        self.after_idle(self._render)

    def _render_header(self) -> None:
        total_width = sum(self.widths[column] for column in self.columns)
        self.header.configure(scrollregion=(0, 0, total_width, self.header_height))
        self.header.delete("all")
        x = 0
        for column in self.columns:
            width = self.widths[column]
            self.header.create_rectangle(x, 0, x + width, self.header_height, fill="#ffffff", outline="#dfe1e5")
            heading = self.headings[column]
            if column == self.sort_column:
                heading = f"{heading} {'v' if self.sort_reverse else '^'}"
            self.header.create_text(x + width // 2, self.header_height // 2, text=heading, font=("Segoe UI", 9, "bold"), fill="#202124")
            x += width

    def _render_body(self) -> None:
        total_width = sum(self.widths[column] for column in self.columns)
        total_height = max(len(self.items) * self.row_height, self.row_height)
        self.body.configure(scrollregion=(0, 0, total_width, total_height))
        self.body.delete("all")
        for row_index, item in enumerate(self.items):
            self._draw_row(item, row_index)

    def _draw_row(self, item: str, row_index: int) -> None:
        y = row_index * self.row_height
        selected = item in self.selected_items
        background = self._row_background(item, selected)
        values = self.rows.get(item, ())
        x = 0
        for column_index, column in enumerate(self.columns):
            width = self.widths[column]
            value = values[column_index] if column_index < len(values) else ""
            self.body.create_rectangle(x, y, x + width, y + self.row_height, fill=background, outline="#eef0f2")
            fill = self._cell_foreground(column, value, selected)
            font = ("Segoe UI", 9, "bold") if column in self.STATUS_COLUMNS and value else ("Segoe UI", 9)
            anchor = "center" if column in self.STATUS_COLUMNS else "w"
            text_x = x + width // 2 if anchor == "center" else x + 6
            self.body.create_text(text_x, y + self.row_height // 2, text=value, anchor=anchor, fill=fill, font=font)
            x += width

    def _row_background(self, item: str, selected: bool) -> str:
        if selected:
            return "#d9eaf7"
        for tag in self.row_tags.get(item, ()):
            background = self.tag_styles.get(tag, {}).get("background")
            if background:
                return background
        return "#ffffff"

    def _cell_foreground(self, column: str, value: str, selected: bool) -> str:
        if column in self.STATUS_COLUMNS:
            cleaned = value.strip().lower()
            if cleaned == "yes":
                return "#137333"
            if cleaned == "no":
                return "#a50e0e"
            return "#5f6368"
        return "#202124"


class OutreachApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"{config.CAMPAIGN_NAME} Outreach Control")
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width, height = _initial_window_size(screen_width, screen_height)
        self.geometry(f"{width}x{height}")
        self.minsize(1100, 680)

        self.status_var = tk.StringVar(value="Ready")
        self.lineage_var = tk.StringVar(value="")
        self.email_progress_var = tk.StringVar(value="Sent: 0 | Not sent: 0 | Failed: 0")
        self.hide_sent_var = tk.BooleanVar(value=False)
        self.hide_disabled_var = tk.BooleanVar(value=False)
        self.rows: list[TargetStatus] = []
        self.undo_stack: list[tuple[str, list[tuple[str, str]]]] = []

        self._build_layout()
        self.refresh_status()

    def _build_layout(self) -> None:
        top = ttk.Frame(self, padding=10)
        top.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(top)
        header.pack(fill=tk.X)
        ttk.Label(
            header,
            text=f"{config.CAMPAIGN_NAME} Outreach Pipeline",
            font=("Segoe UI", 14, "bold"),
        ).pack(side=tk.LEFT)
        self.refresh_button = ttk.Button(header, text="Refresh", command=self.refresh_with_scheduled_check)
        self.refresh_button.pack(side=tk.RIGHT)
        Tooltip(
            self.refresh_button,
            "Check Outlook for sent scheduled emails, then reload target status, file-lineage status, and send-log status from disk.",
        )
        ttk.Label(
            header,
            textvariable=self.email_progress_var,
            font=("Segoe UI", 10, "bold"),
        ).pack(side=tk.RIGHT, padx=(0, 8))
        hide_sent_check = ttk.Checkbutton(
            header,
            text="Hide sent",
            variable=self.hide_sent_var,
            command=self.render_status_table,
        )
        hide_sent_check.pack(side=tk.RIGHT, padx=(0, 8))
        Tooltip(hide_sent_check, "Hide rows already logged as SENT, DELIVERY_FAILED, or REFUSED_TO_ANSWER from the table.")
        hide_disabled_check = ttk.Checkbutton(
            header,
            text="Hide disabled",
            variable=self.hide_disabled_var,
            command=self.render_status_table,
        )
        hide_disabled_check.pack(side=tk.RIGHT, padx=(0, 8))
        Tooltip(hide_disabled_check, "Hide rows where Enabled is not Yes.")

        lineage_label = ttk.Label(top, textvariable=self.lineage_var, foreground="#7a4b00")
        lineage_label.pack(fill=tk.X, pady=(8, 0))

        pipeline_controls = ttk.LabelFrame(top, text="Pipeline", padding=10)
        pipeline_controls.pack(fill=tk.X, pady=(10, 0))
        self.add_button(pipeline_controls, "1 Validate", lambda: self.run_command(["validate"]), 0, 0, "Check source workbook templates and enabled target rows.")
        self.add_button(pipeline_controls, "2 Generate Handoff", lambda: self.run_command(["generate-handoff", "--include-generated"]), 0, 1, "Create the workbook to upload to ChatGPT Plus.")
        self.add_button(pipeline_controls, "3 Import Completed", self.import_handoff, 0, 2, "Import a completed ChatGPT workbook into the local review queue.")
        self.add_button(pipeline_controls, "Open Handoff", lambda: self.open_file(config.CHATGPT_HANDOFF_WORKBOOK_PATH), 1, 0, "Open the current ChatGPT handoff workbook.")
        self.add_button(pipeline_controls, "Open Review Queue", lambda: self.open_file(config.LOCAL_REVIEW_WORKBOOK_PATH), 1, 1, "Open the local review queue workbook.")
        self.add_button(pipeline_controls, "Open Send Log", lambda: self.open_file(config.SEND_LOG_PATH), 1, 2, "Open the CSV log of draft/send attempts.")
        for index in range(4):
            pipeline_controls.columnconfigure(index, weight=1)

        table_frame = ttk.Frame(top)
        table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(10, 0))
        columns = (
            "test_id",
            "target",
            "recipient",
            "enabled",
            "handoff",
            "review",
            "ready",
            "validation",
            "scheduled",
            "send_status",
            "send_time",
            "issues",
            "next_action",
            "template",
            "language",
        )
        self.table = ColoredStatusTable(table_frame, columns=columns, height=10, selectmode="extended")
        headings = {
            "test_id": "Test ID",
            "target": "Target",
            "recipient": "Recipient",
            "enabled": "Enabled",
            "handoff": "Handoff",
            "review": "Review",
            "ready": "Ready",
            "validation": "Validation",
            "scheduled": "Scheduled Send",
            "send_status": "Last Send",
            "send_time": "Last Send Time",
            "issues": "Issues",
            "next_action": "Next Action",
            "template": "Template",
            "language": "Language",
        }
        widths = {
            "test_id": 90,
            "target": 160,
            "recipient": 240,
            "enabled": 80,
            "handoff": 80,
            "review": 80,
            "ready": 80,
            "validation": 90,
            "scheduled": 160,
            "send_status": 140,
            "send_time": 180,
            "issues": 170,
            "next_action": 170,
            "template": 180,
            "language": 100,
        }
        for column in columns:
            self.table.heading(column, text=headings[column])
            self.table.column(column, width=widths[column], minwidth=60, stretch=False)
        self.table.tag_configure("complete", background="#f1f8f3")
        self.table.tag_configure("blocked", background="#fff3f1")
        self.table.tag_configure("pending", background="#fffdf0")

        scroll_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        scroll_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        scroll_y.configure(command=self._table_yview)
        scroll_x.configure(command=self._table_xview)
        self.table.configure(
            yscrollcommand=lambda *args: self._scrollbar_set(scroll_y, *args),
            xscrollcommand=lambda *args: self._scrollbar_set(scroll_x, *args),
        )
        self.table.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        bottom_pane = ttk.Frame(top)
        bottom_pane.pack(fill=tk.BOTH, expand=False, pady=(10, 0))
        bottom_pane.columnconfigure(0, weight=1, uniform="lower_columns")
        bottom_pane.columnconfigure(1, weight=1, uniform="lower_columns")
        bottom_pane.rowconfigure(0, weight=1, uniform="lower_rows")
        bottom_pane.rowconfigure(1, weight=1, uniform="lower_rows")

        review_controls = ttk.LabelFrame(bottom_pane, text="Review", padding=8)
        prepare_controls = ttk.LabelFrame(bottom_pane, text="Prepare", padding=8)
        delivery_controls = ttk.LabelFrame(bottom_pane, text="Outlook Delivery", padding=8)
        log_frame = ttk.LabelFrame(bottom_pane, text="Command Output", padding=8)
        review_controls.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))
        prepare_controls.grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(4, 0))
        delivery_controls.grid(row=0, column=1, sticky="nsew", padx=(4, 0), pady=(0, 4))
        log_frame.grid(row=1, column=1, sticky="nsew", padx=(4, 0), pady=(4, 0))

        self.add_button(review_controls, "Compare", self.compare_selected_email, 0, 0, "Compare the base email with the ChatGPT-reviewed email for the selected target.")
        self.add_button(review_controls, "View Prompt", self.view_selected_prompt_packet, 0, 1, "Open the ChatGPT prompt packet for the selected target.")
        self.add_button(review_controls, "Reset Entry", self.confirm_reset_selected, 1, 0, "Reset the selected target or targets by removing handoff rows, review rows, send log entries, and logged Outlook drafts.")
        self.add_button(review_controls, "Regenerate Handoff", self.confirm_regenerate_handoff_selected, 1, 1, "Regenerate handoff workbook rows for the selected target or targets.")
        self.add_button(prepare_controls, "Ready", lambda: self.set_selected_ready(True), 0, 0, "Set Ready To Send to Yes for the selected target or targets in the local review queue.")
        self.add_button(prepare_controls, "Not Ready", lambda: self.set_selected_ready(False), 0, 1, "Set Ready To Send to No for the selected target or targets.")
        self.add_button(prepare_controls, "Set Schedule", self.set_selected_schedule, 1, 0, "Open single-target scheduling for one selection, or batch scheduling for multiple selected ready/valid targets.")
        self.add_button(prepare_controls, "Undo Change", self.undo_last_prepare_change, 1, 1, "Undo the last Ready or Schedule change made in this UI session.")
        self.add_button(delivery_controls, "Make Draft", self.create_selected_draft, 0, 0, "Create Outlook drafts for the selected target or targets. This does not send the email.")
        self.add_button(delivery_controls, "Schedule In Outlook", self.confirm_schedule_selected, 0, 1, "Hand the selected target or targets to Outlook delayed delivery using their scheduled times.")
        self.add_button(delivery_controls, "Send Now", self.confirm_send_selected, 0, 2, "Send the selected target or targets immediately through Outlook, ignoring any scheduled time.")
        self.add_button(delivery_controls, "Toggle Failed", self.confirm_toggle_delivery_failed, 1, 0, "Mark the selected target as delivery failed after a bounce/undelivered message, or clear that manual flag.")
        self.add_button(delivery_controls, "Toggle Refused", self.confirm_toggle_refused_to_answer, 1, 1, "Mark the selected target as refused to answer, or clear that manual flag.")
        self.add_button(delivery_controls, "Cancel Outbox", self.confirm_cancel_outbox_selected, 1, 2, "Cancel selected scheduled Outlook outbox messages and set Ready To Send to No.")
        self.add_button(delivery_controls, "Fix Outbox Send", self.confirm_fix_outbox_locked_selected, 2, 0, "Send selected OUTLOOK_OUTBOX_LOCKED messages now from Outlook Outbox.", columnspan=3)
        for frame in (review_controls, prepare_controls):
            frame.columnconfigure(0, weight=1)
            frame.columnconfigure(1, weight=1)
        delivery_controls.columnconfigure(0, weight=1)
        delivery_controls.columnconfigure(1, weight=1)
        delivery_controls.columnconfigure(2, weight=1)

        self.output = tk.Text(log_frame, height=5, wrap=tk.WORD)
        self.output.pack(fill=tk.BOTH, expand=True)

        footer = ttk.Frame(self, padding=(10, 0, 10, 10))
        footer.pack(fill=tk.X)
        ttk.Label(footer, textvariable=self.status_var).pack(side=tk.LEFT)

    def add_button(self, parent, text: str, command, row: int, column: int, tooltip: str, columnspan: int = 1) -> ttk.Button:
        button = ttk.Button(parent, text=text, command=command)
        button.grid(row=row, column=column, columnspan=columnspan, padx=4, pady=3, sticky="ew")
        Tooltip(button, tooltip)
        return button

    def _table_yview(self, *args) -> None:
        self.table.yview(*args)

    def _table_xview(self, *args) -> None:
        self.table.xview(*args)

    def _scrollbar_set(self, scrollbar: ttk.Scrollbar, *args) -> None:
        scrollbar.set(*args)

    def refresh_status(self) -> None:
        lineage = load_file_lineage_status()
        prefix = "File lineage OK" if lineage.status == "OK" else f"File lineage {lineage.status}"
        self.lineage_var.set(f"{prefix}: {lineage.message}")

        try:
            self.rows = load_target_statuses()
        except Exception as exc:
            messagebox.showerror("Refresh failed", str(exc))
            return

        self.render_status_table()

    def render_status_table(self) -> None:
        for item in self.table.get_children():
            self.table.delete(item)

        visible_rows = _visible_rows(
            self.rows,
            self.hide_sent_var.get(),
            self.hide_disabled_var.get(),
        )
        for row in visible_rows:
            self.table.insert(
                "",
                tk.END,
                values=(
                    row.test_id,
                    row.target_name,
                    row.recipient_email,
                    row.enabled,
                    row.handoff_generated,
                    row.review_imported,
                    row.ready_to_send,
                    self._display_validation_status(row),
                    row.scheduled_send_at,
                    row.last_send_status,
                    row.last_send_time,
                    row.issues,
                    row.next_action,
                    row.template_type,
                    row.language,
                ),
                tags=(self._status_row_tag(row),),
            )
        self.table.apply_sort()

        ready = sum(1 for row in self.rows if row.ready_to_send.lower() == "yes")
        sent = sum(1 for row in self.rows if row.last_send_status == "SENT")
        failed = sum(1 for row in self.rows if row.last_send_status == "DELIVERY_FAILED")
        self.email_progress_var.set(_email_progress_text(self.rows))
        hidden_suffix = ""
        if len(visible_rows) != len(self.rows):
            hidden_suffix = f" Showing {len(visible_rows)} active target(s)."
        self.status_var.set(
            f"Loaded {len(self.rows)} target(s). Ready: {ready}. Sent: {sent}. Failed: {failed}."
            f"{hidden_suffix}"
        )

    def refresh_with_scheduled_check(self) -> None:
        self.status_var.set("Checking scheduled sends")
        self.refresh_button.configure(state="disabled")
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, f"{_please_wait('Checking Outlook sent folders for scheduled sends')}\n\n")
        thread = threading.Thread(target=self._refresh_with_scheduled_check_worker, daemon=True)
        thread.start()

    def _refresh_with_scheduled_check_worker(self) -> None:
        try:
            result = reconcile_scheduled_sends()
            output = _scheduled_check_output(result)
            returncode = 1 if result.errors else 0
        except Exception as exc:
            output = f"Scheduled send check failed: {exc}\n"
            returncode = 1
        self.after(0, self._refresh_with_scheduled_check_finished, returncode, output)

    def _refresh_with_scheduled_check_finished(self, returncode: int, output: str) -> None:
        self.output.insert(tk.END, output)
        self.output.see(tk.END)
        self.refresh_button.configure(state="normal")
        self.refresh_status()
        if returncode:
            self.status_var.set(f"{self.status_var.get()} Scheduled-send check had errors.")

    def _display_validation_status(self, row: TargetStatus) -> str:
        prerequisites = (
            row.enabled,
            row.handoff_generated,
            row.review_imported,
            row.ready_to_send,
        )
        if all(value.strip().lower() == "yes" for value in prerequisites):
            return row.validation_status
        return ""

    def _status_row_tag(self, row: TargetStatus) -> str:
        if row.last_send_status == "SENT":
            return "complete"
        if row.issues:
            return "blocked"
        return "pending"

    def import_handoff(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select completed ChatGPT handoff workbook",
            initialdir=str(config.REVIEW_DIR),
            filetypes=(("Excel workbooks", "*.xlsx"), ("All files", "*.*")),
        )
        if selected:
            self.run_command(["import-review", "--completed-handoff", selected])

    def selected_test_id(self) -> str | None:
        selected = self.table.selection()
        if not selected:
            messagebox.showwarning("No row selected", "Select one target row first.")
            return None
        values = self.table.item(selected[0], "values")
        return str(values[0])

    def compare_selected_email(self) -> None:
        test_id = self.selected_test_id()
        if not test_id:
            return
        try:
            base = get_base_email(test_id)
            ai_email = get_ai_checked_email(test_id)
        except Exception as exc:
            messagebox.showerror("Could not compare email", str(exc))
            return
        self.show_compare_window(
            f"Email Comparison - {test_id}",
            f"Subject:\n{base.subject}\n\nBody:\n{base.body}",
            f"Subject:\n{ai_email.subject}\n\nBody:\n{ai_email.body}",
        )

    def view_selected_prompt_packet(self) -> None:
        test_id = self.selected_test_id()
        if not test_id:
            return
        try:
            prompt = get_prompt_packet(test_id)
        except Exception as exc:
            messagebox.showerror("Could not load prompt packet", str(exc))
            return
        self.show_text_window(f"Prompt Packet - {test_id}", prompt)

    def confirm_regenerate_handoff_selected(self) -> None:
        rows = self.selected_statuses()
        if not rows:
            return
        enabled_blocked = [row.test_id for row in rows if row.enabled != "Yes"]
        if enabled_blocked:
            messagebox.showwarning(
                "Cannot regenerate handoff",
                f"Selected rows must be enabled in the source workbook.\n\nBlocked: {', '.join(enabled_blocked)}",
            )
            return
        ok = messagebox.askyesno(
            "Regenerate handoff rows",
            f"This will replace handoff workbook rows for {len(rows)} selected target(s). Continue?",
        )
        if not ok:
            return
        self.run_entry_action(
            f"Regenerating handoff rows for {len(rows)} selected target(s)",
            lambda: self._regenerate_handoff_rows(rows),
        )

    def create_selected_draft(self) -> None:
        rows = self.selected_statuses()
        if not rows:
            return

        if len(rows) == 1:
            row = rows[0]
            if row.last_send_status != "DRAFT_CREATED":
                create = messagebox.askyesno(
                    "Create Outlook draft",
                    (
                        f"{row.test_id} is not currently marked as having a draft ready.\n\n"
                        "Create an Outlook draft now?"
                    ),
                )
                if not create:
                    return
            self.run_entry_action(
                f"Creating Outlook draft for {row.test_id}",
                lambda: create_outlook_draft_for_entry(row.test_id),
            )
            return

        rows = self._confirm_batch_preflight("Create Outlook drafts", rows, "draft")
        if not rows:
            return
        self.run_entry_action(
            f"Creating Outlook drafts for {len(rows)} selected targets",
            lambda: self._run_for_selected_rows(rows, create_outlook_draft_for_entry),
        )

    def confirm_send_selected(self) -> None:
        rows = self.selected_statuses()
        if not rows:
            return
        single_selection = len(rows) == 1

        if single_selection:
            row = rows[0]
            message = (
                f"This will send {row.test_id} through Outlook now, ignoring any scheduled time. Continue?"
            )
            title = f"Sending {row.test_id} through Outlook now"
        else:
            rows = self._confirm_batch_preflight("Send selected targets now", rows, "send_now")
            if not rows:
                return
            message = (
                f"This will send {len(rows)} selected target(s) through Outlook now, "
                "ignoring any scheduled times. Continue?"
            )
            title = f"Sending {len(rows)} selected targets through Outlook now"

        if single_selection:
            ok = messagebox.askyesno("Confirm Outlook send now", message)
            if not ok:
                return
        self.run_entry_action(
            title,
            lambda: self._run_for_selected_rows(rows, send_outlook_for_entry),
        )

    def confirm_schedule_selected(self) -> None:
        rows = self.selected_statuses()
        if not rows:
            return

        if len(rows) == 1:
            row = rows[0]
            if not row.scheduled_send_at.strip():
                messagebox.showwarning("No schedule", "Set a send date/hour first.")
                return
            ok = messagebox.askyesno(
                "Confirm Outlook scheduled send",
                (
                    f"This will hand {row.test_id} to Outlook for delayed delivery at "
                    f"{row.scheduled_send_at}. Continue?"
                ),
            )
            if ok:
                self.run_entry_action(
                    f"Scheduling {row.test_id} in Outlook",
                    lambda: schedule_outlook_for_entry(row.test_id),
                )
            return

        rows = self._confirm_batch_preflight("Schedule selected targets in Outlook", rows, "schedule_outlook")
        if not rows:
            return
        ordered_rows = sorted(rows, key=lambda row: row.scheduled_send_at)
        self.run_entry_action(
            f"Scheduling {len(ordered_rows)} selected emails in Outlook",
            lambda: self._run_for_selected_rows(ordered_rows, schedule_outlook_for_entry),
        )

    def confirm_cancel_outbox_selected(self) -> None:
        rows = self.selected_statuses()
        if not rows:
            return

        scheduled_rows = [row for row in rows if row.last_send_status == "OUTLOOK_SCHEDULED"]
        non_scheduled = [row for row in rows if row.last_send_status != "OUTLOOK_SCHEDULED"]
        if not scheduled_rows:
            messagebox.showwarning(
                "Cancel Outlook scheduled sends",
                "No selected rows are currently logged as OUTLOOK_SCHEDULED.",
            )
            return

        blocked_preview = ""
        if non_scheduled:
            blocked_preview = (
                "\n\nThese selected rows are not logged as OUTLOOK_SCHEDULED and will be skipped:\n"
                + "\n".join(f"- {row.test_id}: {row.last_send_status or '<blank>'}" for row in non_scheduled[:12])
            )
            if len(non_scheduled) > 12:
                blocked_preview += f"\n...and {len(non_scheduled) - 12} more"

        if len(scheduled_rows) == 1:
            row = scheduled_rows[0]
            message = (
                f"This will cancel the scheduled Outlook outbox message for {row.test_id} "
                "and set Ready To Send to No.\n\nContinue?"
                + blocked_preview
            )
            title = f"Cancelling Outlook scheduled send for {row.test_id}"
        else:
            message = (
                f"This will cancel scheduled Outlook outbox messages for {len(scheduled_rows)} selected target(s) "
                "and set Ready To Send to No for each cancelled row.\n\nContinue?"
                + blocked_preview
            )
            title = f"Cancelling {len(scheduled_rows)} Outlook scheduled sends"

        if not messagebox.askyesno("Confirm cancel Outlook scheduled sends", message):
            return

        self.run_entry_action(
            title,
            lambda: self._run_for_selected_rows(scheduled_rows, cancel_outlook_scheduled_for_entry),
        )

    def confirm_fix_outbox_locked_selected(self) -> None:
        rows = self.selected_statuses()
        if not rows:
            return

        locked_rows = [row for row in rows if row.last_send_status == "OUTLOOK_OUTBOX_LOCKED"]
        non_locked = [row for row in rows if row.last_send_status != "OUTLOOK_OUTBOX_LOCKED"]
        if not locked_rows:
            messagebox.showwarning(
                "Fix Outlook outbox locked sends",
                "No selected rows are currently logged as OUTLOOK_OUTBOX_LOCKED.",
            )
            return

        blocked_preview = ""
        if non_locked:
            blocked_preview = (
                "\n\nThese selected rows are not OUTLOOK_OUTBOX_LOCKED and will be skipped:\n"
                + "\n".join(f"- {row.test_id}: {row.last_send_status or '<blank>'}" for row in non_locked[:12])
            )
            if len(non_locked) > 12:
                blocked_preview += f"\n...and {len(non_locked) - 12} more"

        if len(locked_rows) == 1:
            row = locked_rows[0]
            message = (
                f"This will find the stuck Outlook Outbox message for {row.test_id} "
                "and send it now.\n\nContinue?"
                + blocked_preview
            )
            title = f"Sending locked Outlook outbox message for {row.test_id}"
        else:
            message = (
                f"This will send {len(locked_rows)} selected stuck Outlook Outbox message(s) now.\n\n"
                "Continue?"
                + blocked_preview
            )
            title = f"Sending {len(locked_rows)} locked Outlook outbox messages now"

        if not messagebox.askyesno("Confirm fix Outlook outbox sends", message):
            return

        self.run_entry_action(
            title,
            lambda: self._run_for_selected_rows(locked_rows, send_outbox_locked_for_entry),
        )

    def confirm_toggle_delivery_failed(self) -> None:
        self._confirm_toggle_manual_flag(
            active_status="DELIVERY_FAILED",
            title_active="delivery failure",
            mark_title_prefix="Mark delivery failed for",
            clear_title_prefix="Clear delivery failure for",
            mark_description=(
                "Use this when Outlook shows an undelivered/bounce message for this recipient. "
                "The row will be blocked from future delivery actions until corrected or cleared."
            ),
            toggle_action=toggle_delivery_failed_for_entry,
        )

    def confirm_toggle_refused_to_answer(self) -> None:
        self._confirm_toggle_manual_flag(
            active_status="REFUSED_TO_ANSWER",
            title_active="refused-to-answer",
            mark_title_prefix="Mark refused to answer for",
            clear_title_prefix="Clear refused to answer for",
            mark_description=(
                "Use this when the recipient explicitly refuses to answer. "
                "The row will be excluded from future handoffs and blocked from delivery until cleared."
            ),
            toggle_action=toggle_refused_to_answer_for_entry,
        )

    def _confirm_toggle_manual_flag(
        self,
        *,
        active_status: str,
        title_active: str,
        mark_title_prefix: str,
        clear_title_prefix: str,
        mark_description: str,
        toggle_action,
    ) -> None:
        row = self.selected_status()
        if row is None:
            return

        if row.last_send_status == active_status:
            title = f"{clear_title_prefix} {row.test_id}"
            message = (
                f"{row.test_id} is currently marked {active_status}.\n\n"
                "Clear this manual flag?"
            )
        else:
            title = f"{mark_title_prefix} {row.test_id}"
            message = (
                f"Mark {row.test_id} as {active_status}?\n\n"
                f"{mark_description}"
            )

        if not messagebox.askyesno(title, message):
            return
        try:
            new_status = toggle_action(row.test_id)
        except Exception as exc:
            messagebox.showerror(f"Could not update {title_active} flag", str(exc))
            return

        self.status_var.set(f"{row.test_id}: logged {new_status}.")
        self.refresh_status()

    def selected_status(self) -> TargetStatus | None:
        test_id = self.selected_test_id()
        if not test_id:
            return None
        for row in self.rows:
            if row.test_id == test_id:
                return row
        return None

    def selected_statuses(self) -> list[TargetStatus]:
        selected = self.table.selection()
        if not selected:
            messagebox.showwarning("No rows selected", "Select one or more target rows first.")
            return []

        selected_ids = {
            str(self.table.item(item, "values")[0])
            for item in selected
        }
        return [row for row in self.rows if row.test_id in selected_ids]

    def set_selected_schedule(self) -> None:
        rows = self.selected_statuses()
        if not rows:
            return

        if len(rows) == 1:
            row = rows[0]
            test_id = row.test_id
            if row.last_send_status == "SENT":
                messagebox.showwarning("Already sent", f"{test_id} is already logged as SENT.")
                return
            if row.last_send_status != "DRAFT_CREATED":
                proceed = messagebox.askyesno(
                    "Draft not logged",
                    (
                        f"{test_id} is not currently logged as DRAFT_CREATED.\n\n"
                        "Set a scheduled send time anyway?"
                    ),
                )
                if not proceed:
                    return

            dialog = ScheduleDialog(self, row.scheduled_send_at)
            if dialog.result is None:
                return
            try:
                set_scheduled_send_at(config.LOCAL_REVIEW_WORKBOOK_PATH, test_id, dialog.result)
            except Exception as exc:
                messagebox.showerror("Could not set schedule", str(exc))
                return
            self._push_undo(
                "schedule",
                [(test_id, row.scheduled_send_at)],
            )
            self.status_var.set(f"{test_id}: scheduled send set to {dialog.result or '<blank>'}.")
            self.refresh_status()
            return

        blocked = [
            row.test_id
            for row in rows
            if row.ready_to_send.lower() != "yes"
            or row.validation_status != "VALID"
            or row.last_send_status in {"SENT", "OUTLOOK_SCHEDULED"}
        ]
        if blocked:
            messagebox.showwarning(
                "Cannot batch schedule",
                (
                    "Only selected rows that are Ready = Yes, Validation = VALID, "
                    "and not already SENT/OUTLOOK_SCHEDULED can be scheduled.\n\n"
                    f"Blocked: {', '.join(blocked)}"
                ),
            )
            return

        ordered_rows = sorted(rows, key=lambda row: row.test_id)
        dialog = BatchScheduleDialog(self, len(ordered_rows))
        if dialog.result is None:
            return
        start, end, spacing, randomize, max_spacing = dialog.result
        previous_schedules = [(row.test_id, row.scheduled_send_at) for row in ordered_rows]
        try:
            assignments = apply_batch_schedule(
                config.LOCAL_REVIEW_WORKBOOK_PATH,
                [row.test_id for row in ordered_rows],
                start,
                end,
                spacing,
                randomize=randomize,
                max_spacing_minutes=max_spacing,
            )
        except Exception as exc:
            messagebox.showerror("Could not batch schedule", str(exc))
            return
        self._push_undo("schedule", previous_schedules)

        first = assignments[0].scheduled_send_at if assignments else ""
        last = assignments[-1].scheduled_send_at if assignments else ""
        self.status_var.set(
            f"Scheduled {len(assignments)} selected target(s)"
            + (f" from {first} to {last}." if assignments else ".")
        )
        self.refresh_status()

    def confirm_reset_selected(self) -> None:
        rows = self.selected_statuses()
        if not rows:
            return
        test_ids = [row.test_id for row in rows]
        if len(test_ids) == 1:
            prompt = (
                f"This will reset {test_ids[0]} by removing its handoff row, review row, "
                "send log entries, and any logged Outlook draft. Continue?"
            )
            title = f"Resetting {test_ids[0]}"
            command_label = f"reset {test_ids[0]}"
        else:
            prompt = (
                f"This will reset {len(test_ids)} selected targets by removing their handoff rows, "
                "review rows, send log entries, and any logged Outlook drafts. Continue?"
            )
            title = f"Resetting {len(test_ids)} selected targets"
            command_label = f"reset {len(test_ids)} selected targets"

        ok = messagebox.askyesno(
            "Reset target entry" if len(test_ids) == 1 else "Reset selected entries",
            prompt,
        )
        if not ok:
            return
        self.status_var.set(title)
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, f"{_please_wait(title)}\n\n")
        thread = threading.Thread(target=self._reset_worker, args=(test_ids, command_label), daemon=True)
        thread.start()

    def _reset_worker(self, test_ids: list[str], command_label: str) -> None:
        output_parts: list[str] = []
        failures = 0
        for test_id in test_ids:
            try:
                reset = reset_target_entry(test_id)
                output = (
                    f"Reset {test_id}\n"
                    f"Review row removed: {'Yes' if reset.review_removed else 'No'}\n"
                    f"Handoff row removed: {'Yes' if reset.handoff_removed else 'No'}\n"
                    f"Send log entries removed: {reset.log_entries_removed}\n"
                    f"Outlook drafts deleted: {reset.outlook_drafts_deleted}\n"
                )
                if reset.outlook_errors:
                    output += "Outlook draft delete errors:\n"
                    output += "\n".join(f"- {error}" for error in reset.outlook_errors) + "\n"
                    failures += 1
            except Exception as exc:
                output = f"Reset {test_id}\nFailed: {exc}\n"
                failures += 1
            output_parts.append(output)

        returncode = 1 if failures else 0
        self.after(0, self._command_finished, [command_label], returncode, "\n".join(output_parts))

    def run_entry_action(self, title: str, action) -> None:
        self.status_var.set(title)
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, f"{_please_wait(title)}\n\n")
        thread = threading.Thread(target=self._entry_action_worker, args=(title, action), daemon=True)
        thread.start()

    def _entry_action_worker(self, title: str, action) -> None:
        try:
            completed, skipped, blocked = action()
            output = (
                f"{title}\n"
                f"Completed: {completed}\n"
                f"Skipped: {skipped}\n"
                f"Blocked/failed: {blocked}\n"
                f"Send log: {config.SEND_LOG_PATH}\n"
            )
            returncode = 1 if blocked else 0
        except Exception as exc:
            output = f"{title}\nFailed: {exc}\n"
            returncode = 1
        self.after(0, self._command_finished, [title], returncode, output)

    def set_selected_ready(self, ready: bool) -> None:
        rows = self.selected_statuses()
        if not rows:
            return
        previous_ready = [(row.test_id, row.ready_to_send) for row in rows]
        try:
            set_ready_to_send_many(
                config.LOCAL_REVIEW_WORKBOOK_PATH,
                {row.test_id: ready for row in rows},
            )
        except Exception as exc:
            messagebox.showerror("Could not update review queue", str(exc))
            return
        self._push_undo("ready", previous_ready)
        if len(rows) == 1:
            self.status_var.set(f"{rows[0].test_id}: Ready To Send set to {'Yes' if ready else 'No'}.")
        else:
            self.status_var.set(
                f"{len(rows)} selected targets: Ready To Send set to {'Yes' if ready else 'No'}."
            )
        self.refresh_status()

    def _run_for_selected_rows(self, rows: list[TargetStatus], action) -> tuple[int, int, int]:
        completed = 0
        skipped = 0
        blocked = 0
        for row in rows:
            try:
                result = action(row.test_id)
            except Exception as exc:
                print(f"{row.test_id}: failed: {exc}")
                blocked += 1
                continue
            completed += result[0]
            skipped += result[1]
            blocked += result[2]
        return completed, skipped, blocked

    def _regenerate_handoff_rows(self, rows: list[TargetStatus]) -> tuple[int, int, int]:
        completed, path = regenerate_handoff_for_entries([row.test_id for row in rows])
        print(f"Handoff workbook updated: {path}")
        return completed, 0, 0

    def _push_undo(self, kind: str, values: list[tuple[str, str]]) -> None:
        self.undo_stack.append((kind, values))
        self.undo_stack = self.undo_stack[-20:]

    def undo_last_prepare_change(self) -> None:
        if not self.undo_stack:
            messagebox.showinfo("Undo Change", "No Ready or Schedule change to undo.")
            return
        kind, values = self.undo_stack.pop()
        try:
            if kind == "ready":
                set_ready_to_send_many(
                    config.LOCAL_REVIEW_WORKBOOK_PATH,
                    {
                        test_id: previous_ready.strip().lower() == "yes"
                        for test_id, previous_ready in values
                    },
                )
            elif kind == "schedule":
                set_scheduled_send_at_many(
                    config.LOCAL_REVIEW_WORKBOOK_PATH,
                    dict(values),
                )
            else:
                raise ValueError(f"Unsupported undo operation: {kind}")
        except Exception as exc:
            messagebox.showerror("Could not undo change", str(exc))
            return
        self.status_var.set(f"Undid {kind} change for {len(values)} target(s).")
        self.refresh_status()

    def _confirm_batch_preflight(
        self,
        title: str,
        rows: list[TargetStatus],
        mode: str,
    ) -> list[TargetStatus]:
        eligible: list[TargetStatus] = []
        blocked: list[str] = []
        for row in rows:
            issue = self._preflight_issue(row, mode)
            if issue:
                blocked.append(f"{row.test_id}: {issue}")
            else:
                eligible.append(row)

        if not eligible:
            messagebox.showwarning(title, "No selected rows are eligible.\n\n" + "\n".join(blocked[:12]))
            return []
        if blocked:
            more = "" if len(blocked) <= 12 else f"\n...and {len(blocked) - 12} more"
            ok = messagebox.askyesno(
                title,
                (
                    f"Eligible rows: {len(eligible)}\n"
                    f"Blocked rows: {len(blocked)}\n\n"
                    + "\n".join(blocked[:12])
                    + more
                    + "\n\nContinue with eligible rows only?"
                ),
            )
            if not ok:
                return []
        else:
            ok = messagebox.askyesno(title, f"All {len(eligible)} selected rows are eligible. Continue?")
            if not ok:
                return []
        return eligible

    def _preflight_issue(self, row: TargetStatus, mode: str) -> str:
        if row.enabled != "Yes":
            return "disabled"
        if row.review_imported != "Yes":
            return "not imported into review queue"
        if row.ready_to_send != "Yes":
            return "not ready"
        if row.validation_status != "VALID":
            return "validation is not VALID"
        if row.last_send_status == "SENT":
            return "already sent"
        if row.last_send_status == "OUTLOOK_SCHEDULED":
            return "already scheduled in Outlook"
        if row.last_send_status == "OUTLOOK_OUTBOX_LOCKED":
            return "stuck in Outlook Outbox; use Fix Outbox Send"
        if row.last_send_status == "SEND_STATUS_UNKNOWN":
            return "unknown send status requires manual review"
        if row.last_send_status == "DELIVERY_FAILED":
            return "delivery failed; review recipient email"
        if row.last_send_status == "REFUSED_TO_ANSWER":
            return "recipient refused to answer"
        if mode == "draft" and row.last_send_status == "DRAFT_CREATED":
            return "draft already created"
        if mode == "schedule_outlook":
            if not row.scheduled_send_at.strip():
                return "missing schedule"
            if row.last_send_status != "DRAFT_CREATED":
                return "draft not created"
        return ""

    def show_email_window(self, title: str, subject: str, body: str) -> None:
        window = tk.Toplevel(self)
        window.title(title)
        window.geometry("900x650")
        frame = ttk.Frame(window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Subject", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        subject_box = tk.Text(frame, height=3, wrap=tk.WORD)
        subject_box.insert(tk.END, subject)
        subject_box.configure(state=tk.DISABLED)
        subject_box.pack(fill=tk.X, pady=(2, 10))

        ttk.Label(frame, text="Body", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        body_box = tk.Text(frame, wrap=tk.WORD)
        body_box.insert(tk.END, body)
        body_box.configure(state=tk.DISABLED)
        body_box.pack(fill=tk.BOTH, expand=True)

    def show_text_window(self, title: str, text: str) -> None:
        window = tk.Toplevel(self)
        window.title(title)
        window.geometry("900x650")
        frame = ttk.Frame(window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        text_box = tk.Text(frame, wrap=tk.WORD)
        scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text_box.yview)
        text_box.configure(yscrollcommand=scroll.set)
        text_box.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        text_box.insert(tk.END, text)
        text_box.configure(state=tk.DISABLED)

    def show_compare_window(self, title: str, base_text: str, ai_text: str) -> None:
        window = tk.Toplevel(self)
        window.title(title)
        window.geometry("1200x700")
        frame = ttk.Frame(window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(1, weight=1)

        ttk.Label(frame, text="Base Email", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(frame, text="AI Email", font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky="w")

        base_box = tk.Text(frame, wrap=tk.WORD)
        ai_box = tk.Text(frame, wrap=tk.WORD)
        base_box.grid(row=1, column=0, sticky="nsew", padx=(0, 6), pady=(4, 0))
        ai_box.grid(row=1, column=1, sticky="nsew", padx=(6, 0), pady=(4, 0))
        for text_box in (base_box, ai_box):
            text_box.tag_configure("changed", background="#fff4ce")
            text_box.tag_configure("inserted", background="#e7f4ea")
            text_box.tag_configure("deleted", background="#fce8e6")

        self._insert_diff_text(base_box, ai_box, base_text, ai_text)
        base_box.configure(state=tk.DISABLED)
        ai_box.configure(state=tk.DISABLED)

    def _insert_diff_text(self, base_box: tk.Text, ai_box: tk.Text, base_text: str, ai_text: str) -> None:
        base_lines = base_text.splitlines()
        ai_lines = ai_text.splitlines()
        matcher = difflib.SequenceMatcher(None, base_lines, ai_lines)

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                for line in base_lines[i1:i2]:
                    base_box.insert(tk.END, line + "\n")
                for line in ai_lines[j1:j2]:
                    ai_box.insert(tk.END, line + "\n")
                continue

            base_tag = "deleted" if tag == "delete" else "changed"
            ai_tag = "inserted" if tag == "insert" else "changed"
            base_chunk = base_lines[i1:i2] or [""]
            ai_chunk = ai_lines[j1:j2] or [""]
            row_count = max(len(base_chunk), len(ai_chunk))
            for index in range(row_count):
                if index < len(base_chunk):
                    self._insert_tagged_line(base_box, base_chunk[index], base_tag)
                else:
                    base_box.insert(tk.END, "\n")
                if index < len(ai_chunk):
                    self._insert_tagged_line(ai_box, ai_chunk[index], ai_tag)
                else:
                    ai_box.insert(tk.END, "\n")

    def _insert_tagged_line(self, text_box: tk.Text, line: str, tag: str) -> None:
        start = text_box.index(tk.END)
        text_box.insert(tk.END, line + "\n")
        end = text_box.index(tk.END)
        text_box.tag_add(tag, start, end)

    def run_command(self, args: list[str]) -> None:
        self.status_var.set(f"Running: {' '.join(args)}")
        self.output.delete("1.0", tk.END)
        command_text = f"Running: python main.py {' '.join(args)}"
        self.output.insert(tk.END, f"{_please_wait(command_text)}\n\n")
        thread = threading.Thread(target=self._run_command_worker, args=(args,), daemon=True)
        thread.start()

    def _run_command_worker(self, args: list[str]) -> None:
        process = subprocess.run(
            [sys.executable, "main.py", *args],
            cwd=BASE_DIR,
            text=True,
            capture_output=True,
        )
        output = process.stdout
        if process.stderr:
            output += "\nSTDERR:\n" + process.stderr
        self.after(0, self._command_finished, args, process.returncode, output)

    def _command_finished(self, args: list[str], returncode: int, output: str) -> None:
        self.output.insert(tk.END, output)
        self.output.see(tk.END)
        self.status_var.set(f"Finished: {' '.join(args)} | exit code {returncode}")
        self.refresh_status()

    def open_file(self, path: Path) -> None:
        if not path.exists():
            messagebox.showwarning("File not found", str(path))
            return
        os.startfile(path)


def main() -> None:
    app = OutreachApp()
    app.mainloop()


if __name__ == "__main__":
    main()
