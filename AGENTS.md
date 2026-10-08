# AGENTS.md

Guidance for coding agents working in this repository.

## Project Overview

This is a configurable local Python workflow for controlled outreach campaigns.
It reads a campaign source workbook, generates deterministic base emails and
ChatGPT handoff workbooks, imports completed AI-reviewed workbook output into a
local review queue, and supports controlled Outlook or Microsoft Graph delivery.

Email must only be sent through explicit delivery commands or UI actions.
Validation, preview, handoff generation, imports, migrations, preflight checks,
and tests must not send email.

## Campaign Separation

Keep campaign-specific material out of reusable Python modules.

Campaign-specific material belongs in:

- `campaigns/<campaign>/campaign.json`
- `campaigns/<campaign>/target_selection_guidance.md`
- `campaigns/<campaign>/prompt_rules.md`
- the configured source workbook

The default campaign is `campaigns/example_commercial`. The active campaign can
be changed with `OUTREACH_CAMPAIGN_DIR`.

## Repository Layout

- `main.py`: CLI entry point. Running without arguments opens the Tkinter UI.
- `ui.py`: desktop control UI and single-row or batch workflow actions.
- `config.py`: loads campaign config and exposes runtime constants.
- `outreach/`: reusable workflow logic:
  - `models.py`: frozen dataclasses for stakeholders, emails, review rows, and queue rows.
  - `excel_reader.py`: reads source workbook subjects and templates.
  - `template_engine.py`: renders deterministic base email text.
  - `prompt_packet.py`: builds ChatGPT prompt packets.
  - `handoff_workbook.py`: creates and updates ChatGPT handoff workbooks.
  - `handoff_rules.py`: loads campaign prompt guardrails.
  - `review_importer.py`: imports completed handoff workbooks and validates AI output.
  - `review_queue.py`: reads and mutates the local review queue workbook.
  - `status.py`: builds per-target UI/status rows from source, handoff, review, and log state.
  - `service.py`: pipeline state, runtime file migration, and delivery preflight checks.
  - `scheduler.py`, `outlook_client.py`, `graph_client.py`: delivery backends and send queue processing.
  - `logger.py`: append-oriented send-log helpers and status readers.
  - `workbook_state.py`, `workbook_schema.py`, `lineage.py`: workbook metadata, backups, migrations, and lineage checks.
- `tests/`: `unittest` tests.
- `data/`, `review/`, `logs/`, and `previews/`: runtime workbooks, logs, review files, and previews.

## Runtime Data Rules

Treat runtime files as user/project data, not source code.

- The source workbook configured by `config.EXCEL_PATH` is input data unless the
  user explicitly asks for a data edit.
- The ChatGPT handoff workbook intentionally excludes real recipient email
  addresses.
- The local review queue reattaches real recipient email addresses by `Test ID`.
- `logs/send_log.csv` is an append-oriented delivery audit trail; do not rewrite
  or truncate it casually.
- Preserve workbook lineage metadata when creating or importing handoff/review
  workbooks.
- Use existing backup helpers before code changes that intentionally mutate
  workbook data.
- Do not commit or churn generated Excel, CSV, preview, cache, or `.venv` files.
- The checked-in `data/commercial_outreach_targets.xlsx` is synthetic example
  data; real campaign workbooks should normally remain untracked.

## Common Commands

Use the project virtual environment on Windows:

```powershell
.\.venv\Scripts\python.exe main.py
.\.venv\Scripts\python.exe main.py validate
.\.venv\Scripts\python.exe main.py preview
.\.venv\Scripts\python.exe main.py generate-handoff
.\.venv\Scripts\python.exe main.py import-review
.\.venv\Scripts\python.exe main.py preflight --preflight-mode draft
.\.venv\Scripts\python.exe main.py preflight --preflight-mode schedule_outlook
.\.venv\Scripts\python.exe main.py preflight --preflight-mode send_now
.\.venv\Scripts\python.exe main.py migrate-files
.\.venv\Scripts\python.exe -m unittest
```

Delivery-related commands must be handled carefully:

```powershell
.\.venv\Scripts\python.exe main.py send-ready
.\.venv\Scripts\python.exe main.py create-outlook-drafts --max-daily 1
.\.venv\Scripts\python.exe main.py send-ready --delivery outlook --confirm-send --max-daily 1
```

Do not run commands with `--confirm-send`, Outlook scheduling, or Graph delivery
unless the user explicitly requests real delivery behavior.

## Development Conventions

- Keep reusable workflow behavior independent from one named campaign.
- Prefer adding configurable values to `campaign.json` over hard-coding a new
  campaign rule in Python.
- Keep changes small and aligned with existing module boundaries.
- Keep data records as frozen dataclasses where the surrounding code does so.
- Use `pathlib.Path` for filesystem paths.
- Use `openpyxl` APIs for workbook reads/writes.
- Preserve dry-run and preflight behavior when touching delivery code.
- Delivery eligibility logic belongs in `OutreachService.delivery_preflight` or
  the scheduler/status layer, not only in UI handlers.
- Use the existing status vocabulary where possible: `SENT`, `DRAFT_CREATED`,
  `DRAFT_FAILED`, `OUTLOOK_SCHEDULED`, `OUTLOOK_SCHEDULE_CANCELLED`,
  `OUTLOOK_OUTBOX_LOCKED`, `SEND_STATUS_UNKNOWN`, `DELIVERY_FAILED`,
  `DELIVERY_FAILURE_CLEARED`, `REFUSED_TO_ANSWER`,
  `REFUSED_TO_ANSWER_CLEARED`, `VALID`, and `INVALID`.
- Keep UI worker actions non-blocking by following the existing thread/`after`
  pattern in `ui.py`.
- Use ASCII for source changes unless editing existing non-ASCII content
  intentionally.

## Testing Expectations

Run the full suite after behavioral changes:

```powershell
.\.venv\Scripts\python.exe -m unittest
```

Add or update focused `unittest` tests when changing:

- workbook schemas or workbook import/export behavior,
- prompt packet or template rendering rules,
- validation and risk flag logic,
- status/next-action calculation,
- preflight or delivery-blocking behavior,
- UI behavior that can be covered with mocks/fakes.

## Safety Notes

- Do not expose real recipient email addresses in ChatGPT handoff workbook
  fields.
- Do not weaken validation that blocks recipient email leaks, missing locked
  values, unresolved placeholders, or unsupported localized wording.
- Do not weaken prompt-version checks or campaign contact-fit guardrails.
- Do not silently bypass file lineage checks before delivery.
- Do not launch Outlook invisibly or assume Outlook is configured.
- Microsoft Graph sending depends on `MS_GRAPH_CLIENT_ID` and optionally
  `MS_GRAPH_TENANT_ID`; do not hard-code credentials.
- If the worktree has existing edits, assume they are user changes and preserve
  them.
