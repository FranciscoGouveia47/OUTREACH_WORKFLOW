# Configurable Outreach Workflow

This repository is a generic version of the local outreach workflow originally
built for a single campaign. It reads a campaign workbook, generates
deterministic base emails, creates a ChatGPT handoff workbook, imports completed
AI-reviewed output into a local review queue, and supports controlled Outlook or
Microsoft Graph delivery.

The included `example_commercial` campaign uses placeholder data only. Replace
it before running real outreach.

## Safety Model

Email is sent only by explicit delivery commands or UI actions.

These actions do not send email:

- validation,
- preview generation,
- ChatGPT handoff generation,
- completed handoff import,
- migration,
- preflight,
- tests.

Real delivery requires one of:

- `main.py create-outlook-drafts`,
- `main.py send-ready --delivery outlook --confirm-send`,
- `main.py send-ready --confirm-send` for Microsoft Graph,
- the matching buttons in the desktop UI.

Do not run delivery commands against real recipients until you have reviewed the
local review queue and understand the send log state.

## Repository Layout

- `main.py`: CLI entry point. Running without arguments opens the Tkinter UI.
- `ui.py`: desktop control panel for validation, handoff, review, scheduling,
  Outlook draft/send actions, and status review.
- `config.py`: loads the active campaign from `campaign.json` and exposes the
  paths/settings used by the workflow.
- `outreach/`: reusable workflow logic.
- `campaigns/example_commercial/`: placeholder campaign config, prompt rules,
  and target-selection guidance.
- `data/commercial_outreach_targets.xlsx`: synthetic example source workbook.
- `review/`: generated ChatGPT handoff and local review queue workbooks.
- `logs/`: append-only delivery audit log.
- `previews/`: deterministic base email previews.
- `tests/`: `unittest` coverage for workbook, queue, status, scheduling, and
  delivery-blocking behavior.

## Campaign Configuration

The active campaign is a directory containing at least:

```text
campaign.json
target_selection_guidance.md
prompt_rules.md
```

By default the app uses:

```text
campaigns/example_commercial
```

To use a different campaign, set:

```powershell
$env:OUTREACH_CAMPAIGN_DIR = "C:\path\to\your\campaign"
```

The campaign JSON controls:

- campaign name and description,
- prompt packet version,
- source workbook path,
- generated handoff/review/log paths,
- sender identity,
- locked URLs and reference codes,
- allowed recipient types,
- allowed template types,
- supported languages,
- Outlook folder names,
- send window, interval, daily cap, and timeout.

Keep project-specific facts in campaign files, not in Python modules.

## Source Workbook Schema

The source workbook needs two sheets.

### `Test Subjects`

Required columns:

```text
Test ID
Target Name
Salutation
Organisation
Role / Position
Professional Context
Relevant Sector / Field
Personalisation Fact
Language
Test Recipient Email
Enabled
Recipient Type
Template Type
```

Important rules:

- `Test ID` must be unique.
- `Enabled` must be `Yes` for a row to enter the handoff.
- `Recipient Type` must match `allowed_recipient_types`.
- `Template Type` should match `allowed_template_types`; otherwise the fallback
  template is used.
- `Personalisation Fact` should be an evidence-backed note, not polished email
  copy.
- Real recipient emails stay in the source workbook and local review queue.
  They are not exported into the ChatGPT handoff workbook.

### `Email Templates`

Required columns:

```text
Template Type
Typical Audience
Subject Template
Body Template
Usage Notes
```

Templates can use these common placeholders:

```text
{Greeting}
{Personalisation Paragraph}
{Relevant Sector / Field}
{Organisation Routing Sentence}
{Campaign Name}
{Campaign Description}
{Campaign URL}
{CTA Label}
{CTA URL}
{Survey Link}
{Company Website}
{Project Website}
{Reference Code}
{Project Grant Agreement}
{Company LinkedIn}
{Project LinkedIn}
{Sender Name}
{Sender Position}
{Sender Organisation}
```

`Survey Link`, `Project Website`, `Project Grant Agreement`, and
`Project LinkedIn` are retained as compatibility aliases. For a commercial
campaign they mean CTA URL, company website, reference code, and company
LinkedIn.

## Workflow

1. Edit or replace the campaign config and source workbook.
2. Run validation.
3. Generate the ChatGPT handoff workbook.
4. Upload the handoff workbook to ChatGPT and ask it to fill the AI output
   columns.
5. Import the completed workbook.
6. Review every row in the local review queue.
7. Mark approved rows `Ready To Send = Yes`.
8. Create Outlook drafts, schedule in Outlook, or send explicitly.

## Commands

Use a virtual environment with the dependencies in `requirements.txt`.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Open the desktop UI:

```powershell
.\.venv\Scripts\python.exe main.py
```

Validate the source workbook:

```powershell
.\.venv\Scripts\python.exe main.py validate
```

Generate deterministic previews:

```powershell
.\.venv\Scripts\python.exe main.py preview
```

Generate the ChatGPT handoff workbook:

```powershell
.\.venv\Scripts\python.exe main.py generate-handoff
```

Import a completed ChatGPT workbook:

```powershell
.\.venv\Scripts\python.exe main.py import-review
```

Run delivery preflight without touching Outlook or Graph:

```powershell
.\.venv\Scripts\python.exe main.py preflight --preflight-mode draft
.\.venv\Scripts\python.exe main.py preflight --preflight-mode schedule_outlook
.\.venv\Scripts\python.exe main.py preflight --preflight-mode send_now
```

Create Outlook drafts:

```powershell
.\.venv\Scripts\python.exe main.py create-outlook-drafts --max-daily 1
```

Send via classic Outlook:

```powershell
.\.venv\Scripts\python.exe main.py send-ready --delivery outlook --confirm-send --max-daily 1
```

Send via Microsoft Graph:

```powershell
$env:MS_GRAPH_CLIENT_ID = "your-entra-app-client-id"
.\.venv\Scripts\python.exe main.py send-ready --confirm-send --max-daily 1
```

Run tests:

```powershell
.\.venv\Scripts\python.exe -m unittest
```

## ChatGPT Handoff Instructions

Upload `review/chatgpt_handoff.xlsx` and ask:

```text
Fill the AI output columns for each row:
- AI Subject
- AI Email Body
- AI Quality Notes
- AI Ready For Human Review
- AI Risk Flags

Use each row's ChatGPT Prompt Packet as the controlling instruction.
Return an updated Excel workbook with the same columns.
```

## Validation And Delivery Blocks

Rows are blocked from delivery when:

- recipient email is missing,
- the source row is disabled,
- no handoff row exists,
- no completed review has been imported,
- `Ready To Send` is not `Yes`,
- validation status is not `VALID`,
- prompt version is stale,
- the row is already logged as `SENT`,
- the row is already logged as `OUTLOOK_SCHEDULED`,
- the latest durable status is `SEND_STATUS_UNKNOWN`,
- the latest durable status is `OUTLOOK_OUTBOX_LOCKED`,
- the latest durable status is `DELIVERY_FAILED`,
- the latest durable status is `REFUSED_TO_ANSWER`,
- the review queue lineage does not match the current handoff workbook.

The send log is append-oriented. Do not edit it manually unless you know exactly
what operational state you are changing.

## Creating A Real Campaign

1. Copy `campaigns/example_commercial` to a new campaign directory.
2. Edit `campaign.json`.
3. Replace `target_selection_guidance.md`.
4. Replace `prompt_rules.md`.
5. Create a real source workbook using the same two-sheet schema.
6. Set `OUTREACH_CAMPAIGN_DIR` to the new directory.
7. Run `validate`, then `preview`, before generating a handoff.

Keep real prospect data out of Git unless your organisation explicitly approves
storing it there.
