# Configuration Guide

This project separates reusable outreach mechanics from campaign-specific
content. A campaign directory contains the data and rules for one outreach
project.

## Active Campaign

Default:

```text
campaigns/example_commercial
```

Override:

```powershell
$env:OUTREACH_CAMPAIGN_DIR = "C:\Users\you\campaigns\my_campaign"
```

The path can be inside or outside this repository.

## `campaign.json`

Required or commonly used fields:

| Field | Purpose |
| --- | --- |
| `campaign_name` | Display name used in the UI, console output, and prompt packets. |
| `campaign_description` | Short description available to templates and prompt packets. |
| `prompt_packet_version` | Increment whenever prompt rules materially change. Stale review rows are blocked from delivery. |
| `source_workbook` | Path to the input workbook. Relative paths are resolved from the repo root. |
| `handoff_workbook` | Path for the generated ChatGPT handoff workbook. |
| `review_workbook` | Path for the local review queue workbook. |
| `send_log` | Path for the append-oriented delivery log. |
| `target_selection_guidance` | Markdown guidance file in the campaign directory. |
| `prompt_rules` | Markdown prompt guardrails file in the campaign directory. |
| `campaign_url` | General campaign or offer URL. |
| `cta_label` | Human label for the primary call to action. |
| `cta_url` | Primary locked call-to-action URL. |
| `company_url` | Locked company website. |
| `company_linkedin_url` | Locked company LinkedIn URL. |
| `reference_code` | Optional internal reference, campaign code, grant code, or tracking label. |
| `sender_name` | Sender name used in templates and prompt packets. |
| `sender_position` | Sender role/title. |
| `sender_organisation` | Sender organisation. |
| `subject_prefix` | Optional prefix added to generated subjects, for example `[DRAFT]`. |
| `allowed_recipient_types` | Usually `Individual` and `Organisation`. |
| `allowed_template_types` | Template categories accepted from the source workbook. |
| `fallback_template_type` | Template used when a row's template type is not available. |
| `supported_languages` | Accepted language labels. |
| `outlook_sent_search_folder_names` | Extra Outlook folders searched when reconciling scheduled sends. |

## Workbook Data Flow

```text
Source workbook
  -> deterministic base email
  -> ChatGPT handoff workbook
  -> completed ChatGPT handoff workbook
  -> local review queue
  -> Outlook / Graph delivery
  -> append-only send log
```

The ChatGPT handoff workbook does not contain real recipient email addresses.
The local review queue reattaches them by `Test ID`.

## Prompt Rules

`prompt_rules.md` is inserted into every row's prompt packet and into the
handoff workbook's instruction sheet. Use it for:

- target-fit rules,
- personalisation limits,
- commercial claims that must not be made,
- tone and compliance expectations,
- shared-inbox or routing behavior,
- risk flags.

When you materially strengthen or change these rules, increment
`prompt_packet_version` in `campaign.json`. Existing imported review rows with
old prompt versions will then be blocked from delivery until regenerated and
reimported.

## Locked Values

The importer requires non-empty locked values to appear in the AI output:

- `cta_url`
- `company_url`
- `reference_code`
- `company_linkedin_url`
- `sender_name`

Leave optional locked values blank if they should not be required.

## Real Campaign Checklist

Before real delivery:

1. Replace placeholder sender and URL values.
2. Replace `target_selection_guidance.md`.
3. Replace `prompt_rules.md`.
4. Replace the source workbook.
5. Run `validate`.
6. Run `preview` and inspect generated base emails.
7. Generate a small handoff first with `--limit 1`.
8. Import the completed handoff.
9. Review the local queue manually.
10. Run preflight.
11. Create one Outlook draft before scheduling or sending anything.
