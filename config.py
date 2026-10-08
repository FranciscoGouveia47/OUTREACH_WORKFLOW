from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
PREVIEWS_DIR = BASE_DIR / "previews"
LOGS_DIR = BASE_DIR / "logs"
REVIEW_DIR = BASE_DIR / "review"
CAMPAIGNS_DIR = BASE_DIR / "campaigns"

DEFAULT_CAMPAIGN_DIR = CAMPAIGNS_DIR / "example_commercial"
ACTIVE_CAMPAIGN_DIR = Path(os.getenv("OUTREACH_CAMPAIGN_DIR", DEFAULT_CAMPAIGN_DIR)).resolve()
CAMPAIGN_CONFIG_PATH = ACTIVE_CAMPAIGN_DIR / "campaign.json"


def _load_campaign(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(
            f"Campaign config not found: {path}. "
            "Set OUTREACH_CAMPAIGN_DIR or create campaigns/example_commercial/campaign.json."
        )
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError(f"Campaign config must be a JSON object: {path}")
    return data


def _path(value: str, default: Path) -> Path:
    if not value:
        return default
    raw = Path(value)
    if raw.is_absolute():
        return raw
    return (BASE_DIR / raw).resolve()


def _campaign_path(value: str, default_name: str) -> Path:
    raw = Path(value or default_name)
    if raw.is_absolute():
        return raw
    return (ACTIVE_CAMPAIGN_DIR / raw).resolve()


def _string_set(values: Any) -> set[str]:
    if not values:
        return set()
    if not isinstance(values, list):
        raise ValueError("Expected a JSON list of strings.")
    return {str(value).strip() for value in values if str(value).strip()}


CAMPAIGN = _load_campaign(CAMPAIGN_CONFIG_PATH)

CAMPAIGN_NAME = str(CAMPAIGN.get("campaign_name", "Commercial Outreach")).strip()
CAMPAIGN_DESCRIPTION = str(CAMPAIGN.get("campaign_description", "")).strip()
PROMPT_PACKET_VERSION = str(CAMPAIGN.get("prompt_packet_version", "2026-10-08.1")).strip()

EXCEL_PATH = _path(
    str(CAMPAIGN.get("source_workbook", "data/commercial_outreach_targets.xlsx")),
    DATA_DIR / "commercial_outreach_targets.xlsx",
)
CHATGPT_HANDOFF_WORKBOOK_PATH = _path(
    str(CAMPAIGN.get("handoff_workbook", "review/chatgpt_handoff.xlsx")),
    REVIEW_DIR / "chatgpt_handoff.xlsx",
)
LOCAL_REVIEW_WORKBOOK_PATH = _path(
    str(CAMPAIGN.get("review_workbook", "review/email_review_queue.xlsx")),
    REVIEW_DIR / "email_review_queue.xlsx",
)
SEND_LOG_PATH = _path(
    str(CAMPAIGN.get("send_log", "logs/send_log.csv")),
    LOGS_DIR / "send_log.csv",
)

TARGET_SELECTION_GUIDANCE_PATH = _campaign_path(
    str(CAMPAIGN.get("target_selection_guidance", "target_selection_guidance.md")),
    "target_selection_guidance.md",
)
PROMPT_RULES_PATH = _campaign_path(
    str(CAMPAIGN.get("prompt_rules", "prompt_rules.md")),
    "prompt_rules.md",
)

TEST_MODE = bool(CAMPAIGN.get("test_mode", True))
DRY_RUN = bool(CAMPAIGN.get("dry_run", True))

CAMPAIGN_URL = str(CAMPAIGN.get("campaign_url", "")).strip()
CTA_LABEL = str(CAMPAIGN.get("cta_label", "campaign link")).strip()
CTA_URL = str(CAMPAIGN.get("cta_url", CAMPAIGN_URL)).strip()
COMPANY_URL = str(CAMPAIGN.get("company_url", "")).strip()
COMPANY_LINKEDIN_URL = str(CAMPAIGN.get("company_linkedin_url", "")).strip()
REFERENCE_CODE = str(CAMPAIGN.get("reference_code", "")).strip()
SENDER_NAME = str(CAMPAIGN.get("sender_name", "")).strip()
SENDER_POSITION = str(CAMPAIGN.get("sender_position", "")).strip()
SENDER_ORGANISATION = str(CAMPAIGN.get("sender_organisation", "")).strip()
TEST_SUBJECT_PREFIX = str(CAMPAIGN.get("subject_prefix", "")).strip()

# Backwards-compatible aliases for older template placeholders.
SURVEY_URL = CTA_URL
PROJECT_URL = COMPANY_URL
PROJECT_GRANT_AGREEMENT = REFERENCE_CODE
PROJECT_LINKEDIN_URL = COMPANY_LINKEDIN_URL

LOCKED_VALUES = [
    value
    for value in (
        CTA_URL,
        COMPANY_URL,
        REFERENCE_CODE,
        COMPANY_LINKEDIN_URL,
        SENDER_NAME,
    )
    if value
]

MAX_DAILY_EMAILS = int(CAMPAIGN.get("max_daily_emails", 30))
SEND_INTERVAL_MINUTES = int(CAMPAIGN.get("send_interval_minutes", 20))
SEND_WINDOW_START = str(CAMPAIGN.get("send_window_start", "09:00")).strip()
SEND_WINDOW_END = str(CAMPAIGN.get("send_window_end", "17:00")).strip()
OUTLOOK_TIMEOUT_SECONDS = int(CAMPAIGN.get("outlook_timeout_seconds", 60))
OUTLOOK_SENT_SEARCH_FOLDER_NAMES = tuple(
    str(value).strip()
    for value in CAMPAIGN.get("outlook_sent_search_folder_names", [])
    if str(value).strip()
)
DISABLED_TEMPLATE_REASONS = {
    str(key).strip(): str(value).strip()
    for key, value in CAMPAIGN.get("disabled_template_reasons", {}).items()
    if str(key).strip() and str(value).strip()
}
DISABLED_PERSONALISATION_KEYWORDS = {
    str(key).strip().lower(): str(value).strip()
    for key, value in CAMPAIGN.get("disabled_personalisation_keywords", {}).items()
    if str(key).strip() and str(value).strip()
}

ALLOWED_RECIPIENT_TYPES = _string_set(CAMPAIGN.get("allowed_recipient_types")) or {
    "Individual",
    "Organisation",
}
ALLOWED_TEMPLATE_TYPES = _string_set(CAMPAIGN.get("allowed_template_types")) or {
    "General / Fallback",
}
FALLBACK_TEMPLATE_TYPE = str(CAMPAIGN.get("fallback_template_type", "General / Fallback")).strip()
SUPPORTED_LANGUAGES = _string_set(CAMPAIGN.get("supported_languages")) or {"English"}

HANDOFF_GENERATED_COLUMN = "ChatGPT Handoff Generated"
HANDOFF_GENERATED_AT_COLUMN = "ChatGPT Handoff Generated At"
HANDOFF_WORKBOOK_COLUMN = "ChatGPT Handoff Workbook"
REVIEW_IMPORTED_COLUMN = "Review Imported"
REVIEW_IMPORTED_AT_COLUMN = "Review Imported At"
SCHEDULED_SEND_AT_COLUMN = "Scheduled Send At"

GRAPH_CLIENT_ID_ENV = "MS_GRAPH_CLIENT_ID"
GRAPH_TENANT_ID_ENV = "MS_GRAPH_TENANT_ID"
GRAPH_DEFAULT_TENANT = "organizations"
GRAPH_SCOPES = ["https://graph.microsoft.com/Mail.Send"]

TEMPLATE_VALUES = {
    "Campaign Name": CAMPAIGN_NAME,
    "Campaign Description": CAMPAIGN_DESCRIPTION,
    "Campaign URL": CAMPAIGN_URL,
    "CTA Label": CTA_LABEL,
    "CTA URL": CTA_URL,
    "Survey Link": SURVEY_URL,
    "Company Website": COMPANY_URL,
    "Project Website": PROJECT_URL,
    "Reference Code": REFERENCE_CODE,
    "Project Grant Agreement": PROJECT_GRANT_AGREEMENT,
    "Company LinkedIn": COMPANY_LINKEDIN_URL,
    "Project LinkedIn": PROJECT_LINKEDIN_URL,
    "Sender Name": SENDER_NAME,
    "Sender Position": SENDER_POSITION,
    "Sender Organisation": SENDER_ORGANISATION,
}


def read_text_file(path: Path) -> str:
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8").strip()
