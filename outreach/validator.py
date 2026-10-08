from __future__ import annotations

import re
from collections import Counter

import config
from outreach.models import EmailTemplate, Stakeholder


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
UNRESOLVED_PLACEHOLDER_RE = re.compile(r"\{[^{}\n]+\}")


def validate_templates(templates: dict[str, EmailTemplate]) -> list[str]:
    errors: list[str] = []

    if config.FALLBACK_TEMPLATE_TYPE not in templates:
        errors.append(
            f"Missing required fallback template: {config.FALLBACK_TEMPLATE_TYPE}"
        )

    for name, template in templates.items():
        if not template.subject_template:
            errors.append(f"{name}: missing subject template")
        if not template.body_template:
            errors.append(f"{name}: missing body template")

    return errors


def validate_unique_ids(subjects: list[Stakeholder]) -> list[str]:
    ids = [subject.test_id for subject in subjects if subject.test_id]
    duplicate_ids = sorted(test_id for test_id, count in Counter(ids).items() if count > 1)
    return [f"Duplicate Test ID: {test_id}" for test_id in duplicate_ids]


def validate_stakeholder(subject: Stakeholder) -> list[str]:
    errors: list[str] = []

    if not subject.test_id:
        errors.append("missing Test ID")
    if not subject.recipient_email:
        errors.append("missing Test Recipient Email")
    elif not EMAIL_RE.match(subject.recipient_email):
        errors.append(f"invalid Test Recipient Email: {subject.recipient_email}")

    if subject.recipient_type not in config.ALLOWED_RECIPIENT_TYPES:
        errors.append(f"invalid Recipient Type: {subject.recipient_type or '<blank>'}")

    if subject.recipient_type == "Individual":
        if not subject.salutation and not subject.target_name:
            errors.append("Individual recipient requires Salutation or Target Name")

    if subject.recipient_type == "Organisation" and not subject.organisation:
        errors.append("Organisation recipient requires Organisation")

    if subject.language not in config.SUPPORTED_LANGUAGES:
        errors.append(f"unsupported Language: {subject.language or '<blank>'}")

    return errors


def unresolved_placeholders(text: str) -> list[str]:
    return sorted(set(UNRESOLVED_PLACEHOLDER_RE.findall(text)))
