from __future__ import annotations

import re

import config
from outreach.models import EmailMessage, EmailTemplate, Stakeholder
from outreach.validator import unresolved_placeholders


PLACEHOLDER_RE = re.compile(r"\{([^{}\n]+)\}")


def render_email(
    subject: Stakeholder,
    templates: dict[str, EmailTemplate],
) -> EmailMessage:
    template_type = _select_template_type(subject.template_type, templates)
    template = templates[template_type]

    values = {
        "Greeting": _build_greeting(subject),
        "Personalisation Paragraph": "",
        "Relevant Sector / Field": _sector_text(subject.sector),
        "Organisation Routing Sentence": _organisation_routing_sentence(subject),
    }
    values.update(config.TEMPLATE_VALUES)

    subject_text = _render(template.subject_template, values)
    subject_text = _prefix_subject(subject_text)
    body = _render(template.body_template, values)
    body = _collapse_empty_paragraphs(body)

    unresolved = unresolved_placeholders(f"{subject_text}\n{body}")
    if unresolved:
        raise ValueError(f"unresolved template placeholder(s): {', '.join(unresolved)}")

    return EmailMessage(
        stakeholder_id=subject.test_id,
        target_name=subject.target_name,
        recipient=subject.recipient_email,
        recipient_type=subject.recipient_type,
        template_type=template_type,
        language=subject.language,
        subject=subject_text,
        body=body,
    )


def _select_template_type(
    raw_template_type: str,
    templates: dict[str, EmailTemplate],
) -> str:
    if raw_template_type in config.ALLOWED_TEMPLATE_TYPES and raw_template_type in templates:
        return raw_template_type
    if config.FALLBACK_TEMPLATE_TYPE not in templates:
        raise ValueError(f"fallback template missing: {config.FALLBACK_TEMPLATE_TYPE}")
    return config.FALLBACK_TEMPLATE_TYPE


def _build_greeting(subject: Stakeholder) -> str:
    if subject.recipient_type == "Individual":
        name = subject.salutation or subject.target_name
        return f"Dear {name},"
    if subject.recipient_type == "Organisation":
        return f"Dear {subject.organisation} team,"
    raise ValueError(f"invalid Recipient Type: {subject.recipient_type or '<blank>'}")


def _organisation_routing_sentence(subject: Stakeholder) -> str:
    if subject.recipient_type != "Organisation":
        return ""
    return (
        "We would appreciate it if this invitation could be shared with the "
        "colleague in your organisation best placed to provide input on this topic."
    )


def _sector_text(sector: str) -> str:
    if not sector or sector.strip().lower() == "other":
        return "your professional context"
    return sector


def _render(template: str, values: dict[str, str]) -> str:
    return PLACEHOLDER_RE.sub(lambda match: values.get(match.group(1), match.group(0)), template)


def _prefix_subject(subject: str) -> str:
    subject = subject.strip()
    if not config.TEST_SUBJECT_PREFIX:
        return subject
    if subject.startswith(config.TEST_SUBJECT_PREFIX):
        return subject
    return f"{config.TEST_SUBJECT_PREFIX} {subject}"


def _collapse_empty_paragraphs(text: str) -> str:
    normalised = text.replace("\r\n", "\n").replace("\r", "\n")
    normalised = "\n".join(line.rstrip() for line in normalised.splitlines())
    normalised = re.sub(r"\n{3,}", "\n\n", normalised)
    return normalised.strip()
