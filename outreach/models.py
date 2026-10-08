from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Stakeholder:
    source_row: int
    test_id: str
    target_name: str
    salutation: str
    organisation: str
    role: str
    professional_context: str
    sector: str
    personalisation_fact: str
    language: str
    recipient_email: str
    enabled: bool
    recipient_type: str
    template_type: str
    handoff_generated: bool


@dataclass(frozen=True)
class EmailTemplate:
    template_type: str
    typical_audience: str
    subject_template: str
    body_template: str
    usage_notes: str


@dataclass(frozen=True)
class EmailMessage:
    stakeholder_id: str
    target_name: str
    recipient: str
    recipient_type: str
    template_type: str
    language: str
    subject: str
    body: str


@dataclass(frozen=True)
class HandoffEmail:
    stakeholder: Stakeholder
    email: EmailMessage
    prompt_packet: str


@dataclass(frozen=True)
class ReviewEmail:
    stakeholder: Stakeholder
    subject: str
    body: str
    ai_ready_for_human_review: str
    ai_quality_notes: str
    ai_risk_flags: str
    prompt_version: str
    validation_status: str
    validation_errors: str


@dataclass(frozen=True)
class QueueEmail:
    test_id: str
    target_name: str
    recipient_email: str
    recipient_type: str
    template_type: str
    language: str
    subject: str
    body: str
    ready_to_send: bool
    validation_status: str
    scheduled_send_at: str = ""
    ai_risk_flags: str = ""
    prompt_version: str = ""
