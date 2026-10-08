from __future__ import annotations

from pathlib import Path

from outreach.models import EmailMessage


def clear_previews(previews_dir: Path) -> None:
    previews_dir.mkdir(parents=True, exist_ok=True)
    for preview_file in previews_dir.glob("*.txt"):
        preview_file.unlink()


def save_preview(email: EmailMessage, previews_dir: Path) -> Path:
    previews_dir.mkdir(parents=True, exist_ok=True)
    path = previews_dir / f"{email.stakeholder_id}.txt"
    path.write_text(format_preview(email), encoding="utf-8")
    return path


def format_preview(email: EmailMessage) -> str:
    return (
        "=" * 60
        + "\n"
        + f"{email.stakeholder_id}\n"
        + f"Actual destination:\n{email.recipient}\n"
        + f"Role-play stakeholder:\n{email.target_name}\n"
        + f"Recipient Type:\n{email.recipient_type}\n"
        + f"Template:\n{email.template_type}\n"
        + f"Language:\n{email.language}\n"
        + f"Subject:\n{email.subject}\n"
        + "BODY:\n"
        + f"{email.body}\n"
        + "=" * 60
        + "\n"
    )
