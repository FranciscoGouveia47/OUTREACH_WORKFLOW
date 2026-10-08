from __future__ import annotations

import unittest
from dataclasses import replace

from outreach.excel_reader import read_workbook
from outreach.template_engine import render_email

import config


class TemplateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.subjects, cls.templates = read_workbook(config.EXCEL_PATH)

    def test_salutation_does_not_duplicate_dear(self) -> None:
        subject = next(
            item
            for item in self.subjects
            if item.enabled and item.recipient_type == "Individual" and item.salutation
        )

        email = render_email(subject, self.templates)

        self.assertIn(f"Dear {subject.salutation},", email.body)
        self.assertNotIn("Dear Dear", email.body)

    def test_other_sector_is_not_inserted_literally(self) -> None:
        subject = replace(next(item for item in self.subjects if item.enabled), sector="Other")

        email = render_email(subject, self.templates)

        self.assertNotIn("activity in Other", email.body)
        self.assertIn("your professional context", email.body)

    def test_no_unresolved_placeholders(self) -> None:
        for subject in self.subjects:
            if subject.enabled:
                email = render_email(subject, self.templates)
                self.assertNotRegex(email.subject + email.body, r"\{[^{}]+\}")


if __name__ == "__main__":
    unittest.main()
