from __future__ import annotations

import unittest
from dataclasses import replace

from outreach.excel_reader import read_workbook
from outreach.validator import validate_stakeholder, validate_templates, validate_unique_ids

import config


class ValidationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.subjects, cls.templates = read_workbook(config.EXCEL_PATH)

    def test_workbook_has_required_fallback_template(self) -> None:
        self.assertEqual([], validate_templates(self.templates))

    def test_workbook_test_ids_are_unique(self) -> None:
        self.assertEqual([], validate_unique_ids(self.subjects))

    def test_invalid_recipient_type_blocks_row(self) -> None:
        subject = replace(self.subjects[0], recipient_type="Person")

        errors = validate_stakeholder(subject)

        self.assertIn("invalid Recipient Type: Person", errors)


if __name__ == "__main__":
    unittest.main()
