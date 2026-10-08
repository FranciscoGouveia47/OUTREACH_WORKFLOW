from __future__ import annotations


import config


DEFAULT_CONTACT_FIT_HANDOFF_RULES = """Contact-fit and personalisation guardrails:
- Before drafting, perform a local relevance check using only the row context.
- Do not assume the contact is suitable merely because their organisation,
  department, role title, sector label, or company category appears adjacent to
  the campaign topic.
- The email may only make claims directly supported by the individual's own
  role, responsibility, profile, public materials, or recorded evidence note.
- If the row does not contain direct individual or role-based evidence that the
  recipient is a good fit, do not produce a ready email. Set AI Ready For Human
  Review to No and add AI Risk Flags such as weak individual evidence,
  organisation-level evidence, wrong-person risk, overbroad personalisation, or
  routing-only contact.
- If the person appears to be a departmental, organisational, or administrative
  contact but their own relevance is unclear, either mark the row not ready or
  draft only a routing/referral request asking whether they can direct the
  message to the appropriate colleague.
- Never create personalisation by concatenating workbook fields, for example
  "Given your role as [Role], together with your experience in [Sector]..." or
  "we value your perspective on [list of topics]".
- Do not explain why the contact was selected or describe the evidence source
  to the recipient. Never write source/provenance narration such as "the row
  says", "the workbook lists", "the source describes", or "for that reason,
  your perspective would be valuable".
- Avoid generic fit/rationale formulas such as "your work aligns closely with
  this campaign" when they only restate that the recipient matches the target
  list. Instead, make the sentence substantive and narrow.
- Do not mix third-person evidence text with second-person address. Phrases
  like "his expertise", "her role", or "their profile" should not appear in an
  email addressed directly to that person unless they are part of an unavoidable
  quoted title.
- Avoid list-like topic strings. Prefer one or two specific, supported reasons
  for relevance over broad lists.
- Use normal sentence case for workbook-derived role, sector, and topic phrases
  unless the capitalisation is part of an official title, acronym, campaign
  name, organisation name, or other proper noun.
"""


CONTACT_FIT_HANDOFF_RULES = (
    config.read_text_file(config.PROMPT_RULES_PATH)
    or DEFAULT_CONTACT_FIT_HANDOFF_RULES
)
