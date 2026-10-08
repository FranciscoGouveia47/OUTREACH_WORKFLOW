from __future__ import annotations

import config
from outreach.handoff_rules import CONTACT_FIT_HANDOFF_RULES
from outreach.models import EmailMessage, Stakeholder


def build_prompt_packet(stakeholder: Stakeholder, base_email: EmailMessage) -> str:
    return f"""You are preparing one professional outreach email for this campaign: {config.CAMPAIGN_NAME}.

Prompt version: {config.PROMPT_PACKET_VERSION}

Task:
- Use the base email as the starting point.
- Customize and improve it using only the stakeholder context below.
- You may substantially rewrite the email if needed for tone, clarity, relevance, translation quality, or consistency.
- Do not invent facts, credentials, publications, or expertise.
- Treat the stakeholder context below as the only source available in this handoff; do not add new facts unless externally verified.
- If you add any stakeholder-specific fact, role, publication, credential, or expertise beyond what is already in the stakeholder context, verify it externally first. If external checking is not available, do not add new facts; instead rewrite using only the supplied context and flag any uncertainty in AI Quality Notes or AI Risk Flags.
- Do not include the real recipient email address.
- Keep the message concise, professional, and suitable for human review.
- Vary the email structure and phrasing naturally across rows so messages do not read like mechanically repeated bulk mail. Use only truthful, stakeholder-relevant variation, such as changing sentence order, combining or splitting short paragraphs, and choosing different natural transitions. Do not use deceptive tactics, hidden text, obfuscation, misleading subjects, or any wording whose only purpose is to evade mail filtering.
- Use the requested language. If the requested language is Portuguese, use European Portuguese; if it is Portuguese (Brazil), use Brazilian Portuguese.
- Keep the locked values unchanged.
- Keep organisation/shared-inbox language careful: do not imply the mailbox owner is the decision maker or end recipient.
- Do not mechanically paste the Organisation field into phrases such as "your work at [Organisation]". Many organisation values include departments, product lines, regions, slashes, commas, or combined affiliations that sound unnatural if inserted verbatim. Rewrite affiliation details into natural professional prose before using them. Examples:
  - Bad: "Your work at Acme Group, Procurement / EMEA Supplier Operations..."
    Better: "Given your work on supplier operations across EMEA..."
  - Bad: "Your work at Northstar Operations Ltd, Manufacturing / Supplier onboarding..."
    Better: "Given your responsibility for supplier onboarding in a manufacturing context..."
  - Bad: "Your work at BrightPath Software / Partner ecosystem / Integrations..."
    Better: "Given your team's focus on partner integrations..."
  When unsure how to phrase a complex affiliation, use the role, field, or personalisation fact instead of forcing the raw organisation string into the sentence.

{CONTACT_FIT_HANDOFF_RULES}

Localisation rules:
- Formality preservation: maintain the same professional and formal register as the base template in every target language.
- Greeting localisation: preserve the function and formality of the source greeting. Use a natural professional salutation appropriate to the recipient type, target language, and country. Do not downgrade a formal greeting to a casual one such as "Hello", "Hi", "Olá", or equivalents unless the base template itself is informal.
- Natural localisation: treat the source template as a definition of meaning, facts, structure, and tone, not as sentence-level syntax. Rewrite each sentence naturally in the requested language rather than translating sentence by sentence.
- Localise, do not phrase-map: examples in instructions are guidance only. Do not apply fixed phrase replacements mechanically; choose wording a native professional speaker would normally use.
- Recipient classification: before drafting, classify the target as an individual recipient, an organisation/team/shared inbox, or uncertain. Use the stakeholder name, role, organisation, recipient type, and email context without inventing facts.
- Recipient consistency is mandatory: maintain the selected recipient perspective throughout the greeting, pronouns, verb forms, possessives, requests, routing sentence, opt-out wording, and closing. Do not address a team collectively in the greeting and later switch to singular-person wording.
- For organisation, team, department, or shared-inbox recipients, use plural, collective, or naturally impersonal wording throughout. If recipient type is genuinely uncertain, prefer neutral or impersonal wording that avoids assuming one specific individual.
- Regional language variant: use the requested regional variant and professional conventions for that country. For Portuguese, use European Portuguese, not Brazilian Portuguese.
- Pay close attention to capital letters that appear unnecessarily in the middle of sentences. Use normal sentence case unless the capitalisation is required for a proper noun, official project name, acronym, title, grant label, URL, or other locked value.
- Native-language QA: before setting AI Ready For Human Review to Yes, reread the entire email as if it had originally been written by a native professional speaker in the target language. If any sentence sounds translated, unnatural, overly literal, overly bureaucratic, or inconsistent in recipient perspective, rewrite it before completion.

Return the result as five fields:
AI Subject:
AI Email Body:
AI Quality Notes:
AI Ready For Human Review: Yes or No
AI Risk Flags: comma-separated risk flags, or None

Target ID: {stakeholder.test_id}
Target Name: {stakeholder.target_name}
Salutation: {stakeholder.salutation}
Organisation: {stakeholder.organisation}
Role / Position: {stakeholder.role}
Professional Context: {stakeholder.professional_context}
Relevant Sector / Field: {stakeholder.sector}
Personalisation Fact: {stakeholder.personalisation_fact}
Language: {stakeholder.language}
Recipient Type: {stakeholder.recipient_type}
Template Type: {base_email.template_type}

Locked values:
Subject prefix, if any: {config.TEST_SUBJECT_PREFIX or "<none>"}
Campaign name: {config.CAMPAIGN_NAME}
Campaign URL: {config.CAMPAIGN_URL}
Call-to-action label: {config.CTA_LABEL}
Call-to-action URL: {config.CTA_URL}
Company website: {config.COMPANY_URL}
Reference code: {config.REFERENCE_CODE}
Company LinkedIn: {config.COMPANY_LINKEDIN_URL}
Sender name: {config.SENDER_NAME}
Sender position: {config.SENDER_POSITION}
Sender organisation: {config.SENDER_ORGANISATION}

Base subject:
{base_email.subject}

Base email:
{base_email.body}
"""
