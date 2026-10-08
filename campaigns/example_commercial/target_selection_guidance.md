# Example Commercial Outreach Target Selection Guidance

This placeholder guidance shows how to document target-selection rules for a
commercial campaign. Replace it before using the workflow with real prospects.

## Core Rule

Select a target only when there is direct evidence that the person or mailbox is
relevant to the commercial offer. Do not infer relevance from a company category
alone.

## Acceptable Evidence

- A public profile showing that the person owns the relevant business function.
- A company page naming the person as responsible for partnerships,
  procurement, operations, product, growth, sustainability, IT, or another
  offer-relevant area.
- A job title that directly maps to the campaign objective, when the objective
  is also recorded in the evidence note.
- A team/shared inbox that is explicitly meant for sales, partnerships,
  procurement, supplier onboarding, press, or another relevant routing purpose.
- A recent company announcement that ties the target organisation to the
  problem your offer solves, with careful wording that does not overstate the
  individual's role.

## Not Enough By Itself

- The company is in the right sector, but the named person has no supported
  responsibility.
- A senior title such as CEO, founder, director, or manager without evidence of
  relevance to the offer.
- A generic contact page that does not indicate who handles the relevant topic.
- A LinkedIn headline or database category that sounds promising but does not
  support a concrete business reason for outreach.
- A broad phrase such as innovation, digital transformation, sustainability, or
  growth when the offer is narrower than that phrase.

## Minimum Row Data

Each target row should include:

- Test ID.
- Target Name.
- Organisation.
- Role / Position.
- Recipient Type: Individual or Organisation.
- Language.
- Template Type.
- Professional Context.
- Relevant Sector / Field.
- Personalisation Fact.
- Test Recipient Email.
- Enabled.

The personalisation fact should be written as a short, evidence-backed note. It
is not recipient-facing copy; it is input for the AI handoff and human reviewer.

## Review Rule

Before marking an imported row `Ready To Send = Yes`, check that:

- the person or shared mailbox is a defensible target,
- the email makes no unsupported claims,
- the locked campaign links and sender details are present,
- the message contains no real recipient email address,
- the message does not imply a prior relationship,
- the tone is appropriate for a first commercial contact,
- any weak fit is flagged for manual review.

## Example Prompt For Future Target Research

```text
Find potential commercial outreach targets for [offer/category].

Use a strict individual-level evidence gate. Approve a person only when a public
source supports that they own or influence a relevant business problem,
procurement area, partnership area, or operational function. Do not approve a
contact from company sector alone.

Return a table with: Target Name, Organisation, Role / Position, Recipient
Type, Language, Template Type, Evidence Source URL, Evidence Note,
Personalisation Fact, Risk Flags, Approve For Outreach, Reason If Not Approved.
```
