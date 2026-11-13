---
status: accepted
---

# Ground truth first, documents written backwards from it

Each Case starts as facts in `case.yaml` (customer, object, scope, deliberate flaws). A script computes the correct Offer from the Company world rates, catalogue and company rules, and only then are the request side, operations side and Offer documents written to match. Writing documents first and pricing them afterwards would read more naturally, but the correct Offer could then neither be guaranteed nor checked, and flaws would appear by accident rather than by design.

## Considered Options

- Free generation of documents, priced afterwards: rejected, no verifiable ground truth.
- Ground truth first (chosen): every total and every flaw is known and listed.
