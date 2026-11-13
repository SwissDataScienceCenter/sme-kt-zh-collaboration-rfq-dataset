---
status: accepted
---

# Advisory Offers are estimated from past offers and graded by content and a fee band

For Custom build and System product Cases every Offer line can be traced to a written rule (price list, item master, costing manual, Faustregeln), and the expected total is exact. Advisory work is different: what will be delivered is only roughly known when the Offer is written, and Fronalp has no table of hours per task. Like real consultants, the advisor estimates by comparing with similar past offers and their final hours. A1 therefore gets no hour table. The basis is two past project folders (offer, invoice by actual hours, project sheet), the hours export and the base history of fee offers. The reference Offer in `case.yaml` is one sound answer; an Offer is accepted when its net fee lies within the `fee_band` around the reference and it contains every item in `evaluation.must_contain`.

## Considered Options

- Hour table for Advisory in the costing manual or a personal sheet: rejected (maintainer, Q36). Unrealistic for consulting work, and it would turn judgement into a lookup.
- Exact expected total as for the other business lines: rejected, it would punish reasonable estimates.
- Past offers as basis, graded by content and fee band (chosen): realistic, still checkable. The price is a weaker ground truth for Advisory Cases.
