# Plan: synthetic RFQ dataset

Status: confirmed 2026-10-30. Vocabulary follows `GLOSSARY.md`; decisions are recorded in `docs/adr/`.

## 1. Purpose

A synthetic dataset that traces how a fictional Swiss contractor turns RFQs into Offers. First use: a shared dataset for the 2026 cohort of the SME AI programme (Canton of Zürich) to prototype a Pipeline that automates parts of the business process. Later use: an evaluation set with known correct answers. The dataset must make the complexity of real SME data visible: incomplete and outdated information, verbal agreements, legacy documents, systems that are hard to interface with.

## 2. Ground rules

| Rule | Source |
|---|---|
| Build the Company world first, then one Case at a time with a review after each | ADR 0001 |
| Fictional company, no cohort SME data, Swiss legal and business framework | ADR 0002 |
| Dataset documents in Swiss German (no ß), framework documents in English | ADR 0003 |
| Norms named and summarised only; no price depends on norm details; federal law may be quoted | ADR 0004 |
| Past projects generated with the Case that needs them, then frozen | ADR 0005 |
| Ground truth in `case.yaml` first, documents written backwards from it | ADR 0006 |
| Real samples live in `_private_samples/` and are never copied | `.gitignore` |

## 3. Company world

**Fronalp Metall & Gebäudehülle AG**, Ibach (municipality of Schwyz). Founded 1962 as a locksmith and metal workshop, sheet metal and roofing added in the 1980s, System products since about 2005, Advisory since about 2015, third generation family business, about 35 people. Customers are B2B (architects, property managers, general contractors, housing cooperatives, public bodies) and B2C (homeowners).

Business lines: Advisory, Custom build, System product.

Roles (names fictional): owner and managing director, three project managers, one estimator, back office (two), workshop foreman, A fitters, B fitters, apprentices.

Systems: CRM and ERP exist but appear only as exports and as reference numbers inside documents (customer, item, project and offer numbers). Built in friction: IDs that differ between CRM and ERP, outdated records, fields kept only in people's heads. Further sources: shared offers mailbox, personal mailboxes, Excel costing sheets, a file server with one folder per project, phone and site notes.

Long history shows in the data: scanned past projects, an old Word offer template next to the current ERP layout, rules of thumb that live only in emails ("wie damals beim Auftrag Meier").

## 4. Business process

1. Intake: shared mailbox, phone, web form, tender portal, and File notes of verbal agreements.
2. Record the customer and lead in the CRM.
3. Assign the RFQ to a business line.
4. Clarify with the customer, with or without a site visit.
5. Cost: ERP item master, rate table, supplier and subcontractor quotes, similar past projects.
6. Review: the owner approves Offers above a threshold (value set in phase 1).
7. Send, follow up, record won or lost.

Variants: Advisory needs several meetings before the scope is clear. Custom build needs a site visit or written assumptions. System product skips most clarification and sends only adaptations to the workshop for checking.

## 5. Offer types

| Offer type | Business line | Shape |
|---|---|---|
| Fee offer | Advisory | Short (hours x rate per phase) to long (many pages describing work, citing regulations, assumptions, exclusions) |
| Unit price offer | Custom build | Grouped line items with quantity, unit, unit price; time and material, flat rate and optional items |
| Lump sum offer | Custom build | Narrative description of work and process, few priced positions, one fixed price |
| Tender response | Custom build | Architect's bill of quantities, priced |
| Configuration offer | System product | Catalogue item plus option and adaptation surcharges |
| Cost estimate | any | Non binding, states its accuracy (for example plus or minus 15%) |

### Conventions observed in the private samples (inform structure only, never copied)

- Cover letter, then description of work, process, client provisions (bauseitige Leistungen), costs, conditions, dates, signature block for order confirmation.
- Client provisions listed as text or as positions priced at 0.00 to mark the scope boundary.
- Optional items shown in brackets and excluded from the total, in a separate "Optionale Positionen" block, or as a unit price in the text.
- Assumptions stated per position ("Annahme 96 m2") or as a general note when no site visit took place.
- Waste and reserve added explicitly ("96 m2 + 15 %").
- Totals: discount, early payment discount (Skonto), VAT, rounding to 5 Rappen, sometimes an explicit rounding line.
- VAT depends on the year: 7.7% for documents dated 2018 to 2023, 8.1% from 2024. Past projects must use the rate of their year.
- Validity ranges from 14 days to 6 months; payment terms from 30 days net to instalments (1/2 on order, or 1/3 and 2/3).
- Reservations (Vorbehalte) listing what is billed by effort if conditions differ.
- Enclosures listed (general terms, test reports, declarations of conformity).
- Norms and regulations cited as selling arguments, not as price drivers (consistent with ADR 0004).
- Realistic noise: typos, unfilled template fields, totals in the heading that differ from the line items.

## 6. Cases

Generated one at a time in this order: C1, S1, A1, C2, S2, A3, C3, S3, A2.

| ID | Case | Offer type | Difficulty | Built in flaws |
|---|---|---|---|---|
| C1 | Terrace renovation | Unit price offer | Medium | No site visit, assumptions; area mismatch between heading and positions |
| S1 | Bike shelter for an apartment block via web form | Configuration offer | Easy | Clean baseline |
| A1 | Condition assessment of a 1970s flat roof for a property manager | Fee offer | Medium | Scope changes during the email thread |
| C2 | External steel staircase for a homeowner | Lump sum offer | Hard | Outdated dimensions in the thread, hand sketch, budget hint in a File note |
| S2 | Carport, wider than standard, in a high snow load area | Configuration offer | Medium | Adaptation rule from the catalogue, not from a norm |
| A3 | Railing safety check for a housing cooperative | Cost estimate, then Fee offer | Medium | Verbal agreement contradicts a later email |
| C3 | Facade cladding from an architect's tender | Tender response | Easy to read, large | Two items missing in the bill of quantities, tight deadline |
| S3 | Repeat customer quoting an old offer number | Cost estimate, then Configuration offer | Medium | Old prices, CRM and ERP IDs do not match |
| A2 | Straw roof feasibility for a listed farmhouse | Fee offer | Hard | Outside core expertise, subcontractor needed, long Offer |

## 7. Generation procedure per Case (ADR 0006)

1. Write `case.yaml`: customer (CRM ID), object, timeline, scope facts, line items referencing ERP item codes, deliberate flaws with their location, past projects.
2. Run the pricing script on `case.yaml` and the Company world: line totals, discount, Skonto on gross, VAT of the Offer year, 5 Rappen rounding.
3. Generate one or two past projects (Offer, one page summary, final invoice, measured quantities; older ones scanned) and freeze them into `company-world/past-projects/`.
4. Write request side documents backwards from the facts, placing each flaw on purpose.
5. Write case specific operations side documents (supplier quotes, internal costing sheet, File notes).
6. Render the Offer from the computed line items using the template of its Offer type.
7. Run validation: totals, cross references, IDs, glossary terms, every listed flaw present.
8. Review by the maintainer before the next Case.

## 8. File formats

| Content | Format |
|---|---|
| Emails with attachments and quoted history | `.eml` |
| Meeting transcripts | `.vtt` or `.txt` |
| File notes | `.docx` |
| Web form entries | `.json` plus PDF printout |
| Sketches | `.png` or `.jpg`, hand drawn style |
| Plans | `.pdf` |
| CRM and ERP exports | `.csv`, `.xlsx` |
| Price lists, rate tables, catalogue | `.xlsx` |
| Manuals, templates, general terms | `.docx` |
| Offers | `.pdf` plus `.docx` source |
| Old documents | scanned image `.pdf` |
| Ground truth | `case.yaml` |

No native CAD formats.

## 9. Repository layout

```
README.md
GLOSSARY.md
docs/plan.md
docs/adr/
company-world/
  organisation/     profile, history, org chart, staff list
  operations/       rate table, catalogue, costing manual, offer templates, general terms, operations manual, subcontractors
  regulations/      federal law excerpts, Norm register
  systems/          CRM and ERP exports
  past-projects/
cases/<case-id>/
  request/  operations/  offer/  case.yaml
generation/         specs, prompts, pricing and rendering scripts, validation
_private_samples/   ignored by git
```

## 10. Phases

| Phase | Content | Done when |
|---|---|---|
| 0 Framework | Glossary, ADRs 0001 to 0006, this plan | Done |
| 1 Company world core | Organisation, rate table, item master, System product catalogue with adaptation rules, costing manual (discounts, Skonto, VAT by year, rounding, validity, payment terms, approval threshold), offer templates (old Word, current ERP), general terms, operations manual, subcontractor list, federal law excerpts, Norm register, CRM and ERP exports, generation specs and pricing script | Done 2026-11-02 |
| 2 Pilot Case C1 | Full procedure from section 7, including past projects | Built 2026-11-02, awaiting review |
| 3 Remaining Cases | S1, A1, C2, S2, A3, C3, S3, A2 | Each reviewed |
| 4 Packaging | README, licence, Case index, GitHub repository | Published |
| Later | Evaluation set from `case.yaml` ground truth | Separate decision |

## 11. Parked

Maintenance contracts: see the note "parked/maintenance-contracts-extension.md" in the Claude project "Kt ZH Proj". Not stored in this repository and not used when creating Cases.

## 12. Settled at the start of phase 1 (2026-11-02)

- Glossary: Line item, Client provision, Optional item, Assumption, Reservation, Time and material, Item master, Measured quantities.
- Commercial rules: owner approval above CHF 50'000 net and for any discount above 3%; validity 3 months (current template) and 30 days (old Word template); B2C 30 days net; B2B Custom build above CHF 30'000 in instalments 30% on order, 60% by progress, 10% after handover; Skonto 2% within 10 days on the gross total; discount 0% standard, 2% for regular customers; VAT by document year; totals rounded to 5 Rappen.
- Rate table 2026 (CHF/h excl. VAT): A fitter 125, B fitter 115, apprentice 65, workshop 118, project manager 145; Advisory: owner 185, project manager 155, technician 125. One table per year from 2018.
- Item master about 250 articles in about 10 groups; System product catalogue with three families (bike shelter, carport, canopy).
- Staff list of about 35 fictional people with role, start year, initials and traits that leave traces in the data.

## 13. Before publishing

- Choose the licence.
- Check company name (zefix.ch), web domain, UID, phone numbers, supplier names (generation/world/operations.yaml) and customer names for collisions with real entities. Phone numbers use the unassigned block 041 000 xx xx and the UID has an invalid check digit on purpose.
- Known limitation: no site photos in Cases; findings are described in file notes (DL-185). Mention in the published README.
