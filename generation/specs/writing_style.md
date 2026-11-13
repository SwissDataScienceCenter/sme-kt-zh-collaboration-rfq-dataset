# Writing style for Case documents

All Case documents are Swiss German (ADR 0003): no ß, Swiss vocabulary (Offerte, Velo, parkieren, Abwart, Pendenz, allfällig), numbers as 12'345.60, dates as 07.07.2026 in documents and 2026-07-07 only in file names, units m2 or m², CHF before the amount.

## Staff voices (from organisation.yaml traits)

| Initials | Voice |
|---|---|
| CSc | Long, precise, polite emails; signs "Freundliche Grüsse, Claudia Schuler"; approves Offers by phone, so approvals show up only as File notes or CRM activities |
| MSc | Forwards old cases with "wie damals bei ..."; short, dialect coloured; no signature |
| RBe | Very short, often from phone ("Gesendet von meinem iPhone"), no greeting, own abbreviations (Terr., Abd., Spe., ca.) |
| SKä | Structured, numbered questions, uses CRM numbers in subject lines |
| DIn | Friendly, copies catalogue text verbatim, sometimes quotes outdated prices |
| BUl | Writes little; knowledge lives in Excel comments and his rules of thumb |
| EFö | Forwards RFQs with a one line note ("zur Bearbeitung, Lead L-2026-123 erfasst") |
| LMa | CRM notes with English abbreviations (fup, asap, tbd) |
| NSu | Reports with photo references and measured values |
| TRe | Handwritten notes, delivered as photos |

## Customer voices

| Category | Voice |
|---|---|
| Privat | Informal or formal mix, incomplete information, photos from phone, own measurements in cm or "ca.", references to neighbours or earlier work |
| Verwaltung | Formal, ticket or object numbers, several decision makers, budget and deadline pressure |
| Genossenschaft | Board decisions, meeting dates, minutes excerpts, cost sensitivity |
| Architektur | Precise, tender documents, deadlines, NPK structure, plans as PDF |
| GU | Short deadlines, price pressure, own contract terms |
| Öffentlich | Formal procedure, references to regulations, approvals |

## Document conventions

- Emails as `.eml` with full headers, quoted history below the reply, attachments embedded. Threads may contain outdated information further down.
- File notes use `templates/Vorlage_Aktennotiz.docx`, sometimes as a photo of a handwritten note instead.
- Meetings: notes written by one participant (formal Besprechungsnotiz from property managers and planners, terse File notes from Fronalp staff). Notes can omit or misstate details. Transcripts (`.vtt`) only for short calls, with speech recognition errors.
- Web form entries as `.json` plus the automatic notification email.
- Offers are rendered only from pricing.py output.

## Flaw types

| Type | Meaning |
|---|---|
| missing_info | Information needed for the Offer is absent |
| outdated_info | Information superseded later in the thread or by a newer document |
| contradiction | Two sources state different facts |
| ambiguity | A reference can point to several entities (people, customers, objects) |
| wrong_reference | Old or wrong offer, customer or article number |
| unit_mismatch | Quantities in different units or inconsistent rounding |
| out_of_scope | Request partly outside what the company offers |
| missing_attachment | Attachment announced but not included |
| stale_template | Template or price list from an earlier year used |
| verbal_only | Agreement exists only in a File note or a transcript |
| noise | Typos, unfilled fields, irrelevant content |
