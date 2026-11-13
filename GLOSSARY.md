# Synthetic RFQ Dataset

Vocabulary for a synthetic dataset that traces how a fictional Swiss contractor turns customer requests into priced offers. The dataset serves the SME AI programme cohort (Canton of Zürich) for prototyping RFQ automation.

## Process

**RFQ** (Kundenanfrage):
A customer's request for a priced offer, including every message, meeting note and attachment that belongs to it.
_Avoid_: Inquiry, request, Anfrage

**Offer** (Offerte):
The priced document the company sends to the customer in answer to an RFQ, ending in a total in CHF.
_Avoid_: Quote, quotation, proposal, Angebot

**Business process**:
The fictional company's manual workflow from receiving an RFQ to sending an Offer.
_Avoid_: Process (unqualified), workflow

**Pipeline**:
The AI system that cohort participants build to automate parts of the business process.
_Avoid_: Process (unqualified), workflow

## Offer content

**Line item** (Position):
One priced row of an Offer.
_Avoid_: Item, row

**Client provision** (bauseitige Leistung):
Work or resources the customer provides, listed to mark where the company's scope ends.
_Avoid_: Customer contribution

**Optional item** (Eventualposition):
A Line item that is offered but not included in the total.
_Avoid_: Option, alternative

**Assumption** (Annahme):
A stated basis for a quantity or price that has not been checked on site.

**Reservation** (Vorbehalt):
A condition under which additional work is billed by effort.

**Time and material** (Regie):
Work billed by hours and material actually used.
_Avoid_: Effort billing, hourly work

**Item master** (Artikelstamm):
The ERP list of articles with codes, units and prices.

**Measured quantities** (Ausmass):
The quantities measured after the work, on which the final invoice is based.

## Company

**Company world**:
Everything that stays fixed across Cases: organisation, people, history, rates, catalogue, manuals, regulations and past projects.
_Avoid_: Scenario, framework, setting

**Advisory** (Beratung):
Business line that sells expertise: condition assessments, renovation concepts and feasibility studies, priced by effort.
_Avoid_: Consulting, planning

**Custom build** (Projektgeschäft):
Business line that designs and builds one off work such as roof renovations, stairs, railings and cladding, priced item by item.
_Avoid_: Custom project, Sonderanfertigung

**System product** (Systemprodukt):
Business line that sells the company's own catalogue items, such as canopies and carports, with minor adaptations.
_Avoid_: Standard product, ready made product, catalogue product

## Offer types

**Fee offer** (Honorarofferte):
An Advisory Offer priced by estimated hours per phase and rate category, with deliverables, assumptions, exclusions and cited regulations.
_Avoid_: Consulting offer, quote

**Unit price offer** (Einheitspreisofferte):
A Custom build Offer of line items priced per unit of quantity, settled on measured quantities; may include time and material, flat rate and optional items.
_Avoid_: Itemised quote

**Lump sum offer** (Pauschalofferte):
A Custom build Offer with one fixed price for a precisely defined scope.
_Avoid_: Fixed price quote, Globalofferte

**Tender response** (Eingabe Leistungsverzeichnis):
A Custom build Offer made by pricing a bill of quantities that the customer's architect or planner has prepared.
_Avoid_: Bid, Submission

**Configuration offer** (Konfigurationsofferte):
A System product Offer made of a catalogue item plus surcharges for the chosen options and adaptations.
_Avoid_: Product quote, catalogue offer

**Cost estimate** (Richtofferte):
A non binding Offer giving an approximate price early in the business process, in any business line.
_Avoid_: Budget offer, Kostenvoranschlag, rough quote

## Data

**Case**:
One RFQ with all its documents, up to and including its Offer; the unit in which the dataset is generated.
_Avoid_: Scenario, example, sample

**File note** (Aktennotiz):
A short internal record of a phone call, site conversation or verbal agreement, often incomplete and sometimes contradicted by later emails.
_Avoid_: Memo, call note

**Meeting transcript** (Gesprächstranskript):
The full transcript of a recorded meeting, for example from Teams.
_Avoid_: Minutes, protocol

**Request side**:
Documents that come from or concern the customer before the Offer exists: emails, meeting transcripts, request forms, sketches.

**Operations side**:
Internal documents the company uses to price and plan work: line item specifications, hourly rates, past projects, operations manuals, regulations.

**Offer side**:
The finished Offer, whose level of detail depends on the kind of work offered.

**Norm register** (Normenverzeichnis):
A Company world document listing each relevant norm by number and title with a short description in our own words.
_Avoid_: Norm glossary, norms book
