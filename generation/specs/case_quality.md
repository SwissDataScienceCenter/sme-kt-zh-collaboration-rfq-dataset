# Case quality concept

How every Case is built so that it holds together down to the details. Applies from Case S1 on; C1 was reworked to meet it.

## 1. One source per fact

| Layer | Where | Rule |
|---|---|---|
| Company world | `generation/world/*.yaml`, `item_master.csv`, Faustregeln and manuals | Every rule an estimator uses must be written here |
| Case facts | `cases/<id>/case.yaml` → `facts` | Every number of the Case exists once (dimensions, counts, dates, hours) |
| Derived values | computed by `build_case.py` (area, edges, perimeter) | Never typed by hand |
| Documents | rendered from facts and `generation/cases/<id>/content.yaml` | Texts may restate facts, never invent new ones |

## 2. Story bible before rendering

Before any document is written, the Case gets a short story bible in `case.yaml` and is confirmed by the maintainer:
- facts and scope, who the client is and who only initiates
- timeline with every event and who is available when (holidays, absences, deadlines, meetings)
- systems state: what the CRM and ERP show at the time the estimator works (company exports are a snapshot of 2026-04-30; Case state lives in the Case folder, for example a CRM print)
- the Offer lines with a `src` for each, the deliberate flaws with a `marker`, and `acceptable_alternatives` where the written rules allow more than one answer

## 3. Automatic gates in build_case.py

The build fails if any of these fails:
- totals and every line amount printed in the Offer PDF match `pricing.py`
- every `checks` expression in `case.yaml` holds (quantities against facts and plan)
- the costing sheet formulas evaluate to the offered quantities
- building plans: rooms do not overlap and tile the footprint, every door sits on a wall shared by the two rooms it joins, every room is reachable from the entrance, window doors and windows sit on exterior walls, counts match the facts
- chronology: no mail after the Offer date, quoted mails earlier than the reply, staff mails on working days 06 to 20 h, file name date equals the Date header, validity three months
- every identifier (L-, KC-, D-, OF-, P-) in the Case exists in the systems data
- no ß anywhere; every deliberate flaw marker is present; every Offer line has a source

## 4. Domain checklist (extend per trade)

Terrace and flat roof: build-up height and door thresholds, fall protection and railing during works, drains and emergency overflows, disposal volumes, permit for public ground and crane, protection of inhabited flats, client provisions, cost split in condominiums, test method, warranty reference.
System products: number of bikes or cars vs modules, available length or width, terrain slope (column extension), altitude rule (carports), wind side walls, foundations and utilities in the ground, power for lighting, building permit, access for the delivery truck, lead time vs desired date and permit procedure (publication, objection period, production only after a binding permit), colour surcharge, transport zone, surcharge bases and rounding when several rules apply, snow load class and type structural calculation (SIA 261), clearance vs vehicle height, frost safe foundations, downpipe connection, profile poles, deposit and price hold when installation moves to the next year.
Metal construction (stairs, railings): height from finished floor to finished floor, rise and going (2h + a), clearance before the first step, open riser gap with children, railing height, infill and graspable handrail (SIA 358, bfu), connection to existing slabs (engineer, anchors, reservation), foundations, stringer weight vs crane, surface (galvanized, duplex, warping of thin sheet), opening of existing railings, temporary fall protection, permit filing and realistic schedule, deposit for private customers, partner quote validity, saving variant that keeps safety.
Advisory (condition assessment of roofs): client vs representative and who signs, area from plans, phases and what is excluded after scope changes, pollutants before opening (asbestos, PCB, PAH by building year), fall protection with two persons, number of probe openings per roof, moisture measurement, report content (build-up, drainage, overflows, upstands, remaining life, immediate measures), variants and cost estimate accuracy, board or owners' meeting dates and company holidays, cost ceiling with written consent for extra effort, lab costs as pass-through, travel, options and when they must be ordered, independence if the company also bids.
Advisory (railing checks): number of railings vs number of flats, stairwells and other railings in scope, anchors per railing, pull test sample per risk group with test load, protocol and handling of a failed test, insurer acceptance of the method, access to flats twice, immediate securing of red findings, meaning and deadline per traffic light colour, existing railings vs today's SIA 358 / bfu without legal judgement, difference to an earlier Richtofferte explained, who signs for the cooperative.

## 5. QA loop with independent agents

1. `python3 generation/scripts/build_review_bundle.py <id>` creates `_review/<id>/` (text and PNG renders).
2. Three reviewers run in parallel with the prompts in `generation/specs/qa_prompts.md`:
   coherence (finds every inconsistency), domain realism (Swiss practice), blind derivability (rebuilds the Offer without seeing it).
3. Every finding is triaged in `docs/qa/<id>.md`: fixed, intended (flaw id) or accepted (reason).
4. Rebuild, then a second round of coherence and derivability.
5. Done when no high or medium finding is unexplained and the blind Offer differs from the expected Offer only by documented alternatives.
