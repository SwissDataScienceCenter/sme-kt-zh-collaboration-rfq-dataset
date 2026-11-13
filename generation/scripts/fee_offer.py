"""Word fee offer (Honorarofferte) for Advisory, built on Vorlage_Honorarofferte_Beratung.docx (DL-184).

fee_rows(sections)   group priced lines into the printed rows: hour rows per task with B1/B2/B3 hours,
                     flat fee rows (article group 98.1xx) and travel costs (98.120.01)
render(...)          write .docx and convert to PDF next to it; returns the list of amounts printed
"""
from decimal import Decimal
from pathlib import Path
import yaml
from docx.shared import Pt
from docx_helpers import ROOT, C, chf, new_doc, table, to_pdf
import renderers as R

CAT = {"99.100.06": "B1", "99.100.07": "B2", "99.100.08": "B3", "99.100.01": "MA"}
UNIT = {"Pau": "psch"}
_NR = yaml.safe_load(open(ROOT / "generation/world/norm_register.yaml", encoding="utf-8"))
NORMS = {n["nr"]: n["title"] for n in _NR["norms"] + _NR.get("laws", [])}

def _h(x):
    x = float(x); return ("%g" % x) if x else ""

def fee_rows(sections):
    """Return phases [{title, optional, rows:[{text, B1, B2, B3, amount}], flat:[line], travel:[line], hours_sum}]"""
    out = []
    for sec in sections:
        ph = {"title": sec["title"], "rows": [], "flat": [], "travel": [], "optional": False}
        for l in sec["lines"]:
            if l.get("optional"): ph["optional"] = True
            code = l.get("code", "")
            if code in CAT:
                row = next((r for r in ph["rows"] if r["text"] == l["text"]), None)
                if row is None:
                    row = {"text": l["text"], "B1": Decimal(0), "B2": Decimal(0), "B3": Decimal(0), "MA": Decimal(0), "amount": Decimal(0)}
                    ph["rows"].append(row)
                row[CAT[code]] += Decimal(str(l["qty"])); row["amount"] += l["total"]
            elif code.startswith("98.120"):
                ph["travel"].append(l)
            else:
                ph["flat"].append(l)
        if ph["optional"] and ph["rows"]:          # an option shows its own travel inside its table
            for l in ph["travel"]:
                ph["rows"].append({"text": l["text"], "B1": Decimal(0), "B2": Decimal(0), "B3": Decimal(0), "MA": Decimal(0), "amount": l["total"]})
            ph["travel"] = []
        ph["hours_sum"] = sum((r["amount"] for r in ph["rows"]), Decimal(0))
        out.append(ph)
    return out

def render(docx_path, o, result, sections, date, rates):
    """o: offer dict with recipient, number, subject, texts{...}, assumptions, client_provisions, exclusions, reservations,
    norms_named, signed_by, payment_terms, valid_until. rates: {'B1': 185, ...} for the offer year."""
    t = o["texts"]; printed = []
    doc = new_doc(title_line=True)
    for line in o["recipient"]:
        p = doc.add_paragraph(line); p.paragraph_format.left_indent = Pt(260); p.paragraph_format.space_after = Pt(0)
    doc.add_paragraph(f"\nIbach, {R.de_date(date, long=True).lstrip('0')}")
    ce = o.get("kind") == "Richtofferte"
    p = doc.add_paragraph(); r = p.add_run(f'{"RICHTOFFERTE" if ce else "HONORAROFFERTE"} {o["number"]}'); r.bold = True; r.font.size = Pt(13)
    p = doc.add_paragraph(); r = p.add_run("Betreff: " + o["subject"]); r.bold = True
    doc.add_paragraph(t["salutation"])
    for x in t.get("intro", []): doc.add_paragraph(x)

    doc.add_heading("1 Ausgangslage und Zielsetzung", level=2)
    for x in t["situation"]: doc.add_paragraph(x)
    doc.add_heading("2 Grundlagen", level=2)
    for x in t["documents"]: doc.add_paragraph(x, style="List Bullet")
    if o.get("norms_named"):
        doc.add_paragraph("Massgebende Normen (nur genannt, Auslegung im Einzelfall nicht Teil dieser Offerte): " +
                          "; ".join(f"{n} {NORMS.get(n, '')}".strip() for n in o["norms_named"]), style="List Bullet")

    phases = fee_rows(sections)
    base = [p_ for p_ in phases if not p_["optional"] and (p_["rows"] or p_["flat"] or p_["travel"])]
    opts = [p_ for p_ in phases if p_["optional"]]
    doc.add_heading("3 Leistungen und Aufwand nach Phasen", level=2)
    for p_ in [x for x in base if x["rows"]] + opts:
        doc.add_paragraph().add_run(p_["title"] + ("  (nicht im Total)" if p_["optional"] else "")).bold = True
        for x in t["phase_scope"].get(p_["title"], []): doc.add_paragraph(x, style="List Bullet")
        if p_["rows"]:
            ma = any(r["MA"] for ph in phases for r in ph["rows"])
            rows = [(r["text"], _h(r["B1"]), _h(r["B2"]), _h(r["B3"])) + ((_h(r["MA"]),) if ma else ()) + (chf(r["amount"]),) for r in p_["rows"]]
            rows.append(("Total Phase", "", "", "") + (("",) if ma else ()) + (chf(p_["hours_sum"]),))
            hdr = ["Leistung", "B1 Std", "B2 Std", "B3 Std"] + (["MA Std"] if ma else []) + ["Betrag CHF"]
            table(doc, hdr, rows, widths=[8.0, 1.5, 1.5, 1.5] + ([1.5] if ma else []) + [2.4], font=9); doc.add_paragraph()
            printed += [chf(r["amount"]) for r in p_["rows"]] + [chf(p_["hours_sum"])]
    doc.add_paragraph(f'Stundenansätze: B1 Geschäftsleitung CHF {rates["B1"]}, B2 Projektleitung CHF {rates["B2"]}, '
                      f'B3 Technik CHF {rates["B3"]}' + (f', MA Monteur CHF {rates["MA"]}' if any(r["MA"] for ph in phases for r in ph["rows"]) else '') + ' pro Stunde exkl. MWST.')

    flat = [l for p_ in base for l in p_["flat"]]
    doc.add_heading("4 Pauschalleistungen", level=2)
    if flat:
        table(doc, ["Leistung", "Menge", "Preis CHF", "Betrag CHF"],
              [(l["text"], f'{float(l["qty"]):g} {UNIT.get(l["unit"], l["unit"])}', chf(l["unit_price"]), chf(l["total"])) for l in flat], widths=[9.6, 2.2, 2.2, 2.6], font=9); doc.add_paragraph()
        printed += [chf(l["total"]) for l in flat]
    else:
        doc.add_paragraph("Keine.")
    for x in t.get("flat_notes", []): doc.add_paragraph(x)

    doc.add_heading("5 Nicht enthaltene Leistungen", level=2)
    for x in o["exclusions"]: doc.add_paragraph(x, style="List Bullet")
    doc.add_heading("6 Annahmen und Leistungen der Bauherrschaft", level=2)
    for x in o["assumptions"]: doc.add_paragraph(x, style="List Bullet")
    for x in o["client_provisions"]: doc.add_paragraph("Bauseits: " + x, style="List Bullet")
    for x in o.get("reservations", []): doc.add_paragraph("Vorbehalt: " + x, style="List Bullet")

    doc.add_heading("7 Honorar", level=2)
    hours = sum((p_["hours_sum"] for p_ in base), Decimal(0)); fl = sum((l["total"] for l in flat), Decimal(0))
    tr = sum((l["total"] for p_ in base for l in p_["travel"]), Decimal(0))
    nums = "1 bis %d" % len([p_ for p_ in base if p_["rows"]]) if len([p_ for p_ in base if p_["rows"]]) > 1 else "1"
    rows = [(f"Total Aufwand gemäss Phasen {nums}", chf(hours)), ("Pauschalleistungen", chf(fl))]
    if tr:
        km = sum(float(l["qty"]) for p_ in base for l in p_["travel"])
        up = [l for p_ in base for l in p_["travel"]][0]["unit_price"]
        rows.append((f"Spesen (km-Pauschale, {km:g} km à CHF {chf(up)})", chf(tr)))
    if result["discount"]:
        rows += [("Total brutto", chf(result["gross"])), (f'Rabatt {float(result["discount_rate"])*100:g} %', "-" + chf(result["discount"]))]
        printed += [chf(result["gross"]), chf(result["discount"])]
    rows += [("Total exkl. MWST", chf(result["net"])), (f'MWST {result["vat_rate"]*100:.1f} %', chf(result["vat"]))]
    if result["total"] != result["net"] + result["vat"]: rows.append(("Rundung", chf(result["total"] - result["net"] - result["vat"])))
    rows.append(("Total inkl. MWST", chf(result["total"])))
    table(doc, ["", "CHF"], rows, widths=[12, 4.6], font=9.5); doc.add_paragraph()
    printed += [chf(hours), chf(fl), chf(result["net"]), chf(result["vat"]), chf(result["total"])] + ([chf(tr)] if tr else [])
    if opts:
        doc.add_paragraph("Optionen (nicht im Total): " + "; ".join(f'{p_["title"].replace("Option: ", "")} CHF {chf(p_["hours_sum"] + sum((l["total"] for l in p_["flat"] + p_["travel"]), Decimal(0)))} exkl. MWST' for p_ in opts))
        printed += [chf(p_["hours_sum"] + sum((l["total"] for l in p_["flat"] + p_["travel"]), Decimal(0))) for p_ in opts]
    if ce:
        doc.add_paragraph(t["accuracy_note"])
    else:
        doc.add_paragraph("Verrechnung nach effektivem Aufwand. Das Total gilt als Kostendach; Mehraufwand nur nach schriftlicher "
                          "Zustimmung der Auftraggeberin." + (" " + t["options_note"] if opts and t.get("options_note") else ""))

    doc.add_heading("8 Termine", level=2)
    for x in t["schedule"]: doc.add_paragraph(x)
    doc.add_heading("9 Konditionen", level=2)
    table(doc, ["Gültigkeit", f'3 Monate, bis {R.de_date(o["valid_until"])}'], [("Zahlung", o["payment_terms"]),
                          ("AGB", "Es gelten unsere Allgemeinen Geschäftsbedingungen."), ("MWST-Nr.", f"{C['uid']} {C['vat_suffix']}")],
          widths=[4, 12.6], font=9.5); doc.add_paragraph()
    for x in t.get("closing", []): doc.add_paragraph(x)
    s = R.STAFF[o["signed_by"]]
    doc.add_paragraph("Freundliche Grüsse"); doc.add_paragraph(C["name"])
    doc.add_paragraph(f'\n{s["first"]} {s["last"]}\n{t.get("signature_role", s["role"])}')
    if not ce:
        p = doc.add_paragraph(); r = p.add_run("\nAuftrag erteilt" + (" " + t["order_by"] if t.get("order_by") else "")); r.bold = True
        doc.add_paragraph("Ort, Datum: ______________________    Unterschrift: ______________________")
    docx_path = Path(docx_path); doc.save(docx_path); to_pdf(docx_path)
    return printed
