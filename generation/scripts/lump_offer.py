"""Lump sum offers for metal construction (Pauschalofferte Metallbau), costed bottom-up (costing manual 3.2, DL-189).

cost(variant, F, world, year)  evaluate a costing variant from case.yaml: rows [group, label, qty_expr, unit, source]
                               groups M material and O surfaces (EK + 15 %), W workshop, K design, A fitting,
                               T transport and crane (VK), P project management, F partner services; plus 8 % risk and profit,
                               rounded up to CHF 10. Returns rows with qty, unit cost, amount and the lump sum.
sheet(path, variants, ...)     write the estimator's costing sheet (Kalkulationsblatt Metallbau) with Excel formulas
render(path_docx, o, ...)      Word lump sum offer from the 2016 template structure, with variants and an option
"""
import math
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font
from docx.shared import Pt
from docx_helpers import C, chf, new_doc, table, to_pdf
import renderers as R

GROUPS = {"M": "Material", "O": "Oberflächen", "W": "Werkstatt", "K": "Konstruktion", "A": "Montage",
          "T": "Transport und Kran", "P": "Projektleitung", "F": "Fremdleistungen"}
RATE_CAT = {"WS": "workshop", "MA": "fitter_a", "MB": "fitter_b", "PL": "project_manager"}

def _d(x): return Decimal(str(x)).quantize(Decimal("0.01"), ROUND_HALF_UP)

def unit_cost(src, world, year):
    kind, _, val = src.partition(":")
    if kind == "ek":
        it = world.items[val]; return _d(Decimal(it["ek"]) * world.factor(it["family"], year))
    if kind == "vk": return _d(world.price(val, year))
    if kind == "rate": return _d(world.rate(RATE_CAT[val], year))
    if kind == "fix": return _d(val)
    raise ValueError(src)

def cost(var, F, world, year, env_extra=None):
    ops = world.ops["offer_markup"]; env = {"F": F, "ceil": math.ceil, "floor": math.floor, "round": round, "sqrt": math.sqrt, "max": max}
    env.update(env_extra or {})
    rows, sums = [], {g: Decimal(0) for g in GROUPS}
    for g, label, qexpr, unit, src in var["rows"]:
        q = eval(str(qexpr), {"__builtins__": {}}, env) if isinstance(qexpr, str) else qexpr
        q = round(float(q), 2); uc = unit_cost(src, world, year); amt = _d(Decimal(str(q)) * uc)
        rows.append({"group": g, "label": label, "qty": q, "unit": unit, "unit_cost": uc, "amount": amt, "src": src, "expr": str(qexpr)})
        sums[g] += amt
    mat = _d((sums["M"] + sums["O"]) * (1 + Decimal(str(ops["material"]))))
    rest = sum((sums[g] for g in GROUPS if g not in ("M", "O")), Decimal(0))
    total_cost = mat + rest; wg = _d(total_cost * Decimal(str(ops["risk_and_profit"])))
    lump = Decimal(math.ceil((total_cost + wg) / 10) * 10)
    return {"rows": rows, "sums": sums, "material_with_markup": mat, "cost": total_cost, "risk_profit": wg, "lump": lump}

NAMES = {"F.stringer": "Wangenlänge", "F.treads": "Stufen", "F.risers": "Steigungen", "F.railing": "Geländerlänge",
         "F.stair_kg": "Gewicht Treppe", "F.coat_m2": "Pulverfläche", "F.handrail_wall": "Handlauflänge", "F.holders": "Halter"}
BASIS = [("height", "Höhe fertig", "m"), ("risers", "Steigungen", "Stk"), ("rise", "Steigung", "m"), ("treads", "Stufen", "Stk"),
         ("going", "Auftritt", "m"), ("diag", "Lauflinie schräg", "m"), ("stringer", "Wangenlänge", "m"), ("railing", "Geländerlänge", "m"),
         ("handrail_wall", "Handlauflänge", "m"), ("stair_kg", "Gewicht Treppe ohne Stufen", "kg"), ("coat_m2", "Pulverfläche", "m2")]

def sheet(path, variants, title, by, F=None):
    wb = Workbook(); wb.remove(wb.active)
    for key, (name, res) in variants.items():
        ws = wb.create_sheet(f"Variante {key}" if len(key) == 1 else key[:28])
        r = 1
        def put(vals, bold=False):
            nonlocal r
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                if bold: cell.font = Font(bold=True)
            r += 1
        put([title, "", "", "", "", by], bold=True); put([name]); r += 1
        if F:
            put(["", "Grundlagen"], bold=True)
            for k, lab, u in BASIS:
                if F.get(k) is not None: put(["", lab, F[k], u])
            r += 1
        put(["Gr.", "Position", "Menge", "EH", "Ansatz / EK", "Betrag", "Herleitung"], bold=True)
        first = r
        for row in res["rows"]:
            expr = row["expr"]
            for k_, v_ in NAMES.items(): expr = expr.replace(k_, v_)
            kind, _, val = row["src"].partition(":")
            src = {"ek": f"EK {val}", "vk": f"VK {val}", "rate": f"Ansatz {val}", "fix": "Offerte Partner" if row["group"] == "F" else "Richtpreis"}.get(kind, row["src"])
            put([row["group"], row["label"], row["qty"], row["unit"], float(row["unit_cost"]), f"=ROUND(C{r}*E{r},2)", f"{expr}  ({src})"])
        last = r - 1; r += 1
        def s_(g): return f'SUMIF(A{first}:A{last},"{g}",F{first}:F{last})'
        n0 = r
        put(["", "Material und Oberflächen + 15 %", "", "", "", f'=ROUND(({s_("M")}+{s_("O")})*1.15,2)'])
        for g in ("W", "K", "A", "T", "P", "F"): put(["", GROUPS[g], "", "", "", f"={s_(g)}"])
        nc = r; put(["", "Kosten total", "", "", "", f"=SUM(F{n0}:F{nc - 1})"], bold=True)
        put(["", "Risiko und Gewinn 8 %", "", "", "", f"=ROUND(F{nc}*0.08,2)"])
        put(["", "Pauschalpreis (auf CHF 10 aufgerundet)", "", "", "", f"=CEILING(F{nc}+F{nc + 1},10)"], bold=True)
        for col, w in zip("ABCDEFG", [5, 46, 9, 6, 11, 12, 60]): ws.column_dimensions[col].width = w
    wb.save(path)

def render(docx_path, o, variants, option, date, texts):
    """variants: [(key, title, lump, vat, total, description lines)], option: (text, amount) or None"""
    t = texts; printed = []
    doc = new_doc(title_line=True)
    for line in o["recipient"]:
        p = doc.add_paragraph(line); p.paragraph_format.left_indent = Pt(260); p.paragraph_format.space_after = Pt(0)
    doc.add_paragraph()
    doc.add_paragraph(f"Ibach, {R.de_date(date, long=True).lstrip('0')}")
    p = doc.add_paragraph(); r = p.add_run(f'OFFERTE {o["number"]}'); r.bold = True; r.font.size = Pt(13)
    doc.add_paragraph("Objekt: " + o["object"])
    doc.add_paragraph(t["salutation"])
    doc.add_paragraph(t.get("intro", "Besten Dank für Ihre Anfrage. Gerne unterbreiten wir Ihnen unsere Offerte wie folgt:"))
    doc.add_heading("1. Ausgangslage", level=2)
    for x in t["situation"]: doc.add_paragraph(x)
    doc.add_heading("2. Leistungsbeschrieb", level=2)
    for x in t["scope_common"]: doc.add_paragraph(x, style="List Bullet")
    for key, title, *_ , desc in variants:
        doc.add_paragraph().add_run(title).bold = True
        for x in desc: doc.add_paragraph(x, style="List Bullet")
    doc.add_paragraph().add_run(t.get("package_title", "Im Pauschalpreis enthalten")).bold = True
    for x in t["package"]: doc.add_paragraph(x, style="List Bullet")
    doc.add_heading("3. Bauseitige Leistungen", level=2)
    for x in o["client_provisions"]: doc.add_paragraph(x, style="List Bullet")
    if o.get("exclusions"):
        doc.add_paragraph().add_run("Nicht enthalten").bold = True
        for x in o["exclusions"]: doc.add_paragraph(x, style="List Bullet")
    if o.get("assumptions") or o.get("reservations"):
        doc.add_paragraph().add_run("Annahmen und Vorbehalte").bold = True
        for x in o.get("assumptions", []) + o.get("reservations", []): doc.add_paragraph(x, style="List Bullet")
    sec = iter(range(4, 10))
    doc.add_heading(f"{next(sec)}. Kosten", level=2)
    rows = []
    for i, (key, title, lump, vat, total, _) in enumerate(variants, 1):
        rows += [(str(i), title + ", pauschal", chf(lump)), ("", f'MWST {o["vat_rate"] * 100:.1f} %', chf(vat)), ("", "Total inkl. MWST", chf(total))]
        printed += [chf(lump), chf(vat), chf(total)]
    table(doc, ["Pos.", "Beschreibung", "Betrag CHF"], rows, widths=[1.5, 11, 3.5])
    doc.add_paragraph()
    if option:
        doc.add_heading(f"{next(sec)}. Optional", level=2)
        doc.add_paragraph(f"{option[0]}: CHF {chf(option[1])} exkl. MWST, pauschal, Oberfläche wie die gewählte Variante.")
        printed.append(chf(option[1]))
    doc.add_heading(f"{next(sec)}. Konditionen", level=2)
    table(doc, ["Gültigkeit", o["validity_text"]], [("Zahlung", o["payment_text"]), ("Garantie", "2 Jahre gemäss SIA-Norm 118"),
                                                    ("AGB", "Es gelten unsere Allgemeinen Geschäftsbedingungen."),
                                                    ("MWST-Nr.", f"{C['uid']} {C['vat_suffix']}")], widths=[4, 12])
    doc.add_paragraph()
    doc.add_heading(f"{next(sec)}. Termine", level=2)
    for x in t["schedule"]: doc.add_paragraph(x)
    for x in t.get("closing", ["Wir danken Ihnen für Ihr Interesse und freuen uns auf Ihren Auftrag. Für Fragen stehen wir gerne zur Verfügung."]):
        doc.add_paragraph(x)
    s = R.STAFF[o["signed_by"]]
    doc.add_paragraph("Freundliche Grüsse"); doc.add_paragraph(C["name"])
    doc.add_paragraph(f'\n{s["first"]} {s["last"]}\n{s["role"]}')
    doc.add_paragraph()
    p = doc.add_paragraph(); r = p.add_run("Auftragsbestätigung"); r.bold = True
    if len(variants) > 1:
        doc.add_paragraph("Wir bestellen: " + "   ".join(f"[ ] {v[1].split(':')[0]}" for v in variants) +
                          ("   [ ] Option" if option else ""))
    doc.add_paragraph("Ort, Datum: ______________________    Unterschrift: ______________________")
    docx_path = Path(docx_path); doc.save(docx_path); to_pdf(docx_path)
    return printed
