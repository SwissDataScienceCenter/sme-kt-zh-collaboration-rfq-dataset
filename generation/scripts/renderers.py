"""Shared renderers for Case and past project documents.

metabau_offer(...)    MetaBau offer print, layouts "2024" (cover page with totals) and "2023" (older, totals at the end)
metabau_invoice(...)  MetaBau final invoice (Schlussrechnung) based on Measured quantities
eml(...)              RFC 5322 email with quoted history and attachments
vtt(...)              Teams style WebVTT transcript
file_note(...)        File note (Aktennotiz) as .docx from the company template structure
scanify(...)          turn a PDF into a slightly skewed, noisy greyscale scan
"""
import datetime as dt, io, random, subprocess, tempfile, zlib
from email.message import EmailMessage
from email.utils import format_datetime, make_msgid
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, black
from docx_helpers import C, ORG, chf, new_doc, table
from docx.shared import Pt

STAFF = {s["ini"]: s for s in ORG["staff"]}
for s in ORG["staff"]:
    STAFF[s.get("display_ini", s["ini"])] = s
RED = HexColor("#8C1D18")

def staff_mail(ini):
    s = STAFF[ini]
    t = f'{s["first"]}.{s["last"]}'.lower().replace(" ", "")
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ã", "a")):
        t = t.replace(a, b)
    return f"{t}@fronalp-mgh.ch"

def de_date(d, long=False):
    d = dt.date.fromisoformat(str(d))
    if not long:
        return d.strftime("%d.%m.%Y")
    m = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"]
    return f"{d.day:02d}. {m[d.month-1]} {d.year}"

def qty_fmt(q, dec=3):
    return f"{float(q):,.{dec}f}".replace(",", "'")

# ---------------------------------------------------------------------------------------------- offer print
class _Doc:
    def __init__(self, path, footer_left, footer_right_fmt):
        self.c = canvas.Canvas(str(path), pagesize=A4); self.w, self.h = A4
        self.page = 1; self.fl = footer_left; self.fr = footer_right_fmt; self.y = 0
    def footer(self):
        c = self.c; c.setFont("Helvetica", 8); c.setFillColor(black)
        c.line(50, 50, self.w - 50, 50)
        c.drawString(50, 38, self.fl); c.drawRightString(self.w - 50, 38, self.fr.format(n=self.page))
    def new_page(self):
        self.footer(); self.c.showPage(); self.page += 1
    def save(self):
        self.footer(); self.c.save()

def _letterhead(c, w, h, layout):
    c.setFont("Helvetica-Bold", 8.5); c.drawString(50, h - 50, C["name"])
    c.setFont("Helvetica", 8.5)
    for i, t in enumerate([C["street"], f'{C["zip"]} {C["city"]}', C["phone_main"]]):
        c.drawString(50, h - 61 - i * 10.5, t)
    c.drawString(200, h - 50, C["email_general"]); c.drawString(200, h - 61, C["web"])
    if layout == "2024":
        c.setFont("Helvetica-Bold", 8.5); c.drawString(200, h - 82, "Pikett Dach und Spenglerei")
        c.setFont("Helvetica", 8.5); c.drawString(200, h - 93, C["phone_picket"])
    c.setFillColor(RED); c.setFont("Helvetica-Bold", 22); c.drawString(380, h - 66, "FRONALP")
    c.setFont("Helvetica-Bold", 9); c.drawString(381, h - 80, "Metall & Gebäudehülle"); c.setFillColor(black)

def _totals_block(c, x0, x1, y, r, cur="CHF"):
    rows = [("Total brutto", "", "", "", chf(r["gross"]), True)]
    if r["discount"]:
        rows.append(("Rabatt", f'{r["discount_rate"]*100:.2f} %', chf(r["gross"]), "-" + chf(r["discount"]), chf(r["gross"] - r["discount"]), False))
    if r["skonto"]:
        rows.append(("Skonto", f'{r["skonto_rate"]*100:.2f} %', chf(r["gross"]), "-" + chf(r["skonto"]), chf(r["net"]), False))
    rows.append(("Zwischentotal", "", "", "", chf(r["net"]), True))
    rows.append(("Mehrwertsteuer", f'{r["vat_rate"]*100:.2f} %', chf(r["net"]), chf(r["vat"]), chf(r["net"] + r["vat"]), False))
    if r["total"] != r["net"] + r["vat"]:
        rows.append(("Rundung", "", "", chf(r["total"] - r["net"] - r["vat"]), chf(r["total"]), False))
    for label, pct, base, amt, run, bold in rows:
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 9)
        c.drawString(x0, y, label); c.drawString(x0 + 150, y, pct); c.drawString(x0 + 200, y, base)
        c.drawRightString(x0 + 360, y, amt); c.drawRightString(x1, y, run); y -= 12
    c.line(x0, y + 4, x1, y + 4); y -= 10
    c.setFont("Helvetica-Bold", 10); c.drawString(x0, y, f"Total {cur} inkl. MWST"); c.drawRightString(x1, y, chf(r["total"]))
    c.line(x0, y - 5, x1, y - 5); c.line(x0, y - 7, x1, y - 7)
    return y - 24

def metabau_offer(path, case_offer, result, sections, date, layout="2024", texts=None):
    """case_offer: dict from case.yaml 'offer'; result: pricing.compute_offer output; sections: [{title, lines:[priced rows]}]"""
    texts = texts or {}
    pm = STAFF[case_offer["project_manager"]]
    nr = case_offer["number"]; d = _Doc(path, de_date(date), "Seite {n}" if layout == "2024" else nr + " / Seite {n}")
    c, w, h = d.c, d.w, d.h
    c.setTitle(f"Offerte {nr}")
    _letterhead(c, w, h, layout)
    y = h - 170
    c.setFont("Helvetica", 10)
    for t in case_offer["recipient"]:
        c.drawString(50, y, t); y -= 12.5
    y -= 28; c.drawString(50, y, f'Ibach, {de_date(date, long=True)}'); y -= 34
    c.setFont("Helvetica-Bold", 13); c.drawString(50, y, "Offerte")
    c.setFont("Helvetica-Bold", 10); c.drawString(200, y, nr); y -= 14
    c.drawString(50, y, "Objekt")
    for t in case_offer["object_lines"]:
        c.drawString(200, y, t); y -= 12.5
    y -= 6; c.setFont("Helvetica", 9.5)
    c.line(50, y + 8, w - 50, y + 8)
    c.drawString(50, y - 4, "Projektleiter")
    c.drawString(200, y - 4, f'{pm["first"]} {pm["last"]}, {staff_mail(case_offer["project_manager"])}, {C["phone_main"]}')
    y -= 24
    if case_offer.get("copy_to"):
        c.drawString(50, y, "Kopie an"); c.drawString(200, y, case_offer["copy_to"]); y -= 16
    if layout == "2024":
        c.line(50, y + 8, w - 50, y + 8); y -= 8
        y = _totals_block(c, 50, w - 50, y, result)
        c.setFont("Helvetica", 10)
        c.drawString(50, y, "Offerte gültig bis:"); c.drawString(200, y, de_date(case_offer["valid_until"], long=True)); y -= 18
        c.drawString(50, y, "Zahlungskonditionen:"); y2 = y
        for t in _wrap(case_offer["payment_text"], 60):
            c.drawString(200, y2, t); y2 -= 12
        y = y2 - 16
        for t in _wrap(texts.get("closing", "Wir danken für Ihre Anfrage und freuen uns, wenn wir die offerierten Arbeiten für Sie ausführen dürfen."), 95):
            c.drawString(50, y, t); y -= 12.5
        c.setFont("Helvetica", 8.5); c.drawString(50, y - 6, f"Es gelten unsere Allgemeinen Geschäftsbedingungen.   MWST-Nr. {C['uid']} MWST"); c.setFont("Helvetica", 10)
        y = 170
        c.setFont("Helvetica", 10); c.drawString(50, y, C["name"])
        c.setFont("Helvetica", 8.5); c.drawString(330, y, case_offer.get("signatory") or case_offer["recipient"][0]); c.setFont("Helvetica", 10)
        c.line(50, y - 40, 250, y - 40); c.line(330, y - 40, w - 50, y - 40)
        c.drawString(50, y - 52, f'{pm["first"]} {pm["last"]}'); c.drawString(330, y - 52, "Ort, Datum, Unterschrift")
        c.drawString(50, y - 64, pm["role"])
        d.new_page(); y = h - 60
    else:
        c.setFont("Helvetica", 10)
        for t in _wrap(texts.get("intro", "Besten Dank für Ihre Anfrage. Gerne offerieren wir Ihnen wie folgt:"), 95):
            c.drawString(50, y, t); y -= 12.5
        y -= 10

    def header(y):
        c.setFont("Helvetica", 8.5); c.line(50, y + 12, w - 50, y + 12)
        c.drawString(50, y + 2, "Offerte"); c.drawString(110, y + 2, f'{nr}  {case_offer["object_lines"][0]}')
        c.setFont("Helvetica-Bold" if layout == "2023" else "Helvetica", 9)
        c.drawString(50, y - 12, "Nr."); c.drawString(110, y - 12, "Beschreibung")
        c.drawRightString(400, y - 12, "Menge"); c.drawRightString(475, y - 12, "Preis"); c.drawRightString(w - 50, y - 12, "Betrag")
        c.line(50, y - 17, w - 50, y - 17)
        return y - 36
    y = header(y)
    def need(y, space):
        if y - space < 70:
            d.new_page(); return header(h - 60)
        return y
    c.setFont("Helvetica-Bold", 11)
    for t in case_offer.get("heading", []):
        y = need(y, 14); c.drawString(110, y, t); y -= 14
    y -= 6
    if case_offer.get("hint_text"):
        y = need(y, 20); c.setFont("Helvetica-Bold", 11); c.drawString(110, y, "Hinweis:"); y -= 16
        c.setFont("Helvetica-Bold", 9.5)
        for t in _wrap(case_offer["hint_text"], 52):
            y = need(y, 12); c.drawString(110, y, t); y -= 12
        y -= 12
    if case_offer.get("client_provisions"):
        y = need(y, 30); c.setFont("Helvetica", 9.5); c.drawString(110, y, "Leistungen bauseits:"); c.line(110, y - 2, 200, y - 2); y -= 16
        for t in case_offer["client_provisions"]:
            for k, part in enumerate(_wrap(t, 55)):
                y = need(y, 12); c.drawString(110, y, ("- " if k == 0 else "  ") + part); y -= 12
        y -= 12
    for title, items in (("Annahmen:", case_offer.get("assumptions", [])), ("Beschrieb:", case_offer.get("description", []))):
        if not items: continue
        y = need(y, 30); c.setFont("Helvetica", 9.5); c.drawString(110, y, title); c.line(110, y - 2, 110 + c.stringWidth(title, "Helvetica", 9.5), y - 2); y -= 16
        for t in items:
            for k, part in enumerate(_wrap(t, 55)):
                y = need(y, 12); c.drawString(110, y, ("- " if k == 0 else "  ") + part); y -= 12
        y -= 12
    pos_dec = 3 if layout == "2024" else 2
    for si, sec in enumerate(sections, 1):
        y = need(y, 40); c.setFont("Helvetica", 10); c.drawString(110, y, sec["title"])
        c.line(110, y - 2, 110 + c.stringWidth(sec["title"], "Helvetica", 10), y - 2); y -= 22
        for li, ln in enumerate(sec["lines"], 1):
            text = _wrap(ln["text"], 52); details = [p for t in ln.get("detail", []) for p in _wrap(t, 52)]
            y = need(y, 12 * (len(text) + len(details)) + 16)
            c.setFont("Helvetica", 9.5)
            c.drawString(50, y, f"{si}.{li:02d}")
            for k, t in enumerate(text):
                c.drawString(110, y - 12 * k, t)
            c.drawRightString(400, y, f'{qty_fmt(ln["qty"], pos_dec)} {ln["unit"]}')
            c.drawRightString(475, y, chf(ln["unit_price"]))
            amount = chf(ln["total"])
            c.drawRightString(w - 50, y, f"({amount})" if ln.get("optional") else amount)
            y -= 12 * len(text) + 4
            for t in details:
                c.drawString(110, y, t); y -= 11.5
            y -= 14
    y = need(y, 40); c.line(110, y + 6, w - 50, y + 6); c.setFont("Helvetica-Bold", 10)
    cur = "CHF" if layout == "2024" else "Fr."
    c.drawString(110, y - 8, f'Total, {case_offer["object_lines"][1]}'); c.drawRightString(w - 50, y - 8, f'{cur} {chf(result["gross"])}')
    y -= 30
    if result["optional_sum"]:
        c.setFont("Helvetica", 8.5); c.drawString(110, y, "Beträge in Klammern sind Eventualpositionen und im Total nicht enthalten."); y -= 18
    if case_offer.get("reservations"):
        y = need(y, 30); c.setFont("Helvetica", 9.5); c.drawString(110, y, "Vorbehalte:"); y -= 14
        for t in case_offer["reservations"]:
            for k, part in enumerate(_wrap(t, 70)):
                y = need(y, 12); c.drawString(110, y, ("- " if k == 0 else "  ") + part); y -= 12
    if layout == "2024":
        d.new_page(); y = header(h - 60) - 6
    else:
        y = need(y, 120) - 10
    _totals_block(c, 50, w - 50, y, result, cur="CHF" if layout == "2024" else "Fr.")
    if layout == "2023":
        y -= 110; c.setFont("Helvetica", 10)
        c.drawString(50, y, f'Gültigkeit: 30 Tage. Zahlung: {case_offer["payment_text"]}'); y -= 30
        c.drawString(50, y, "Freundliche Grüsse"); c.drawString(50, y - 14, C["name"]); c.drawString(50, y - 40, f'{pm["first"]} {pm["last"]}')
    d.save()

def _wrap(text, width):
    words, lines, cur = str(text).split(), [], ""
    for wd in words:
        if len(cur) + len(wd) + 1 > width and cur:
            lines.append(cur); cur = wd
        else:
            cur = (cur + " " + wd).strip()
    if cur: lines.append(cur)
    return lines or [""]

# ---------------------------------------------------------------------------------------------- invoice
def metabau_invoice(path, inv, result, rows, date):
    """inv: number, project, recipient, object, offer_ref, akonto (list of (text, amount)); rows: priced lines (Measured quantities)"""
    d = _Doc(path, de_date(date), "Seite {n}"); c, w, h = d.c, d.w, d.h
    c.setTitle(f'Schlussrechnung {inv["number"]}')
    _letterhead(c, w, h, "2024" if int(str(date)[:4]) >= 2024 else "2023")
    y = h - 170; c.setFont("Helvetica", 10)
    for t in inv["recipient"]:
        c.drawString(50, y, t); y -= 12.5
    y -= 26; c.drawString(50, y, f'Ibach, {de_date(date, long=True)}'); y -= 30
    c.setFont("Helvetica-Bold", 13); c.drawString(50, y, "Schlussrechnung"); c.setFont("Helvetica-Bold", 10); c.drawString(200, y, inv["number"]); y -= 16
    c.setFont("Helvetica", 9.5)
    for k, v in [("Projekt", inv["project"]), ("Objekt", inv["object"]), ("Offerte", inv["offer_ref"]), ("Ausführung", inv["period"])]:
        c.drawString(50, y, k); c.drawString(200, y, v); y -= 12.5
    y -= 10; c.setFont("Helvetica-Bold", 9)
    c.drawString(50, y, "Pos."); c.drawString(95, y, inv.get("basis") or "Beschreibung (gemäss Ausmass)"); c.drawRightString(400, y, "Menge"); c.drawRightString(475, y, "Preis"); c.drawRightString(w - 50, y, "Betrag")
    c.line(50, y - 5, w - 50, y - 5); y -= 20; c.setFont("Helvetica", 9)
    for i, ln in enumerate(rows, 1):
        if y < 120:
            d.new_page(); y = h - 70; c.setFont("Helvetica", 9)
        parts = _wrap(ln["text"], 50)
        c.drawString(50, y, str(i))
        for k, t in enumerate(parts): c.drawString(95, y - 11 * k, t)
        c.drawRightString(400, y, f'{qty_fmt(ln["qty"], 2)} {ln["unit"]}'); c.drawRightString(475, y, chf(ln["unit_price"])); c.drawRightString(w - 50, y, chf(ln["total"]))
        y -= 11 * len(parts) + 3
    if y < 230:
        d.new_page(); y = h - 70
    y -= 8; c.line(50, y + 6, w - 50, y + 6)
    y = _totals_block(c, 50, w - 50, y - 8, result)
    paid = 0
    c.setFont("Helvetica", 9.5)
    for t, a in inv.get("akonto", []):
        c.drawString(50, y, t); c.drawRightString(w - 50, y, "-" + chf(a)); y -= 12.5; paid += float(a)
    if paid:
        c.setFont("Helvetica-Bold", 10); c.drawString(50, y - 4, "Restbetrag CHF"); c.drawRightString(w - 50, y - 4, chf(float(result["total"]) - paid)); y -= 24
    c.setFont("Helvetica", 9.5)
    if inv.get("skonto_note"):
        for part in _wrap(inv["skonto_note"], 95):
            c.drawString(50, y - 6, part); y -= 12
        y -= 2
    for t in inv.get("notes", []):
        for part in _wrap(t, 95):
            c.drawString(50, y - 6, part); y -= 12
        y -= 2
    c.drawString(50, y - 6, f"Zahlbar innert 30 Tagen netto.   MWST-Nr. {C['uid']} MWST")
    if inv.get("qr", True):
        d.new_page(); _qr_part(c, w, h, inv, result); c.save(); return
    d.save()

def _qr_part(c, w, h, inv, result):
    """Swiss QR bill payment part (layout only; fictional QR-IBAN, pattern not scannable) (DL-160)."""
    rng = random.Random(inv["number"]); y0 = 297 * 2.8346 - 105 * 2.8346
    c.setDash(2, 2); c.line(0, 105 * 2.8346, w, 105 * 2.8346); c.line(62 * 2.8346, 0, 62 * 2.8346, 105 * 2.8346); c.setDash()
    c.setFont("Helvetica-Bold", 11); c.drawString(15, 105 * 2.8346 - 22, "Empfangsschein"); c.drawString(62 * 2.8346 + 15, 105 * 2.8346 - 22, "Zahlteil")
    qx, qy, size = 62 * 2.8346 + 15, 105 * 2.8346 - 200, 130
    n = 33; cell = size / n
    for i in range(n):
        for j in range(n):
            finder = any(fi <= i < fi + 7 and fj <= j < fj + 7 for fi, fj in ((0, 0), (0, n - 7), (n - 7, 0)))
            if (finder and (i % 6 == 0 or j % 6 == 0 or (2 <= i % 26 <= 4 and 2 <= j % 26 <= 4))) or (not finder and rng.random() < 0.5):
                c.rect(qx + j * cell, qy + (n - 1 - i) * cell, cell, cell, fill=1, stroke=0)
    cx_, cy_ = qx + size / 2, qy + size / 2                     # Swiss cross: black square with white cross
    c.setFillColorRGB(1, 1, 1); c.rect(cx_ - 11, cy_ - 11, 22, 22, fill=1, stroke=0); c.setFillColorRGB(0, 0, 0)
    c.rect(cx_ - 9, cy_ - 9, 18, 18, fill=1, stroke=0); c.setFillColorRGB(1, 1, 1)
    c.rect(cx_ - 1.6, cy_ - 6, 3.2, 12, fill=1, stroke=0); c.rect(cx_ - 6, cy_ - 1.6, 12, 3.2, fill=1, stroke=0); c.setFillColorRGB(0, 0, 0)
    tx = qx + size + 20; c.setFont("Helvetica-Bold", 8); c.drawString(tx, 105 * 2.8346 - 40, "Konto / Zahlbar an"); c.setFont("Helvetica", 9)
    for k, t in enumerate(["CH00 3000 0000 0000 0000 0", C["name"], C["street"], f'{C["zip"]} {C["city"]}']):
        c.drawString(tx, 105 * 2.8346 - 52 - k * 11, t)
    c.setFont("Helvetica-Bold", 8); c.drawString(tx, 105 * 2.8346 - 110, "Referenz"); c.setFont("Helvetica", 9)
    c.drawString(tx, 105 * 2.8346 - 121, f'00 00000 00000 {zlib.crc32(inv["number"].encode()) % 100000:05d} 00000 00000')
    c.setFont("Helvetica-Bold", 8); c.drawString(tx, 105 * 2.8346 - 140, "Zahlbar durch"); c.setFont("Helvetica", 9)
    for k, t in enumerate(inv["recipient"]):
        c.drawString(tx, 105 * 2.8346 - 151 - k * 11, t)
    c.setFont("Helvetica-Bold", 8); c.drawString(qx, qy - 20, "Währung"); c.drawString(qx + 60, qy - 20, "Betrag")
    c.setFont("Helvetica", 10); c.drawString(qx, qy - 33, "CHF"); c.drawString(qx + 60, qy - 33, chf(float(result["total"]) - sum(float(a) for _, a in inv.get("akonto", []))))
    c.setFont("Helvetica-Bold", 7); c.drawString(15, 105 * 2.8346 - 40, "Konto / Zahlbar an"); c.setFont("Helvetica", 8)
    for k, t in enumerate(["CH00 3000 0000 0000 0000 0", C["name"], C["street"], f'{C["zip"]} {C["city"]}']):
        c.drawString(15, 105 * 2.8346 - 50 - k * 10, t)
    c.setFont("Helvetica-Bold", 7); c.drawString(15, 105 * 2.8346 - 100, "Referenz"); c.setFont("Helvetica", 8)
    c.drawString(15, 105 * 2.8346 - 110, f'00 00000 00000 {zlib.crc32(inv["number"].encode()) % 100000:05d} 00000 00000')
    c.setFont("Helvetica-Bold", 7); c.drawString(15, 105 * 2.8346 - 126, "Zahlbar durch"); c.setFont("Helvetica", 8)
    for k, t in enumerate(inv["recipient"]): c.drawString(15, 105 * 2.8346 - 136 - k * 10, t)
    c.setFont("Helvetica-Bold", 7); c.drawString(15, 70, "Währung"); c.drawString(60, 70, "Betrag"); c.setFont("Helvetica", 8)
    c.drawString(15, 60, "CHF"); c.drawString(60, 60, chf(float(result["total"]) - sum(float(a) for _, a in inv.get("akonto", []))))
    c.setFont("Helvetica-Bold", 7); c.drawRightString(62 * 2.8346 - 15, 30, "Annahmestelle")

# ---------------------------------------------------------------------------------------------- email
def eml(path, frm, to, subject, date, body, cc=None, attachments=(), quoted=None, in_reply_to=None, mailer=None):
    """frm/to: 'Name <mail>'; date: ISO datetime string; quoted: list of dicts {frm, date, subject, body} rendered below"""
    msg = EmailMessage()
    msg["From"] = frm; msg["To"] = to if isinstance(to, str) else ", ".join(to)
    if cc: msg["Cc"] = cc if isinstance(cc, str) else ", ".join(cc)
    msg["Subject"] = subject
    when = dt.datetime.fromisoformat(date).replace(tzinfo=dt.timezone(dt.timedelta(hours=2)))
    msg["Date"] = format_datetime(when)
    msg["Message-ID"] = make_msgid(domain=frm.split("@")[-1].strip(">"))
    if in_reply_to: msg["In-Reply-To"] = in_reply_to
    if mailer: msg["X-Mailer"] = mailer
    text = body.rstrip() + "\n"
    for q in quoted or []:
        qd = dt.datetime.fromisoformat(q["date"])
        text += (f"\n\nVon: {q['frm']}\nGesendet: {de_date(qd.date())} {qd.strftime('%H:%M')}\n"
                 f"An: {q['to']}\nBetreff: {q['subject']}\n\n" + "\n".join(q["body"].rstrip().splitlines()))
    msg.set_content(text)
    for name, data, mime in attachments:
        maintype, subtype = mime.split("/")
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=name)
    Path(path).write_bytes(bytes(msg))
    return msg["Message-ID"]

# ---------------------------------------------------------------------------------------------- transcript
def vtt(path, cues, start="00:00:05"):
    """cues: list of (speaker, text, seconds_duration)"""
    t = sum(int(x) * f for x, f in zip(start.split(":"), (3600, 60, 1)))
    out = ["WEBVTT", ""]
    for i, (spk, txt, dur) in enumerate(cues, 1):
        a, b = t, t + dur
        fmt = lambda s: f"{s // 3600:02d}:{(s % 3600) // 60:02d}:{s % 60:02d}.{(i * 137) % 1000:03d}"
        out += [f"{i}", f"{fmt(a)} --> {fmt(b)}", f"<v {spk}>{txt}</v>", ""]
        t = b + 1
    Path(path).write_text("\n".join(out), encoding="utf-8")

# ---------------------------------------------------------------------------------------------- file note
def file_note(path, fields, content, agreements, todos):
    doc = new_doc(title_line=True)
    doc.add_heading("Aktennotiz", level=1)
    table(doc, ["", ""], fields, widths=[5, 11])
    doc.add_heading("Inhalt", level=2)
    for t in content: doc.add_paragraph(t)
    doc.add_heading("Vereinbarungen", level=2)
    for t in agreements: doc.add_paragraph(t, style="List Bullet")
    doc.add_heading("Pendenzen", level=2)
    table(doc, ["Was", "Wer", "Bis"], todos, widths=[10, 3, 3])
    doc.save(path)

# ---------------------------------------------------------------------------------------------- scan effect
def scanify(pdf_in, pdf_out, seed=0, dpi=110):
    from PIL import Image, ImageFilter
    rng = random.Random(seed)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["pdftoppm", "-r", str(dpi), "-gray", "-png", str(pdf_in), f"{tmp}/p"], check=True)
        pages = []
        for f in sorted(Path(tmp).glob("p-*.png")):
            im = Image.open(f).convert("L")
            im = im.rotate(rng.uniform(-0.9, 0.9), expand=False, fillcolor=255, resample=Image.BICUBIC)
            px = im.load(); W, H = im.size
            for _ in range(W * H // 900):
                x, y = rng.randrange(W), rng.randrange(H); px[x, y] = rng.randint(120, 200)
            im = im.filter(ImageFilter.GaussianBlur(0.5)).point(lambda v: 255 if v > 235 else int(v * 0.93))
            pages.append(im.convert("RGB"))
        pages[0].save(pdf_out, save_all=True, append_images=pages[1:], resolution=dpi)


# ---------------------------------------------------------------------------------------------- CRM print view
def crm_print(path, cp):
    d = _Doc(path, f'KundenCockpit Druckansicht, {cp["printed"]}, {cp["by"]}', "Seite {n}"); c, w, h = d.c, d.w, d.h
    c.setTitle("KundenCockpit Lead"); y = h - 50
    c.setFillColor(HexColor("#1F4E79")); c.rect(40, y - 8, w - 80, 26, fill=1, stroke=0); c.setFillColor(HexColor("#FFFFFF"))
    c.setFont("Helvetica-Bold", 13); c.drawString(50, y, "KundenCockpit  |  Lead"); c.setFillColor(black); y -= 40
    c.setFont("Helvetica", 9.5)
    for k, v in cp["lead"]:
        c.setFont("Helvetica-Bold", 9.5); c.drawString(50, y, k); c.setFont("Helvetica", 9.5); c.drawString(170, y, v); y -= 14
    y -= 10; c.setFont("Helvetica-Bold", 11); c.drawString(50, y, "Verknüpfte Kontakte"); y -= 16
    for row in cp["contacts"]:
        c.setFont("Helvetica-Bold", 9); c.drawString(50, y, f"{row[0]}  {row[1]}"); y -= 12
        c.setFont("Helvetica", 8.5)
        for t in row[2:]:
            c.drawString(70, y, t); y -= 11
        y -= 6
    y -= 6; c.setFont("Helvetica-Bold", 11); c.drawString(50, y, "Aktivitäten"); y -= 16
    c.setFont("Helvetica-Bold", 8.5)
    for x, t in ((50, "Datum"), (115, "Wer"), (150, "Typ"), (205, "Notiz")): c.drawString(x, y, t)
    c.line(50, y - 3, w - 50, y - 3); y -= 15; c.setFont("Helvetica", 8.5)
    for dte, who, typ, note in sorted(cp["activities"], key=lambda r: r[0][6:] + r[0][3:5] + r[0][:2]):
        c.drawString(50, y, dte); c.drawString(115, y, who); c.drawString(150, y, typ)
        for k, part in enumerate(_wrap(note, 70)):
            c.drawString(205, y - 11 * k, part)
        y -= 11 * max(1, len(_wrap(note, 70))) + 5
    d.save()
