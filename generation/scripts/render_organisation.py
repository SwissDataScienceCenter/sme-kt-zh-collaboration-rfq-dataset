"""Render the organisation documents of the Company world from generation/world/organisation.yaml.

Outputs (Swiss German) in company-world/organisation/:
  Firmenportrait_Fronalp_2026.docx and .pdf
  Organigramm_Fronalp_2026-10.pdf
  MetaBau_Export_Personal_2026-04-30.xlsx
"""
import subprocess, sys, unicodedata
from pathlib import Path
import yaml
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.styles import Font
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor

ROOT = Path(__file__).resolve().parents[2]
W = yaml.safe_load(open(ROOT / "generation/world/organisation.yaml", encoding="utf-8"))
OUT = ROOT / "company-world/organisation"
OUT.mkdir(parents=True, exist_ok=True)
C = W["company"]
STAFF = {s["ini"]: s for s in W["staff"]}
RED = RGBColor(0x8C, 0x1D, 0x18)

def to_pdf(docx_path):
    """Convert with LibreOffice in a scratch dir outside the dataset, then copy the PDF next to the source."""
    import shutil, tempfile
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / docx_path.name
        shutil.copy(docx_path, src)
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(src)],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        shutil.copy(Path(tmp) / (docx_path.stem + ".pdf"), docx_path.with_suffix(".pdf"))

def disp(ini):
    s = STAFF[ini]
    return s.get("display_ini", s["ini"])

def full(ini):
    s = STAFF[ini]
    return f'{s["first"]} {s["last"]}'

def ascii_mail(first, last):
    t = f"{first}.{last}".lower().replace(" ", "")
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ã", "a"), ("é", "e")):
        t = t.replace(a, b)
    t = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode()
    return f"{t}@{C['email_general'].split('@')[1]}"

# ---------------------------------------------------------------- Firmenportrait
def portrait():
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(10.5)
    for sec in doc.sections:
        sec.left_margin = sec.right_margin = Cm(2.2)
    hdr = doc.sections[0].header.paragraphs[0]
    hdr.text = f'{C["name"]}  |  {C["street"]}, {C["zip"]} {C["city"]}  |  {C["web"]}'
    hdr.runs[0].font.size = Pt(8)
    p = doc.add_paragraph(); r = p.add_run("FRONALP"); r.bold = True; r.font.size = Pt(26); r.font.color.rgb = RED
    p = doc.add_paragraph(); r = p.add_run("Metall & Gebäudehülle"); r.font.size = Pt(13)
    p = doc.add_paragraph(); r = p.add_run(C["slogan"]); r.italic = True
    doc.add_heading("Firmenportrait 2026", level=1)

    doc.add_heading("Über uns", level=2)
    doc.add_paragraph(
        "Die Fronalp Metall & Gebäudehülle AG in Ibach ist ein Familienunternehmen in dritter Generation. "
        "Rund 35 Mitarbeitende planen, fertigen und montieren Metallbau, Dach- und Spenglerarbeiten sowie eigene "
        "Systemprodukte. Unsere Kundschaft reicht von Privatpersonen über Liegenschaftsverwaltungen und "
        "Wohnbaugenossenschaften bis zu Architekturbüros, Generalunternehmungen und öffentlichen Bauherrschaften "
        "in der Zentralschweiz und im Raum Zürich.")
    doc.add_paragraph(
        "Werkstatt, Montage und Planung sind unter einem Dach. Das verkürzt die Wege: Wer Ihr Projekt offeriert, "
        "kennt auch die Leute, die es bauen.")

    doc.add_heading("Unsere Geschäftsbereiche", level=2)
    areas = [
        ("Beratung", "Zustandsaufnahmen von Dächern, Fassaden und Geländern, Sanierungskonzepte und Machbarkeitsstudien. "
         "Wir verrechnen nach Aufwand oder pauschal gemäss Honorarofferte."),
        ("Projektgeschäft", "Flachdach- und Terrassensanierungen, Spenglerarbeiten, Treppen, Geländer und Fassadenbekleidungen. "
         "Jedes Projekt wird individuell geplant und nach Einheitspreisen oder pauschal offeriert. "
         "Ausschreibungen von Architekturbüros bearbeiten wir gerne."),
        ("Systemprodukte", "Velounterstände (Fronalp Velo), Carports (Fronalp Carport) und Vordächer (Fronalp Vordach) aus eigener "
         "Entwicklung. Standardmasse ab Katalog, Anpassungen in Breite, Farbe und Befestigung sind möglich."),
    ]
    for t, d in areas:
        p = doc.add_paragraph(style="List Bullet"); r = p.add_run(t + ": "); r.bold = True; p.add_run(d)

    doc.add_heading("Geschichte", level=2)
    de_hist = {
        1962: "Alois Schuler eröffnet die Schlosserei A. Schuler in Ibach: Tore, Geländer und Reparaturen für die Landwirtschaft.",
        1979: "Umzug an die Werkhofstrasse, erste Schweisshalle.",
        1987: "Markus Schuler übernimmt die Spenglerei Betschart in Seewen. Dach- und Spenglerarbeiten kommen dazu.",
        1988: "Gründung der Fronalp Metall & Gebäudehülle AG, benannt nach dem Fronalpstock, den man von der Werkstatt aus sieht.",
        1998: "Flachdach- und Terrassenabdichtungen ergänzen das Angebot. Die ersten Projektleiter werden angestellt.",
        2005: "Erstes eigenes Systemprodukt: der Velounterstand Fronalp Velo.",
        2009: "Neue Montagehalle und Lancierung von Fronalp Carport.",
        2014: "Lancierung von Fronalp Vordach.",
        2015: "Claudia Schuler, Bauingenieurin FH, gründet den Bereich Beratung.",
        2019: "Claudia Schuler übernimmt die Geschäftsführung, Markus Schuler bleibt Präsident des Verwaltungsrats.",
        2021: "Einführung des ERP MetaBau als Ersatz für die Kalkulationssoftware KALKO 4.",
        2022: "Einführung des CRM KundenCockpit.",
        2024: "Überarbeitetes Offertlayout in MetaBau.",
    }
    for h in W["history"]:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.6); p.paragraph_format.first_line_indent = Cm(-1.6)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(1.6))
        r = p.add_run(f'{h["year"]}\t'); r.bold = True
        p.add_run(de_hist[h["year"]])

    doc.add_heading("Zahlen", level=2)
    for k, v in [("Gründung", "1962"), ("Aktiengesellschaft seit", "1988"), ("Mitarbeitende", "rund 35, davon 4 Lernende"),
                 ("Werkstatt- und Montagefläche", "2'400 m²"), ("Einsatzgebiet", "Kantone Schwyz, Zug, Luzern, Uri und Zürich")]:
        p = doc.add_paragraph(style="List Bullet"); r = p.add_run(k + ": "); r.bold = True; p.add_run(v)

    doc.add_heading("Ansprechpersonen", level=2)
    tbl = doc.add_table(rows=1, cols=3); tbl.style = "Light List"
    for c, t in zip(tbl.rows[0].cells, ["Name", "Funktion", "E-Mail"]):
        c.text = t
    for ini in ["CSc", "RBe", "SKae", "DIn", "BUl", "EFoe", "TRe"]:
        s = STAFF[ini]; row = tbl.add_row().cells
        row[0].text = full(ini); row[1].text = s["role"]; row[2].text = ascii_mail(s["first"], s["last"])

    doc.add_heading("Kontakt", level=2)
    for line in [C["name"], C["street"], f'{C["zip"]} {C["city"]} {C["canton"]}',
                 f'Telefon {C["phone_main"]}', f'Pikett Dach und Spenglerei {C["phone_picket"]}',
                 f'Offertanfragen: {C["email_offers"]}', f'Allgemein: {C["email_general"]}', C["web"],
                 f'UID {C["uid"]} {C["vat_suffix"]}']:
        doc.add_paragraph(line).paragraph_format.space_after = Pt(0)
    f = OUT / "Firmenportrait_Fronalp_2026.docx"
    doc.save(f)
    to_pdf(f)

# ---------------------------------------------------------------- Organigramm
def orgchart():
    f = OUT / "Organigramm_Fronalp_2026-10.pdf"
    pw, ph = landscape(A4)
    c = canvas.Canvas(str(f), pagesize=(pw, ph))
    c.setTitle("Organigramm Fronalp Metall & Gebäudehülle AG")
    def box(x, y, w, h, title, lines, fill="#F3F1EE", stroke="#8C1D18"):
        c.setFillColor(HexColor(fill)); c.setStrokeColor(HexColor(stroke)); c.setLineWidth(0.8)
        c.roundRect(x, y, w, h, 4, fill=1, stroke=1)
        c.setFillColor(HexColor("#222222")); c.setFont("Helvetica-Bold", 8.5)
        c.drawCentredString(x + w / 2, y + h - 12, title)
        c.setFont("Helvetica", 7.2)
        for i, l in enumerate(lines):
            c.drawCentredString(x + w / 2, y + h - 24 - i * 9.5, l)
        return (x + w / 2, y, y + h)
    def line(x1, y1, x2, y2):
        c.setStrokeColor(HexColor("#555555")); c.setLineWidth(0.6); c.line(x1, y1, x2, y2)

    c.setFont("Helvetica-Bold", 14); c.setFillColor(HexColor("#8C1D18"))
    c.drawString(30, ph - 36, "FRONALP  Metall & Gebäudehülle AG")
    c.setFont("Helvetica", 9); c.setFillColor(HexColor("#222222"))
    c.drawString(30, ph - 50, "Organigramm, Stand Oktober 2026")

    vr = box(pw / 2 - 90, ph - 110, 180, 40, "Verwaltungsrat", [f"{full('MSc')} (Präsident)"])
    gl = box(pw / 2 - 90, ph - 175, 180, 40, "Geschäftsleitung", [f"{full('CSc')} ({disp('CSc')})"])
    line(vr[0], vr[1], gl[0], gl[2])
    adm = box(pw - 230, ph - 175, 170, 52, "Administration",
              [f"{full('EFoe')} ({disp('EFoe')}), Leitung", f"{full('LMa')} ({disp('LMa')})", "1 Lernende Kauffrau"])
    line(gl[0] + 90, gl[1] + 20, pw - 230, gl[1] + 20)  # GL to Administration

    units = [
        ("Beratung", [f"Leitung: {full('CSc')}", f"{full('NSu')} ({disp('NSu')})", f"Unterstützung: {full('SKae')}"]),
        ("Projektgeschäft", [f"Dach/Spenglerei: {full('RBe')}", f"Metallbau: {full('SKae')}", f"Fassaden: {full('DIn')}",
                             f"Kalkulation: {full('BUl')}"]),
        ("Systemprodukte", [f"Leitung: {full('DIn')}", "Velo, Carport, Vordach", f"Technik: {full('TRe')}"]),
        ("Werkstatt", [f"Werkstattchef: {full('TRe')}", "6 Metallbauer/Schweisser", "1 Magaziner/Chauffeur", "2 Lernende"]),
        ("Montage", [f"Vorarbeiter: {full('HIm')}, {full('UEh')}", "8 A-Monteure", "6 B-Monteure", "1 Lernender"]),
    ]
    n = len(units); w = 148; gap = (pw - 60 - n * w) / (n - 1); ytop = ph - 320
    busy = ytop + 72 + 18
    line(gl[0], gl[1], gl[0], busy)
    line(30 + w / 2, busy, 30 + (n - 1) * (w + gap) + w / 2, busy)
    for i, (t, ls) in enumerate(units):
        x = 30 + i * (w + gap)
        b = box(x, ytop, w, 72, t, ls)
        line(b[0], busy, b[0], b[2])
    c.setFont("Helvetica-Oblique", 7)
    c.drawString(30, 40, "Montageequipen werden projektbezogen den Projektleitenden zugeteilt. "
                         "Kürzel gemäss Personalstamm MetaBau.")
    c.drawString(30, 30, f'{C["name"]}, {C["street"]}, {C["zip"]} {C["city"]}')
    c.showPage(); c.save()

# ---------------------------------------------------------------- Personal export
def staff_export():
    kst = {"Verwaltungsrat": "000", "Geschäftsleitung": "100", "Beratung": "200", "Projektgeschäft": "300",
           "Systemprodukte": "400", "Werkstatt": "500", "Montage": "600", "Administration": "900"}
    kat = {"advisory_owner": "B1", "advisory_technician": "B3", "project_manager": "PL", "workshop": "WS",
           "fitter_a": "MA", "fitter_b": "MB", "apprentice": "LE", "none": ""}
    wb = Workbook(); ws = wb.active; ws.title = "PERS_EXPORT"
    cols = ["PERS_NR", "KZ", "VORNAME", "NAME", "FUNKTION", "KST", "EINTRITT", "AUSTRITT", "BG_PROZ",
            "ANSATZ_KAT", "EMAIL", "TEL_INT", "AKTIV"]
    ws.append(cols)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for s in sorted(W["staff"], key=lambda x: x["id"]):
        if s.get("board_only"):
            continue
        ws.append([s["id"], s.get("display_ini", s["ini"]).upper(), s["first"], s["last"], s["role"], kst[s["unit"]],
                   f'{s["since"]}-01-01' if s["since"] < 2000 else (f'01.08.{s["since"]}' if s["rate_category"] == "apprentice" or s["role"].startswith("Lernende") else f'01.0{(s["id"] % 9) + 1}.{s["since"]}'),
                   s.get("until", ""), s["fte"], kat[s["rate_category"]], ascii_mail(s["first"], s["last"]),
                   str(s["id"])[-3:], "N" if s.get("until") else "J"])
    for col, wdt in zip("ABCDEFGHIJKLM", [9, 6, 11, 14, 42, 5, 11, 11, 8, 11, 34, 8, 6]):
        ws.column_dimensions[col].width = wdt
    wb.save(OUT / "MetaBau_Export_Personal_2026-04-30.xlsx")

if __name__ == "__main__":
    portrait(); orgchart(); staff_export()
    print("\n".join(sorted(p.name for p in OUT.iterdir())))
