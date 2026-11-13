"""Render templates, general terms, operations manual and partner list (Swiss German).

company-world/operations/templates/
  Vorlage_Offerte_Word_2016.docx          old general Word template (Lump sum, older offers)
  Vorlage_Honorarofferte_Beratung.docx    Advisory fee offer template (2015, updated 2024)
  Vorlage_Aktennotiz.docx                 File note template
company-world/operations/
  AGB_Fronalp_2022-08.docx/.pdf
  Betriebshandbuch_Auftragsabwicklung_v2.1_2023-05.docx/.pdf
  Subunternehmer_und_Partner.xlsx
Known issues: KI-09, KI-10, KI-11 (see generation/world/known_issues.yaml).
"""
import sys
from pathlib import Path
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Font
sys.path.insert(0, str(Path(__file__).parent))
from docx_helpers import new_doc, table, to_pdf, C, ROOT, ORG
from pricing import World

W = World(); OPS = W.ops
OUT = ROOT / "company-world/operations"; TPL = OUT / "templates"; TPL.mkdir(parents=True, exist_ok=True)
GREY = RGBColor(0x80, 0x80, 0x80)

def field(p, text):
    r = p.add_run(text); r.font.color.rgb = GREY; return r

# ------------------------------------------------------------------ old Word template
def word_template():
    doc = new_doc(title_line=True)
    for t in ["«Firma»", "«Anrede» «Vorname» «Name»", "«Strasse»", "«PLZ» «Ort»"]:
        p = doc.add_paragraph(); p.paragraph_format.left_indent = Pt(260); p.paragraph_format.space_after = Pt(0); field(p, t)
    doc.add_paragraph()
    p = doc.add_paragraph("Ibach, "); field(p, "«Datum»")
    p = doc.add_paragraph(); r = p.add_run("OFFERTE "); r.bold = True; r.font.size = Pt(13); f = field(p, "«Offertnummer»"); f.bold = True
    p = doc.add_paragraph("Objekt: "); field(p, "«Objekt», «Adresse Objekt»")
    p = doc.add_paragraph("Sehr geehrte"); field(p, "«r» «Anrede» «Name»")
    doc.add_paragraph("Besten Dank für Ihre Anfrage. Gerne unterbreiten wir Ihnen unsere Offerte wie folgt:")
    for h, body in [
        ("1. Ausgangslage", "«Beschreibung der Ausgangslage»"),
        ("2. Leistungsbeschrieb", "«Beschreibung der Arbeiten, Material, Ausführung»"),
    ]:
        doc.add_heading(h, level=2); field(doc.add_paragraph(), body)
    doc.add_heading("3. Bauseitige Leistungen", level=2)
    for t in ["Unentgeltlicher Strom- und Wasseranschluss", "Zufahrt und Stellplatz für Firmenfahrzeug",
              "Freier Zugang zu den Arbeitsbereichen", "«weitere»"]:
        doc.add_paragraph(t, style="List Bullet")
    doc.add_heading("4. Kosten", level=2)
    table(doc, ["Pos.", "Beschreibung", "Betrag CHF"],
          [("1", "«Position»", "«0.00»"), ("2", "«Position»", "«0.00»"), ("", "Total exkl. MWST", "«0.00»"),
           ("", "MWST 7.7 %", "«0.00»"), ("", "Total inkl. MWST", "«0.00»")], widths=[1.5, 11, 3.5])   # KI-09
    doc.add_heading("5. Optional", level=2)
    field(doc.add_paragraph(), "«Optionale Leistung, Preis pro Einheit exkl. MWST»")
    doc.add_heading("6. Konditionen", level=2)
    table(doc, ["", ""], [("Gültigkeit", "30 Tage"), ("Zahlung", "Wählen Sie ein Element aus."),   # KI-11
                          ("Garantie", "gemäss SIA-Norm 118"), ("AGB", "Es gelten unsere Allgemeinen Geschäftsbedingungen.")],
          widths=[4, 12])
    doc.add_heading("7. Termine", level=2)
    doc.add_paragraph("Ausführung nach Absprache. Dauer der Arbeiten ca. «x» Arbeitstage.")
    doc.add_paragraph("Wir danken Ihnen für Ihr Interesse und freuen uns auf Ihren Auftrag. Für Fragen stehen wir gerne zur Verfügung.")
    doc.add_paragraph("Freundliche Grüsse"); doc.add_paragraph(C["name"])
    doc.add_paragraph("Peter Kollmann\nProjektleiter")   # KI-10: stale default signature
    doc.add_paragraph()
    p = doc.add_paragraph(); r = p.add_run("Auftragsbestätigung"); r.bold = True
    doc.add_paragraph("Ort, Datum: ______________________    Unterschrift: ______________________")
    doc.save(TPL / "Vorlage_Offerte_Word_2016.docx")

# ------------------------------------------------------------------ Advisory template
def fee_template():
    rc = OPS["rate_categories"]
    doc = new_doc(title_line=True)
    for t in ["«Firma»", "«Anrede» «Vorname» «Name»", "«Strasse»", "«PLZ» «Ort»"]:
        p = doc.add_paragraph(); p.paragraph_format.left_indent = Pt(260); p.paragraph_format.space_after = Pt(0); field(p, t)
    p = doc.add_paragraph("Ibach, "); field(p, "«Datum»")
    p = doc.add_paragraph(); r = p.add_run("HONORAROFFERTE "); r.bold = True; r.font.size = Pt(13); field(p, "«OF-JJ-####»")
    p = doc.add_paragraph("Betreff: "); field(p, "«Objekt», «Auftrag»")
    sections = [
        ("1 Ausgangslage und Zielsetzung", ["«Anlass der Anfrage, Zustand, Ziel der Bauherrschaft»"]),
        ("2 Grundlagen", ["Vorhandene Unterlagen: «Pläne, Fotos, frühere Berichte»",
                          "Massgebende Normen und Vorschriften: «nur nennen, z. B. SIA 271, SIA 358, VKF»"]),
    ]
    for h, items in sections:
        doc.add_heading(h, level=2)
        for t in items: field(doc.add_paragraph(style="List Bullet"), t)
    doc.add_heading("3 Leistungen und Aufwand nach Phasen", level=2)
    table(doc, ["Phase", "Leistungen", "B1 Std", "B2 Std", "B3 Std", "Betrag CHF"],
          [("1", "Grundlagen und Bestandesaufnahme", "«»", "«»", "«»", "«»"),
           ("2", "Analyse und Varianten", "«»", "«»", "«»", "«»"),
           ("3", "Konzept und Kostenschätzung", "«»", "«»", "«»", "«»"),
           ("4", "Begleitung Ausschreibung und Ausführung (optional)", "«»", "«»", "«»", "«»")],
          widths=[1.3, 7.2, 1.6, 1.6, 1.6, 2.7])
    doc.add_paragraph(f'Stundenansätze: B1 Geschäftsleitung CHF {rc["advisory_owner"]["rate"]}, '
                      f'B2 Projektleitung CHF {rc["advisory_pm"]["rate"]}, B3 Technik CHF {rc["advisory_technician"]["rate"]} '
                      'pro Stunde exkl. MWST.')
    doc.add_heading("4 Pauschalleistungen", level=2)
    field(doc.add_paragraph(), "«Drohnenaufnahme, Thermografie, Sondierungen gemäss Artikelgruppe 98»")
    doc.add_heading("5 Nicht enthaltene Leistungen", level=2)
    for t in ["Statische Berechnungen und Nachweise durch Ingenieurbüro", "Bewilligungsgebühren",
              "Schadstoffanalysen (Asbest, PCB) durch Fachlabor", "Ausführung der Bauarbeiten", "«weitere»"]:
        doc.add_paragraph(t, style="List Bullet")
    doc.add_heading("6 Annahmen", level=2)
    field(doc.add_paragraph(style="List Bullet"), "«z. B. Zugang zum Dach ohne Gerüst möglich»")
    doc.add_heading("7 Honorar", level=2)
    table(doc, ["", "CHF"], [("Total Aufwand gemäss Phasen 1 bis 3", "«»"), ("Pauschalleistungen", "«»"),
                             ("Spesen (km-Pauschale)", "«»"), ("Total exkl. MWST", "«»"), ("MWST 8.1 %", "«»"),
                             ("Total inkl. MWST", "«»")], widths=[11, 5])
    doc.add_paragraph("Verrechnung nach effektivem Aufwand. Das Total gilt als Kostendach; Mehraufwand nur nach schriftlicher "
                      "Zustimmung der Auftraggeberin. Phase 4 wird nur nach separatem Auftrag ausgeführt.")
    doc.add_heading("8 Termine", level=2); field(doc.add_paragraph(), "«Start, Dauer, Abgabe Bericht»")
    doc.add_heading("9 Konditionen", level=2)
    table(doc, ["", ""], [("Gültigkeit", "3 Monate"), ("Zahlung", "30 Tage netto; ab CHF 10'000 Akonto 50 % bei Auftrag"),
                          ("AGB", "Es gelten unsere Allgemeinen Geschäftsbedingungen.")], widths=[4, 12])
    doc.add_paragraph("Freundliche Grüsse"); doc.add_paragraph(C["name"])
    doc.add_paragraph("Claudia Schuler\nGeschäftsführerin, Bauingenieurin FH")
    p = doc.add_paragraph(); r = p.add_run("Auftrag erteilt"); r.bold = True
    doc.add_paragraph("Ort, Datum: ______________________    Unterschrift: ______________________")
    doc.save(TPL / "Vorlage_Honorarofferte_Beratung.docx")

def file_note_template():
    doc = new_doc(title_line=True)
    doc.add_heading("Aktennotiz", level=1)
    table(doc, ["", ""], [("Datum / Zeit", ""), ("Erstellt von (Kürzel)", ""), ("Art", "Telefon / Baustelle / Besprechung / anderes"),
                          ("Gesprächspartner", ""), ("Objekt / Projekt", ""), ("CRM-Lead / Offerte", "")], widths=[5, 11])
    doc.add_heading("Inhalt", level=2); doc.add_paragraph("\n\n\n")
    doc.add_heading("Vereinbarungen", level=2); doc.add_paragraph("\n\n")
    doc.add_heading("Pendenzen", level=2)
    table(doc, ["Was", "Wer", "Bis"], [("", "", ""), ("", "", "")], widths=[10, 3, 3])
    doc.add_paragraph("Ablage: P:\\Offerten\\<Jahr>\\<Offertnr> bzw. P:\\Projekte\\... \\01_Anfrage. Im CRM als Aktivität verknüpfen.")
    doc.save(TPL / "Vorlage_Aktennotiz.docx")

# ------------------------------------------------------------------ general terms
def agb():
    doc = new_doc(title_line=True)
    doc.add_heading("Allgemeine Geschäftsbedingungen (AGB)", level=1)
    doc.add_paragraph("Ausgabe August 2022")
    clauses = [
        ("Geltungsbereich", "Diese AGB gelten für alle Offerten, Lieferungen und Leistungen der Fronalp Metall & "
         "Gebäudehülle AG (nachfolgend Fronalp). Abweichende Bedingungen des Bestellers gelten nur, wenn Fronalp sie "
         "schriftlich anerkennt. Ergänzend gilt die SIA-Norm 118, soweit diese AGB nichts anderes regeln."),
        ("Offerten", "Offerten sind während der darin genannten Frist verbindlich. Richtofferten sind unverbindlich und "
         "gelten als ungefährer Kostenansatz. Pläne und Unterlagen der Offerte bleiben Eigentum von Fronalp."),
        ("Preise", "Alle Preise verstehen sich in Schweizer Franken exklusive Mehrwertsteuer. Bei Einheitspreisen wird nach "
         "den tatsächlich ausgeführten Mengen (Ausmass) abgerechnet. Pauschalpreise gelten für den beschriebenen "
         "Leistungsumfang."),
        ("Regiearbeiten", "Arbeiten, die nicht in der Offerte enthalten sind, werden nach den jeweils gültigen Regieansätzen "
         "verrechnet. Regierapporte werden dem Besteller oder seiner Bauleitung wöchentlich zur Unterschrift vorgelegt."),
        ("Leistungen des Bestellers", "Der Besteller sorgt für Zufahrt, Stellplätze, Strom und Wasser sowie für freie "
         "Arbeitsbereiche. Verzögerungen, die aus fehlenden bauseitigen Leistungen entstehen, werden nach Aufwand verrechnet."),
        ("Termine", "Termine werden nach Eingang aller für die Ausführung nötigen Angaben verbindlich. Witterung, "
         "Lieferverzögerungen und höhere Gewalt verlängern die Fristen angemessen."),
        ("Zahlung", "Rechnungen sind innert 30 Tagen netto zahlbar, sofern die Offerte nichts anderes vorsieht. Ein in der "
         "Offerte gewährtes Skonto gilt nur bei Zahlung innert 10 Tagen. Bei grösseren Aufträgen kann Fronalp Akontozahlungen "
         "nach Baufortschritt verlangen."),
        ("Eigentumsvorbehalt", "Gelieferte Bauteile bleiben bis zur vollständigen Bezahlung Eigentum von Fronalp, soweit sie "
         "nicht fest mit dem Bauwerk verbunden sind."),
        ("Abnahme", "Nach Fertigstellung findet eine gemeinsame Abnahme statt. Erscheint der Besteller nicht, gilt das Werk "
         "nach 10 Tagen ab Fertigmeldung als abgenommen."),
        ("Mängelrechte", "Die Rügefrist und die Mängelrechte richten sich nach der SIA-Norm 118. Fronalp hat das Recht, "
         "Mängel zuerst selbst zu beheben. Für Verschleissteile und für Schäden durch unsachgemässen Gebrauch oder fehlenden "
         "Unterhalt besteht keine Haftung."),
        ("Systemprodukte", "Velounterstände, Carports und Vordächer werden nach Katalog gefertigt. Die Bauherrschaft ist für "
         "Baubewilligungen, Fundamente und Elektroanschlüsse verantwortlich, sofern nicht anders vereinbart. Für "
         "Standorte über 1200 m ü. M. gelten besondere Abklärungen."),
        ("Beratung", "Beratungsleistungen werden nach Aufwand oder pauschal gemäss Honorarofferte verrechnet. Berichte und "
         "Empfehlungen beruhen auf dem zum Zeitpunkt der Besichtigung sichtbaren Zustand."),
        ("Haftung", "Fronalp haftet für Schäden, die sie vorsätzlich oder grobfahrlässig verursacht. Die Haftung für "
         "Folgeschäden und entgangenen Gewinn ist ausgeschlossen, soweit gesetzlich zulässig."),
        ("Datenschutz", "Fronalp bearbeitet Personendaten nur zur Abwicklung von Anfragen und Aufträgen. Daten werden nicht "
         "an Dritte weitergegeben, ausser an beteiligte Subunternehmer und Lieferanten."),
        ("Recht und Gerichtsstand", "Es gilt schweizerisches Recht. Gerichtsstand ist Schwyz."),
    ]
    for i, (h, t) in enumerate(clauses, 1):
        doc.add_heading(f"{i}. {h}", level=3); doc.add_paragraph(t)
    f = OUT / "AGB_Fronalp_2022-08.docx"; doc.save(f); to_pdf(f)

# ------------------------------------------------------------------ operations manual
def manual():
    doc = new_doc(title_line=True)
    doc.add_heading("Betriebshandbuch Auftragsabwicklung", level=1)
    table(doc, ["Version", "Datum", "Verfasst", "Freigabe"], [("2.1", "Mai 2023", "E. Föhn", "C. Schuler")], widths=[2.5, 3, 5, 5])
    H = lambda t: doc.add_heading(t, level=2)
    P = lambda t: doc.add_paragraph(t)
    B = lambda t: doc.add_paragraph(t, style="List Bullet")
    H("1 Zweck")
    P("Das Handbuch beschreibt den Weg einer Kundenanfrage vom Eingang bis zur Offerte und zum Auftrag. Es ergänzt die "
      "Kalkulationsrichtlinie.")
    H("2 Zuständigkeiten")
    table(doc, ["Bereich", "Zuständig", "Stellvertretung"], [
        ("Eingang und Erfassung", "Erika Föhn", "Leandro Marty"),
        ("Beratung", "Claudia Schuler", "Nadine Suter"),
        ("Dach und Spenglerei", "Reto Betschart", "Hans Imhof"),
        ("Metallbau", "Sandra Kälin", "Urs Ehrler"),
        ("Fassaden", "Peter Kollmann", "Daniel Inderbitzin"),          # KI-10: left June 2023
        ("Systemprodukte", "Daniel Inderbitzin", "Toni Reichmuth"),
        ("Kalkulation", "Bruno Ulrich", "jeweilige Projektleitung"),
        ("Freigabe", "Claudia Schuler", "Markus Schuler"),
    ], widths=[5, 5, 5])
    H("3 Eingangskanäle")
    for t in ["E-Mail an offerten@fronalp-mgh.ch (Hauptkanal).",
              "Persönliche E-Mails an Mitarbeitende: an offerten@ weiterleiten.",
              "Telefon: Aktennotiz nach Vorlage erstellen und an offerten@ senden.",
              "Webformular Systemprodukte: erzeugt automatisch eine E-Mail an offerten@ mit den Formularangaben.",
              "Ausschreibungen von Planern: per E-Mail oder Plattform; Eingabefrist sofort erfassen.",
              "Mündliche Absprachen auf Baustellen oder bei Besichtigungen: Aktennotiz innert 2 Arbeitstagen."]:
        B(t)
    H("4 Ablauf")
    table(doc, ["Schritt", "Was", "Wer", "System", "Frist"], [
        ("1", "Anfrage sichten, Dubletten prüfen", "EFö", "Outlook, KundenCockpit", "1 Arbeitstag"),
        ("2", "Kunde und Lead erfassen, Unterlagen ablegen", "EFö", "KundenCockpit, P:\\Offerten", "1 Arbeitstag"),
        ("3", "Geschäftsbereich zuteilen", "EFö", "KundenCockpit", "1 Arbeitstag"),
        ("4", "Abklären, Besichtigung vereinbaren", "Projektleitung", "Telefon, E-Mail", "5 Arbeitstage"),
        ("5", "Kalkulieren", "Projektleitung, BUl", "MetaBau, Excel", "siehe 5"),
        ("6", "Freigabe (über CHF 50'000 oder Rabatt über 3 %)", "CSc", "Telefon, CRM-Aktivität", "2 Arbeitstage"),
        ("7", "Offerte versenden, Kopie ablegen", "Projektleitung", "MetaBau oder Word, Outlook", "am selben Tag"),
        ("8", "Nachfassen", "Projektleitung", "Telefon, CRM-Aktivität", "nach 14 Tagen"),
        ("9", "Auftrag oder Absage erfassen", "EFö", "MetaBau und KundenCockpit", "bei Eingang"),
    ], widths=[1.5, 5.5, 2.8, 4, 2.7], font=8.5)
    H("5 Bearbeitungsfristen für Offerten")
    table(doc, ["Geschäftsbereich", "Frist ab Erfassung"], [("Systemprodukte", "3 Arbeitstage"),
          ("Projektgeschäft", "10 Arbeitstage, Ausschreibungen gemäss Eingabefrist"),
          ("Beratung", "5 Arbeitstage nach dem Erstgespräch")], widths=[6, 10])
    H("6 Besichtigung")
    P("Im Projektgeschäft wird bei Sanierungen und bei Offerten über CHF 20'000 eine Besichtigung empfohlen. Ist keine "
      "Besichtigung möglich, wird nach Annahmen offeriert und der Textbaustein H1 verwendet.")
    H("7 Offertvorlagen")
    for t in ["Projektgeschäft Einheitspreise, Ausschreibungen, Systemprodukte und Richtofferten: Offertdruck aus MetaBau.",
              "Pauschalofferten Metallbau: Word-Vorlage Vorlage_Offerte_Word_2016.docx.",
              "Beratung: Word-Vorlage Vorlage_Honorarofferte_Beratung.docx."]:
        B(t)
    H("8 Hinweise")
    for t in ["Status in MetaBau und KundenCockpit separat nachführen; die Systeme sind nicht verbunden.",
              "Alte Kundennummern (K####) stehen in MetaBau im Feld «Alt-Nr.».",
              "Vor dem Versand prüfen, ob die Preisliste des laufenden Jahres verwendet wurde."]:
        B(t)
    f = OUT / "Betriebshandbuch_Auftragsabwicklung_v2.1_2023-05.docx"; doc.save(f); to_pdf(f)

# ------------------------------------------------------------------ partners
def partners():
    wb = Workbook(); ws = wb.active; ws.title = "Partner"
    ws.append(["Firma", "Ort", "Gewerk", "Kontakt", "Telefon", "E-Mail", "Bewertung (1-3)", "Zuletzt beauftragt", "Bemerkung"])
    for c in ws[1]: c.font = Font(bold=True)
    rows = [
        ("Gerüstbau Rigiblick GmbH", "Arth", "Gerüste", "Herr Zehnder", "+41 41 000 31 10", "info@rigiblick-geruest.ch", 1, "2026", "zuverlässig, 3 Tage Vorlauf"),
        ("Kran & Mulden Muotathal AG", "Muotathal", "Kran, Mulden, WC, Lifte", "Disposition", "+41 41 000 31 20", "dispo@kmm-muotathal.ch", 1, "2026", "Kran mind. 4 Std"),
        ("Elektro Lauerzersee GmbH", "Lauerz", "Elektro", "Herr Betschart", "+41 41 000 31 30", "elektro@lauerzersee.ch", 2, "2025", "für LED Velo und Vordach"),
        ("Statik Talkessel AG", "Schwyz", "Ingenieur Tragwerk", "Frau Dr. Ulrich", "+41 41 000 31 40", "info@statik-talkessel.ch", 1, "2026", "Statik Treppen, Carports über 1200 m"),
        ("Bauunternehmung Hinteribach AG", "Ibach", "Baumeister", "Herr Lüönd", "+41 41 000 31 50", "bau@hinteribach.ch", 2, "2025", "Fundamente, Kernbohrungen"),
        ("Malerei Seeboden GmbH", "Brunnen", "Maler", "Herr Ott", "+41 41 000 31 60", "seeboden@maler-brunnen.ch", 2, "2024", ""),
        ("Schadstoff-Analytik Zentralschweiz GmbH", "Goldau", "Schadstoffanalysen", "Labor", "+41 41 000 31 70", "labor@sa-zentralschweiz.ch", 1, "2025", "Asbest, PCB, PAK; Resultat in 5 AT; ca. CHF 90 bis 150 pro Probe"),
        ("Glas & Plan Fronbach AG", "Küssnacht", "Glaser", "Herr Annen", "+41 41 000 31 80", "glas@fronbach.ch", 2, "2026", "auch Lieferant VSG"),
        ("Zimmerei Holzegg AG", "Schwyz", "Holzbau", "Herr Gwerder", "+41 41 000 31 90", "holz@holzegg.ch", 1, "2025", "Unterkonstruktion, Holzroste"),
        ("Spenglerei am See GmbH", "Brunnen", "Spengler (Kapazität)", "Herr Inglin", "+41 41 000 32 00", "info@spenglerei-am-see.ch", 3, "2023", "nur bei Engpass, Qualität schwankend"),
        ("Dachgarten Rigi GmbH", "Goldau", "Dachbegrünung", "Frau Kälin", "+41 41 000 32 10", "info@dachgarten-rigi.ch", 2, "2025", ""),
        ("Strohdach Werkstatt Rotbach GmbH", "Rotbach AG", "Strohdach", "Herr Wüthrich", "+41 62 000 32 20", "", "", "nie", "Kontakt von Markus Schuler, nie beauftragt"),
    ]
    for r in rows: ws.append(r)
    for col, w in zip("ABCDEFGHI", [38, 12, 24, 16, 18, 30, 14, 16, 40]): ws.column_dimensions[col].width = w
    wb.save(OUT / "Subunternehmer_und_Partner.xlsx")

if __name__ == "__main__":
    word_template(); fee_template(); file_note_template(); agb(); manual(); partners()
    for p in sorted(list(OUT.glob("*")) + list(TPL.glob("*"))): print(p.relative_to(ROOT))
