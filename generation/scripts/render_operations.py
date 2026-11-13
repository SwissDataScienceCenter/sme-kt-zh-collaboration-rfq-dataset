"""Render operations documents of the Company world (Swiss German).

company-world/operations/
  Regieansaetze_2018-2026.xlsx
  Preisliste_Systemprodukte_2026.docx/.pdf      current
  Preisliste_Systemprodukte_2024.docx/.pdf       outdated, still circulating (known issue KI-05)
  Kalkulationsrichtlinie_v3.2_2024-03.docx/.pdf
  Faustregeln_BUl.xlsx                          estimator's private rules of thumb
"""
import sys, yaml
from pathlib import Path
from decimal import Decimal
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from docx.shared import Pt, Cm
sys.path.insert(0, str(Path(__file__).parent))
from pricing import World, round_to, D
from docx_helpers import new_doc, table, to_pdf, chf, C, ROOT

W = World(); OPS = W.ops
CAT = yaml.safe_load(open(ROOT / "generation/world/catalogue.yaml", encoding="utf-8"))
OUT = ROOT / "company-world/operations"; OUT.mkdir(parents=True, exist_ok=True)

# ------------------------------------------------------------------ rate tables
def rates():
    wb = Workbook(); wb.remove(wb.active)
    for y in range(2026, 2017, -1):
        ws = wb.create_sheet(str(y))
        ws.append([f"Regieansätze {y}", "", "", f"gültig ab 01.01.{y}"]); ws["A1"].font = Font(bold=True, size=12)
        ws.append([]); ws.append(["Kategorie", "Code", "Ansatz CHF/Std exkl. MWST", "Bemerkung"])
        for c in ws[3]: c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="EEEEEE")
        for k, v in OPS["rate_categories"].items():
            note = "nur Honorarofferten Beratung" if k.startswith("advisory") else ""
            if k.startswith("advisory") and y < 2015: continue
            ws.append([v["de"], v["code"], int(W.rate(k, y)), note])
        ws.append([]); ws.append(["Zuschläge"]); ws.cell(ws.max_row, 1).font = Font(bold=True)
        ws.append(["Nacht- und Wochenendarbeit", "", "50 %", "auf Ansatz"])
        ws.append(["Fahrzeugpauschale pro Tag", "", float(W.price("99.110.01", y)), "Artikel 99.110.01"])          # same source as the offers (A3 review)
        ws.append(["Kleinmaterial Regie", "", "5 %", "auf Lohnbetrag"])
        if y == 2022:
            ws.append([]); ws.append(["Achtung: Ansätze per 01.07.2022 wegen Teuerung nicht angepasst, siehe Mail CSc vom 14.06.2022"])
        for col, w in zip("ABCD", [32, 8, 24, 40]): ws.column_dimensions[col].width = w
    wb.save(OUT / "Regieansaetze_2018-2026.xlsx")

# ------------------------------------------------------------------ price lists
def sp_price(base, year):
    if year == OPS["base_year"]:
        return D(base)
    return round_to(D(base) * W.factor("system_products", year), 10)

STD_DE = {"module_length_m": "Modullänge m", "depth_m": "Tiefe m", "clear_height_m": "lichte Höhe m",
          "bikes_per_module": "Velos pro Modul", "colour": "Farbe", "altitude_max_m": "Standort bis m ü. M.",
          "glass": "Glas", "wall": "Wand"}

def price_list(year):
    doc = new_doc(title_line=True)
    doc.add_heading(f"Preisliste Systemprodukte {year}", level=1)
    doc.add_paragraph(f"Gültig ab 1. Januar {year}. Alle Preise in CHF exkl. MWST, ab Werk Ibach, ohne Montage und Transport.")
    for fam in CAT["families"].values():
        doc.add_heading(fam["name"], level=2)
        doc.add_paragraph(fam["description"])
        std = ", ".join(f"{STD_DE[k]}: {v}" for k, v in fam["standard"].items())
        p = doc.add_paragraph(); r = p.add_run("Standard: "); r.bold = True; p.add_run(std)
        rows = [(m["code"], m["de"], m["unit"], chf(sp_price(m["price"], year))) for m in fam["models"]]
        table(doc, ["Code", "Modell", "Einheit", "Preis CHF"], rows, widths=[3, 9, 2, 3])
        doc.add_paragraph()
        rows = []
        for o in fam["options"]:
            pr = f'{o["pct"]*100:+.0f} % Grundpreis' if "pct" in o else chf(sp_price(o["price"], year))
            rows.append((o["code"], o["de"], o["unit"], pr))
        table(doc, ["Code", "Option", "Einheit", "Preis CHF"], rows, widths=[3, 9, 2, 3])
        cs = fam["colour_surcharge"]
        doc.add_paragraph(f'Farbe: Standard {cs["standard"]}. Andere RAL-Farben +{cs["other_ral"]*100:.0f} %, '
                          f'DB-Farben +{cs["db_colours"]*100:.0f} % auf den Grundpreis.')
        asm = fam["assembly"]
        if "first_module" in asm:
            txt = f'Montage: erstes Modul CHF {chf(sp_price(asm["first_module"], year))}, jedes weitere Modul CHF {chf(sp_price(asm["each_additional"], year))}.'
        elif "flat" in asm:
            txt = f'Montage: Pauschale CHF {chf(sp_price(asm["flat"], year))} pro Vordach.'
        else:
            txt = "Montage: " + ", ".join(f"{k} CHF {chf(sp_price(v, year))}" for k, v in asm.items()) + "."
        doc.add_paragraph(txt)
        p = doc.add_paragraph(); r = p.add_run("Anpassungen: "); r.bold = True
        for rule in fam["adaptation_rules"]:
            doc.add_paragraph(rule, style="List Bullet")
    doc.add_heading("Transport", level=2)
    table(doc, ["Zone", "Distanz ab Ibach", "Pauschale CHF", "Beispiele"],
          [(z["zone"], f'bis {z["max_km"]} km' if z["max_km"] else "über 80 km",
            chf(round_to(D(z["flat"]) * W.factor("other", year), 10)), ", ".join(z["examples"]))
           for z in OPS["transport_zones"]], widths=[1.5, 3.5, 3, 9])
    doc.add_heading("Allgemeines", level=2)
    for t in ["Fundamente und Elektroanschlüsse sind bauseits, sofern nicht als Option bestellt.",
              "Baubewilligungen sind Sache der Bauherrschaft. Wir liefern auf Wunsch Pläne für das Baugesuch.",
              "Lieferfrist ab Auftragsbestätigung und Farbwahl: 6 bis 8 Wochen.",
              "Es gelten unsere Allgemeinen Geschäftsbedingungen."]:
        doc.add_paragraph(t, style="List Bullet")
    f = OUT / f"Preisliste_Systemprodukte_{year}.docx"
    doc.save(f); to_pdf(f)

# ------------------------------------------------------------------ costing manual
def manual():
    doc = new_doc(title_line=True)
    doc.add_heading("Kalkulationsrichtlinie", level=1)
    table(doc, ["Version", "Datum", "Verfasst", "Freigabe"], [("3.2", "März 2024", "B. Ulrich, C. Schuler", "C. Schuler")],
          widths=[2.5, 3, 6, 5])
    H = lambda t: doc.add_heading(t, level=2)
    P = lambda t: doc.add_paragraph(t)
    B = lambda t: doc.add_paragraph(t, style="List Bullet")

    H("1 Zweck und Geltungsbereich")
    P("Diese Richtlinie regelt, wie bei Fronalp Offerten kalkuliert, freigegeben und abgelegt werden. "
      "Sie gilt für alle Geschäftsbereiche: Beratung, Projektgeschäft und Systemprodukte.")
    H("2 Preisgrundlagen")
    for t in ["Artikelstamm MetaBau: Verkaufspreise des laufenden Jahres. Gesperrte Artikel nicht verwenden.",
              "Regieansätze gemäss Tabelle Regieansätze, Blatt des Offertjahres.",
              "Preisliste Systemprodukte des laufenden Jahres.",
              "Lieferanten- und Subunternehmerofferten, nicht älter als 3 Monate.",
              "Vergleichsprojekte auf P:\\Projekte. Preise älterer Projekte nicht ungeprüft übernehmen."]:
        B(t)
    H("3 Offertarten")
    doc.add_heading("3.1 Einheitspreisofferte (Projektgeschäft)", level=3)
    P("Positionen aus dem Artikelstamm übernehmen. Mengen aus Plänen, Aufmass vor Ort oder begründeten Annahmen. "
      "Annahmen pro Position ausweisen, zum Beispiel «Annahme 120 m2». Verschnitt und Reserve auf die Menge schlagen "
      "und auf ganze Einheiten aufrunden; alle Längen und Flächen werden auf ganze m bzw. m2 aufgerundet. "
      "Abgerechnet wird nach Ausmass ohne Verschnitt:")
    wf = OPS["waste_factors"]
    for k, de in [("membranes", "Abdichtungsbahnen"), ("slabs", "Platten"), ("sheet_metal", "Spenglerblech"), ("steel", "Stahl")]:
        B(f"{de}: {wf[k]*100:.0f} %")
    P("Ohne Besichtigung vor Ort ist der Textbaustein H1 zu verwenden und eine Regieposition für Unvorhergesehenes "
      "aufzunehmen.")
    doc.add_heading("3.2 Pauschalofferte Metallbau", level=3)
    P("Kalkulation von unten nach oben im Kalkulationsblatt Metallbau.xlsx:")
    for t in ["Material zum Einstandspreis (EK) plus 15 % Materialzuschlag",
              "Werkstattstunden x Ansatz Werkstatt",
              "Montagestunden x Ansatz A-Monteur bzw. B-Monteur",
              "Oberflächen (Verzinken pro kg, Pulverbeschichten pro m2)",
              "Transportpauschale nach Zone, Kran nach Aufwand",
              "Zwischensumme Kosten plus 8 % Risiko und Gewinn",
              "Pauschalpreis auf ganze CHF 10 aufrunden"]:
        B(t)
    doc.add_heading("3.3 Eingabe Leistungsverzeichnis", level=3)
    P("Einheitspreise aus dem Artikelstamm in das Leistungsverzeichnis des Planers übertragen. Positionen, die im "
      "Leistungsverzeichnis fehlen aber nötig sind, als Eventualposition aufführen und beim Planer schriftlich "
      "nachfragen. Eingabefrist im CRM als Aufgabe erfassen.")
    doc.add_heading("3.4 Konfigurationsofferte Systemprodukte", level=3)
    P("Grundpreis, Optionen und Montagepauschale gemäss Preisliste des laufenden Jahres. Anpassungen nur nach den "
      "Regeln der Preisliste. Sonderfälle (Standort über 1200 m ü. M., Holzfassaden, Hanglage über 30 cm) vor dem "
      "Versand mit dem Werkstattchef klären.")
    doc.add_heading("3.5 Honorarofferte Beratung", level=3)
    P("Aufwand pro Phase und Kategorie schätzen (B1 Geschäftsleitung, B2 Projektleitung, B3 Technik). Übliche Phasen:")
    for t in ["Grundlagen und Bestandesaufnahme", "Analyse und Varianten", "Konzept und Kostenschätzung",
              "Begleitung Ausschreibung und Ausführung (optional)"]:
        B(t)
    P("Pauschalleistungen der Artikelgruppe 98 (Drohne, Thermografie, Sondierung) separat aufführen. Annahmen und "
      "nicht enthaltene Leistungen ausdrücklich nennen. Normen und Vorschriften werden genannt, ihre Auslegung im "
      "Einzelfall ist nicht Teil der Honorarofferte.")
    P("Verrechnung nach Aufwand mit Kostendach; Mehraufwand nur nach schriftlicher Zustimmung der Auftraggeberin. "
      "Mindesthonorar für eine Zustandsaufnahme CHF 2'500. Beratung: 30 Tage netto, kein Skonto. "
      "Die Pauschalen 98.110 (Kurzbericht, Variantenstudie) sind seit 2021 gesperrt.")
    doc.add_heading("3.6 Richtofferte", level=3)
    acc = OPS["cost_estimate_accuracy"]
    P(f"Richtofferten sind unverbindlich. Genauigkeit angeben: ±{acc['with_site_visit']*100:.0f} % mit Besichtigung, "
      f"±{acc['without_site_visit']*100:.0f} % ohne Besichtigung. Hinweis auf den ungefähren Kostenansatz nach "
      "Art. 375 OR aufnehmen.")
    H("4 Konditionen")
    cr = OPS["commercial_rules"]
    table(doc, ["Thema", "Regel"], [
        ("Rabatt", "Standard 0 %. Stammkunden 2 %. Mehr nur mit Freigabe GL."),
        ("Skonto", "2 % bei Zahlung innert 10 Tagen, berechnet auf dem Bruttototal."),
        ("Rundung", "Rabatt, Skonto und MWST auf 5 Rappen runden."),
        ("Gültigkeit", "3 Monate ab Offertdatum."),
        ("Zahlung Privatkunden", cr["payment_terms"]["b2c"]),
        ("Zahlung Firmenkunden Projektgeschäft über CHF 30'000", cr["payment_terms"]["b2b_custom_build_above_30000"]),
        ("Zahlung übrige Firmenkunden", cr["payment_terms"]["b2b_default"]),
        ("Mindermengen", "Netto unter CHF 1'500: Zuschlag CHF 120."),
    ], widths=[5.5, 11])
    H("5 Eventualpositionen und bauseitige Leistungen")
    P("Eventualpositionen werden mit Menge und Preis aufgeführt, der Betrag steht in Klammern und zählt nicht zum Total. "
      "Bauseitige Leistungen stehen als Textblock «Leistungen bauseits» am Anfang der Offerte.")
    H("6 Freigabe")
    P("Offerten über CHF 50'000 netto oder mit Rabatt über 3 % gibt die Geschäftsleitung frei. Die Freigabe kann "
      "telefonisch erfolgen und ist im CRM als Aktivität zu vermerken.")
    H("7 Transportzonen")
    P("Die Transportpauschalen gelten für Systemprodukte und Pauschalofferten Metallbau. Im Projektgeschäft mit "
      "Einheitspreisen ist der Transport in der Baustelleninstallation enthalten.")
    table(doc, ["Zone", "Distanz ab Ibach", "Beispiele"],
          [(z["zone"], f'bis {z["max_km"]} km' if z["max_km"] else "über 80 km", ", ".join(z["examples"]))
           for z in OPS["transport_zones"]], widths=[1.5, 4, 11])
    H("8 Beispiel Totalblock")
    ex = [("Total brutto", "", "24'380.00"), ("Rabatt", "2.00 %", "-487.60"), ("Skonto", "2.00 %", "-487.60"),
          ("Zwischentotal", "", "23'404.80"), ("Mehrwertsteuer", "7.70 %", "1'802.15"), ("Total CHF inkl. MWST", "", "25'206.95")]
    table(doc, ["", "Satz", "Betrag CHF"], ex, widths=[7, 3, 4])   # known issue KI-01: outdated VAT rate in example
    H("9 Nummerierung und Ablage")
    for t in ["Offertnummer aus MetaBau: OF-JJ-####. Projektnummer bei Auftrag: P-JJ-###.",
              "Ablage: P:\\Projekte\\<Jahr>\\<Projektnr> <Objekt>\\ mit den Unterordnern 01_Anfrage, 02_Kalkulation, "
              "03_Offerte, 04_Ausführung, 05_Abrechnung.",
              "Vor dem Auftrag liegen Offerten unter P:\\Offerten\\<Jahr>\\<Offertnr>."]:
        B(t)
    H("Anhang: Textbausteine")
    for code, txt in [
        ("H1 Hinweis ohne Besichtigung", "Die in dieser Offerte aufgeführten Leistungen und Preise beruhen auf Annahmen, da "
         "keine Besichtigung vor Ort stattgefunden hat. Zeigen sich bei der Ausführung abweichende Verhältnisse, behalten "
         "wir uns entsprechende Anpassungen der Leistungen und Kosten vor."),
        ("B1 Leistungen bauseits (Standard)", "Unentgeltlicher Strom- und Wasseranschluss; Zufahrt und Stellplatz für "
         "Firmenfahrzeug, Mulde und Kran; freier Zugang zu den Arbeitsbereichen; Räumen der Arbeitsflächen."),
        ("V1 Vorbehalte (Standard)", "Nicht sichtbare Mängel am Untergrund, zusätzliche Abdichtungsarbeiten an "
         "angrenzenden Bauteilen und Mehrmengen werden nach Aufwand bzw. nach Ausmass verrechnet."),
        ("R1 Richtofferte", "Es handelt sich um eine unverbindliche Richtofferte mit einer Genauigkeit von ±__ %. "
         "Sie gilt als ungefährer Kostenansatz im Sinne von Art. 375 OR.")]:
        p = doc.add_paragraph(); r = p.add_run(code + ": "); r.bold = True; p.add_run(txt)
    f = OUT / "Kalkulationsrichtlinie_v3.2_2024-03.docx"; doc.save(f); to_pdf(f)

# ------------------------------------------------------------------ rules of thumb
def faustregeln():
    wb = Workbook(); ws = wb.active; ws.title = "Faustregeln"
    ws.append(["Thema", "Faustregel", "Bemerkung", "Stand"])
    for c in ws[1]: c.font = Font(bold=True)
    rows = [
        ("Terrasse", "Plattenbelag: Fläche + 10 % Verschnitt", "bei Diagonalverlegung 15 %", "2019"),
        ("Terrasse", "Abdichtung 2-lagig: Fläche + 10 %, Aufbordung separat in m", "", "2017"),
        ("Terrasse", "Ohne Besichtigung: Regie Reserve 12 h A + 12 h B + Material 800.-", "bis 2024: 8 h + 8 h + 500.-", "2025"),
        ("Terrasse", "Notüberlauf: mind. 1 pro Ablauf, sonst Ärger mit Verwaltung", "", "2016"),
        ("Flachdach", "Gefälledämmung im Mittel 60 mm reicht bis ca. 10 m Fliessweg", "darüber 100 mm", "2018"),
        ("Kran", "Ab 3. OG oder Material über 1 t immer Kran", "Kran & Mulden Muotathal anrufen", "2022"),
        ("Kran", "Mindestens 4 Std verrechnen, auch wenn schneller", "", "2020"),
        ("Geländer", "Montage 1.2 h pro m (2 Mann), seitlich befestigt 1.5 h", "", "2015"),
        ("Geländer", "Gewicht Stahlgeländer ca. 18 kg/m inkl. Füllung Flachstahl und Handlauf", "Glas: 28 kg/m", "2015"),
        ("Treppe", "Aussentreppe gerade, Stahl, 1 m breit: Werkstatt 6 h pro Stufe", "gewendelt x 1.5", "2014"),
        ("Treppe", "Steigung aussen max. 18 cm, Schrittmass 2 h + a = 62 bis 64 cm; Steigungen = Höhe / 18 cm aufrunden, Stufen = Steigungen - 1", "Höhe immer ab fertigem Belag oben und unten", "2014"),
        ("Treppe", "Werkstatt 6 h pro Stufe gilt für Wangen, Stufenwinkel, Kopf- und Fussplatten; Geländer separat", "Lauerz 2025: 6.5 h gebraucht", "2016"),
        ("Treppe", "Günstiger: Gitterroststufen mit Seitenblechen verschraubt, keine Stufenwinkel, Werkstatt 3.5 h pro Stufe", "", "2026"),
        ("Treppe", "Auftritt auf ganze cm abrunden, so dass vor dem Antritt mind. 15 cm Platz bleibt; Schrittmass prüfen", "", "2026"),
        ("Treppe", "Offene Stufen: Spalt max. 12 cm. Bei Steigung über 15 cm Zwischenstab Flachstahl 40x8, 1 m pro Stufe, Werkstatt 0.5 h pro Stufe", "Kinder!", "2026"),
        ("Treppe", "Anschluss an bestehende Betonplatte: 4 Injektionsanker M16, Lage nach Statik; Verstärkung oder Stützen separat offerieren", "", "2026"),
        ("Treppe", "Wangen Flachstahl 200x12 bis 5.5 m, Länge = Lauflinie schräg + 0.34 m für Anschlüsse; Statik immer bei Aussentreppen", "Statik Talkessel", "2018"),
        ("Treppe", "Stahlgewicht Treppe: Wangen + 4 kg pro Stufe (Winkel) + 24 kg Kopf- und Fussplatten", "Gitterroststufen separat", "2018"),
        ("Treppe", "Montage Aussentreppe mit Geländer: A + B je 1 Tag bis 16 Stufen mit Mobilkran vormontiert, je 1.5 Tage wenn die Teile einzeln mit dem Ladekran versetzt werden; Fahrzeugpauschale pro angefangenem Montagetag; bei Öffnung in bestehendem Geländer Seitenschutz temporär 2 m", "", "2017"),
        ("Kran", "Teile bis 150 kg und Ausladung bis 12 m: eigener Lastwagen mit Ladekran, mind. 2 Std; schwerer oder weiter: Mobilkran, mind. 4 Std", "Werner fährt", "2018"),
        ("Geländer", "Werkstatt pro m: Flachstahl-Staketen 3.5 h, mit Lochblechfüllung 4 h", "bis 2025: 2.5 / 3.5 h, Lauerz 2025 hat 17 Std für 5 m gebraucht", "2026"),
        ("Geländer", "Treppengeländer: Länge = Lauflinie schräg + 0.5 m (Antritt und Anschluss oben), auf 0.1 m aufrunden", "Handlauf an Wand: Lauflinie + 0.1 m, ebenfalls auf 0.1 m aufrunden", "2019"),
        ("Geländer", "Geländer mit Lochblech: Rahmen 10 kg/m, Handlauf Rohr 42.4, Lochblech 0.9 m2 pro m; Lochblech erst nach dem Verzinken einklemmen", "verzieht sich sonst im Zinkbad", "2026"),
        ("Geländer", "Nach SIA 358 und bfu: keine Öffnung über 12 cm, nicht bekletterbar, Handlauf greifbar; Höhe Balkon 1.00 m", "", "2019"),
        ("Geländer", "Handlauf an Wand Chromstahl: Werkstatt 3 h, Montage A + B je halber Tag, 1 Halter pro angefangener m", "", "2020"),
        ("Geländer", "Bestehendes Geländer öffnen für Treppenanschluss: Werkstatt 4 h, Montage A 2 h", "", "2020"),
        ("Oberfläche", "Pulverfläche: Treppe 0.4 m2 pro Steigung, Staketengeländer 1.2 m2 pro m, Lochblechgeländer 2.0 m2 pro m (beidseitig), Handlauf 0.15 m2 pro m, neuer Pfosten 0.25 m2; zusammen auf ganze m2 aufrunden; unter 20 m2 Mindermengenzuschlag", "Duplex = verzinkt + pulver, nicht zusätzlich verzinken", "2026"),
        ("Oberfläche", "Oberflächen wie Material zum EK + 15 %; Gitterroststufen sind schon verzinkt", "", "2019"),
        ("Allgemein", "Konstruktion Treppe: 6 h Werkstattpläne + 2 h Pläne Baugesuch, zum Werkstattansatz", "RHe", "2021"),
        ("Allgemein", "Privatkunden: Gesamtpaket anbieten (Fundamente über Baumeister, Statik, Pläne Baugesuch, Koordination) = PL 10 h; ohne Gesamtpaket PL 4 h", "Lauerz 2025: Baumeister des Kunden 3 Wochen zu spät", "2025"),
        ("Allgemein", "Fremdleistungen (Baumeister, Ingenieur) zum Offertpreis des Partners in die Kosten, W&G 8 % darauf", "", "2025"),
        ("Allgemein", "Privatkunden über CHF 20'000: 30 % bei Auftrag, 60 % nach Montage, 10 % nach Abnahme, je 30 Tage netto, kein Skonto", "Lauerz 2025: Vorfinanzierung", "2026"),
        ("Treppe", "Verzinken: Gewicht x 1.05 wegen Zinkauflage", "", "2012"),
        ("Spengler", "Steildach: Deckstreifen = Traufe + 1 m pro Ecke", "", "2010"),
        ("Terrasse", "Standardaufbau ab 2024: Dampfsperre, PIR 160, Gefälle EPS 60, 2-lagig, Schutzlage, Drainagematte, Kiesleiste an offenen Rändern", "Förderprogramm verlangt guten U-Wert, vorher PIR 120", "2024"),
        ("Terrasse", "Neuer Aufbau meist 10 bis 15 cm höher: bei Fenstertüren Rinne als Eventualposition, Schwellen vor Ort prüfen", "", "2024"),
        ("Terrasse", "Geländer am Rand: demontieren, zwischenlagern, wieder montieren; Erhöhung als Eventualposition", "Höhe nach Sanierung prüfen", "2023"),
        ("Terrasse", "Seitenschutz temporär entlang aller offenen Ränder, sobald Geländer weg ist", "BauAV, Suva kontrolliert", "2022"),
        ("Terrasse", "Wiederverwendung Platten: ganze Fläche aufnehmen/palettieren und wieder verlegen, dazu 10 % Ersatzplatten liefern", "", "2021"),
        ("Terrasse", "Sanierung mit höherem Aufbau: Sanierungseinlauf mit Aufstockelement statt Ablauf einbinden", "", "2024"),
        ("Terrasse", "Pro Fenstertür ein Türschwellenblech Uginox", "", "2015"),
        ("Spengler", "Demontage immer ganze Länge Dachrand und Fassadenanschluss", "", "2015"),
        ("Terrasse", "Aufbordung Terrasse bis 30 cm, ausser an Brüstungen", "", "2015"),
        ("Terrasse", "Rückbau Dachaufbau: Mulde 7 m3 1 Stk pro 60 m2, aufrunden; Umschlag mit Kranmulde im Kran", "bis 2023 Kranmulden 1 m3", "2024"),
        ("Kran", "Sanierung mit Wiederverwendung Platten: 3 Einsätze (Rückbau, Material, Platten) à mind. 4 Std, total ca. 16 Std", "", "2024"),
        ("Allgemein", "Baustelleninstallation: klein bis 150 m2, mittel bis 400 m2, gross darüber", "Transport ist in der Installation enthalten", "2016"),
        ("Terrasse", "Baustellen-WC: 1 Monat bis 120 m2, darüber 2 Monate", "", "2019"),
        ("Allgemein", "Bewohnte Wohnung darunter: Abdecken und Schützen immer als Pauschale", "Ärger 2022 in Seewen", "2022"),
        ("Abdichtung", "Dichtheitsprüfung offerieren, wenn Verwaltung oder STWEG einen Nachweis will", "", "2023"),
        ("Fassade", "Unterkonstruktion = Bekleidungsfläche + 5 %", "", "2019"),
        ("Systemprod.", "Carport über 800 m: verstärkt nehmen, nicht diskutieren", "Toni fragen wenn über 1200", "2020"),
        ("Systemprod.", "Velounterstand Hang: Stützen max 30 cm verlängern", "", "2018"),
        ("Systemprod.", "Velounterstand mit nur 1 Seitenwand: 2. Seitenwand als Option offerieren", "wird oft nachbestellt", "2021"),
        ("Beratung", "Geländerprüfung: Auszugsversuche an den Ankern, Stichprobe 10 % der Anker pro Gebäude, mind. 3 pro Gebäude (seit 2026 Balkone und Treppenhaus je für sich aufrunden), Auswahl nach der Sichtprüfung nach Risiko; MA 0.5 h pro Versuch inkl. Verschliessen der Prüfstelle; eigenes Prüfgerät, Fahrzeugpauschale pro Prüftag", "Küssnacht 2025: ausgebrochener Anker von aussen nicht sichtbar", "2026"),
        ("Beratung", "Geländerprüfung: SKä als B2 für die Beurteilung der Befestigungen vor Ort und die Massnahmen", "", "2025"),
        ("Allgemein", "Stammkunden 2 %, Genossenschaften auch 2 % wenn sie fragen", "", "2023"),
        ("Allgemein", "Arnold & Gisler Immobilien zahlt immer spät, kein Skonto anbieten", "", "2021"),
        ("Allgemein", "Offerten über 50'000: Claudia anrufen, nicht mailen", "", "2022"),
        ("Allgemein", "MWST 7.7 %", "ab 2024 8.1 % !!", "2024"),
    ]
    for r in rows: ws.append(r)
    for col, w in zip("ABCD", [14, 62, 34, 8]): ws.column_dimensions[col].width = w
    wb.save(OUT / "Faustregeln_BUl.xlsx")

if __name__ == "__main__":
    rates(); price_list(2026); price_list(2024); manual(); faustregeln()
    print("\n".join(sorted(p.name for p in OUT.iterdir())))
